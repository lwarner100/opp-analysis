import httplib2
import os
from googleapiclient.http import MediaFileUpload
from apiclient import discovery
from oauth2client import client, tools
from oauth2client.file import Storage
from argparse import ArgumentParser
from titlecase import titlecase
import pandas as pd
import numpy as np
import auth
import opp_analysis as oa
import json

# NECESSARY AUTHENTICATION WITH GOOGLE'S APIS (USES AUTH.PY)
SCOPES = ['https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/presentations']
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Opportunity Analysis'
authInst = auth.auth(SCOPES, CLIENT_SECRET_FILE, APPLICATION_NAME)
credentials = authInst.getCredentials()
http = credentials.authorize(httplib2.Http())
drive_service = discovery.build('drive', 'v3', http=http)
slides_service = discovery.build('slides', 'v1', http=http)
sheets_service = discovery.build('sheets', 'v4', http=http)
global_dict = {}  # STORES INFO USED BY VARIOUS PARTS OF THIS SCRIPT
template = 'templateid'
analysis_sheet = 'sheetid'
# ASSIGNS A RANDOM ID TO SHEET NAME TO ENSURE THE SERVICE ACCOUNT TARGETS THE RIGHT SHEET
sheet_rand_id = np.random.randint(0, 99999)


# df = None
org_name = "ORG"
recipient = "EMAIL@SITE.COM"


def parse_args():  # PARSES ARGUMENTS PASSED WHEN RUNNING MAIN.PY
    global df, org_name, recipient
    parser = ArgumentParser(description="Parser")
    parser.add_argument('file_path', action='store',
                        help='The file path of the csv')
    parser.add_argument('org_name', action='store',
                        help='Name of the organization')
    parser.add_argument('recipient', action='store',
                        help='The email of the person to share with')
    args = parser.parse_args()
    file_path = str(args.file_path)
    org_name = titlecase(str(args.org_name))
    recipient = args.recipient
    df = pd.read_csv(file_path)


def listfiles(size):  # LISTS FILES IN A DRIVE. ONLY FOR TESTING PURPOSES
    results = drive_service.files().list(
        pageSize=size, fields="nextPageToken, files(id, name)").execute()
    items = results.get('files', [])
    if not items:
        print('No files found.')
    else:
        print('Files:')
        for item in items:
            print('{0} ({1})'.format(item['name'], item['id']))


def get_id(name, search_limit=100):  # GETS FILE ID OF A FILE FROM ITS EXACT NAME VIA BRUTE FORCE, WORKS AS LONG AS IT IS ONE OF THE 'SEARCH_LIMIT' (DEFAULT 100) MOST RECENT FILES IN THE DRIVE
    presID = None
    files_list = drive_service.files().list(
        pageSize=search_limit, fields="files(id, name)").execute()
    for file in files_list['files']:
        if file['name'] == name:
            presID = file['id']
            break
    return presID


def create_folder(name):  # CREATES A NEW FOLDER FOR ALL FILES TO BE PLACED INTO
    file_metadata = {
        'name': str(name),
        'mimeType': 'application/vnd.google-apps.folder'
    }
    file = drive_service.files().create(body=file_metadata, fields='id').execute()
    return file.get('id')


# UPLOADS 1 IMAGE TO A SPECIFIED FOLDER
def upload_image_to_folder(image_to_upload, folder_id, image_name='image.jpg'):
    global global_dict
    file_metadata = {
        'name': image_name,
        'parents': [folder_id],
        'mimeType': 'image/jpeg'
    }
    media = MediaFileUpload(image_to_upload,
                            mimetype='image/jpeg',
                            resumable=True,
                            chunksize=-1)
    file = drive_service.files().create(
        body=file_metadata, media_body=media, fields='id').execute()
    print(f'UPLOADED {image_name}')
    global_dict[image_name] = file.get('id')
    return file.get('id')


def get_info(presID):  # GETS JSON REPRESENTATION OF SPECIFIED GOOGLE SLIDES
    presentation = slides_service.presentations()
    info = presentation.get(presentationId=presID).execute()
    return info


def get_sheet_info(sheetID):  # GETS JSON REPRESENTATION OF SPECIFIED GOOGLE SHEET
    sheet = sheets_service.spreadsheets()
    info = sheet.get(spreadsheetId=sheetID).execute()
    return info


# SENDS UPDATE REQUEST TO SLIDES API TO ALTER TEMPLATE DUPLICATE
def batch_update_pres(presID):
    print('APPLYING CHANGES TO TEMPLATE')
    if has_ages:
        age_reqs = [{"createImage": {  # AGE DISTRIBUTION
            "elementProperties": {
                'pageObjectId': 'g863e3227ca_0_0',
                'size': {
                    'width': {'magnitude': 350, 'unit': 'PT'},
                    'height': {'magnitude': 220, 'unit': 'PT'}},
                "transform": {
                    "scaleX": 1.376,
                    "scaleY": 1.376,
                    "translateX": 1511750.015,
                    "translateY": 994512.5075,
                    "unit": "EMU"}},
            "url": str(get_file_url(global_dict['age_dist.jpg']))}},
            {"createImage": {  # COMPARISON OF AGE DISTRIBUTION TO MARKETSMART DATA
                "elementProperties": {
                    'pageObjectId': 'g8634961b8a_0_96',
                    'size': {
                        'width': {'magnitude': 350, 'unit': 'PT'},
                        'height': {'magnitude': 220, 'unit': 'PT'}},
                    "transform": {
                        "scaleX": 1.376,
                        "scaleY": 1.376,
                        "translateX": 1511750.015,
                        "translateY": 994512.5075,
                        "unit": "EMU"}},
                "url": str(get_file_url(global_dict['age_dist_compare.jpg']))}},
            {"replaceAllText": {"replaceText": str(round(rfm_obj.age_proportion*100, 1))+'%',  # UPDATES TEXT TO CORRECT PROPORTION OF FILE THAT HAS AGES
                                "pageObjectIds": ['g863e3227ca_0_0', 'g8634961b8a_0_96', 'g8634961b8a_0_186', 'g8916f532eb_0_8'],
                                "containsText": {"text": "{age-proportion}"}}},
            {"replaceAllText": {"replaceText": str(round(rfm_obj.over50_age_proportion*100, 1))+'%',  # UPDATES TEXT TO CORRECT PROPORTION OF FILE THAT IS OVER 50
                                "pageObjectIds": ['g863e3227ca_0_0'],
                                "containsText": {"text": "{age-proportion-under50}"}}}]
        # CHECKS IF THE DATA PROVIDED HAS A COLUMN INDICATING LEGACY MEMBERSHIP-- IF YES IT CREATES CHARTS BASED ON LEGACY INFORMATION
        if 'legacy_member' in rfm_obj.rfm_data.columns:
            age_reqs +=\
                [{"createImage": {  # CREATES AGE DISTRIBUTION CHART OF ONLY LEGACY SOCIETY MEMBERS
                    "elementProperties": {
                        'pageObjectId': 'g8634961b8a_0_186',
                        'size': {
                            'width': {'magnitude': 350, 'unit': 'PT'},
                            'height': {'magnitude': 220, 'unit': 'PT'}},
                        "transform": {
                            "scaleX": 1.376,
                            "scaleY": 1.376,
                            "translateX": 1511750.015,
                            "translateY": 994512.5075,
                            "unit": "EMU"}},
                    "url": str(get_file_url(global_dict['age_dist_legacy.jpg']))}},
                    {"createImage": {  # CREATES AGE DISTRIBUTION CHART COMPARING THEIR DATA TO MARKETSMART'S OF ONLY LEGACY SOCIETY MEMBERS
                        "elementProperties": {
                            'pageObjectId': 'g8916f532eb_0_8',
                            'size': {
                                'width': {'magnitude': 350, 'unit': 'PT'},
                                'height': {'magnitude': 220, 'unit': 'PT'}},
                            "transform": {
                                "scaleX": 1.376,
                                "scaleY": 1.376,
                                "translateX": 1511750.015,
                                "translateY": 994512.5075,
                                "unit": "EMU"}},
                        "url": str(get_file_url(global_dict['age_dist_compare_legacy.jpg']))}}]
    else:
        age_reqs = []

    general_reqs = [
        {"createImage": {  # CREATES CHART SHOWING A TYPICAL SURVEY RESPONSE BREAKDOWN GIVEN THE SIZE OF THEIR AUDIENCE FOR MAJOR GIFTS
            "elementProperties": {
                'pageObjectId': 'g889f3efaf4_0_70',
                'size': {
                    'width': {'magnitude': 350, 'unit': 'PT'},
                    'height': {'magnitude': 220, 'unit': 'PT'}},
                "transform": {
                    "scaleX": 0.92,
                    "scaleY": 0.92,
                    "translateX": 180,
                    "translateY": 175,
                    "unit": "PT"}},
            "url": str(get_file_url(global_dict['response_breakdown_major.jpg']))}},
        {"createImage": {  # CREATES CHART SHOWING A TYPICAL SURVEY RESPONSE BREAKDOWN GIVEN THE SIZE OF THEIR AUDIENCE FOR LEGACY GIFTS
            "elementProperties": {
                'pageObjectId': 'g889f3efaf4_0_110',
                'size': {
                    'width': {'magnitude': 350, 'unit': 'PT'},
                    'height': {'magnitude': 220, 'unit': 'PT'}},
                "transform": {
                    "scaleX": 0.92,
                    "scaleY": 0.92,
                    "translateX": 180,
                    "translateY": 175,
                    "unit": "PT"}},
            "url": str(get_file_url(global_dict['response_breakdown_legacy.jpg']))}},
        {"createImage": {  # CREATES CHART SHOWING A TYPICAL SURVEY RESPONSE BREAKDOWN GIVEN THE SIZE OF THEIR AUDIENCE (INCLUDING MEDIUM RFM PEOPLE) FOR MAJOR GIFTS
            "elementProperties": {
                'pageObjectId': 'g8a4fd4bf87_1_2',
                'size': {
                    'width': {'magnitude': 350, 'unit': 'PT'},
                    'height': {'magnitude': 220, 'unit': 'PT'}},
                "transform": {
                    "scaleX": 0.92,
                    "scaleY": 0.92,
                    "translateX": 180,
                    "translateY": 175,
                    "unit": "PT"}},
            "url": str(get_file_url(global_dict['response_breakdown_major_medium.jpg']))}},
        {"createImage": {  # CREATES CHART SHOWING A TYPICAL SURVEY RESPONSE BREAKDOWN GIVEN THE SIZE OF THEIR AUDIENCE (INCLUDING MEDIUM RFM PEOPLE) FOR LEGACY GIFTS
            "elementProperties": {
                'pageObjectId': 'g8a4fd4bf87_1_46',
                'size': {
                    'width': {'magnitude': 350, 'unit': 'PT'},
                    'height': {'magnitude': 220, 'unit': 'PT'}},
                "transform": {
                    "scaleX": 0.92,
                    "scaleY": 0.92,
                    "translateX": 180,
                    "translateY": 175,
                    "unit": "PT"}},
            "url": str(get_file_url(global_dict['response_breakdown_legacy_medium.jpg']))}},
        {"createSheetsChart": {  # CREATES ADJUSTABLE PIE CHART BASED ON JEFF'S RECOMMENDATIONS OF HOW MANY PEOPLE TO DIRECT MAIL
            "elementProperties": {
                "pageObjectId": 'g8a4fd4bf87_3_25',
                "size": {
                    'width': {'magnitude': 340, 'unit': 'PT'},
                    'height': {'magnitude': 250, 'unit': 'PT'}},
                "transform": {
                    "scaleX": 1,
                    "scaleY": 1,
                    "translateX": 340,
                    "translateY": 90,
                    "unit": "PT"}},
            "spreadsheetId": global_dict['working sheetID'],
            "chartId": 1858722116,
            "linkingMode": 'LINKED'}},
        {"createSheetsChart": {  # TIMELINE CHART OF WHEN LEGACY GIFTS WILL BE REALIZED BASED ON MORTALITY RATES. FROM TOM'S EXCEL FORMULA
            "elementProperties": {
                "pageObjectId": 'g889f3efaf4_0_193',
                "size": {
                    'width': {'magnitude': 470, 'unit': 'PT'},
                    'height': {'magnitude': 216, 'unit': 'PT'}},
                "transform": {
                    "scaleX": 1,
                    "scaleY": 1,
                    "translateX": 177,
                    "translateY": 139,
                    "unit": "PT"}},
            "spreadsheetId": global_dict['working sheetID'],
            "chartId": 23131393,
            "linkingMode": 'LINKED'}},
        {"createSheetsChart": {  # TABLE SHOWING RELEVANT STATS ON EXPECTED RESPONSE AND PROJECTIONS
            "elementProperties": {
                "pageObjectId": 'g86bb6be987_0_131',
                "size": {
                    'width': {'magnitude': 648, 'unit': 'PT'},
                    'height': {'magnitude': 245, 'unit': 'PT'}},
                "transform": {
                    "scaleX": 1,
                    "scaleY": 1,
                    "translateX": 35,
                    "translateY": 56,
                    "unit": "PT"}},
            "spreadsheetId": global_dict['working sheetID'],
            "chartId": 2024033113,
            "linkingMode": 'LINKED'}},
        {"createSheetsChart": {  # TABLE SHOWING RELEVANT STATS ON EXPECTED RESPONSE AND PROJECTIONS
            "elementProperties": {
                "pageObjectId": 'g86bb6be987_0_216',
                "size": {
                    'width': {'magnitude': 648, 'unit': 'PT'},
                    'height': {'magnitude': 245, 'unit': 'PT'}},
                "transform": {
                    "scaleX": 1,
                    "scaleY": 1,
                    "translateX": 35,
                    "translateY": 56,
                    "unit": "PT"}},
            "spreadsheetId": global_dict['working sheetID'],
            "chartId": 43298779,
            "linkingMode": 'LINKED'}},
        {"createSheetsChart": {  # CHART SHOWING RELEVANT STATS ON EXPECTED RESPONSE AND PROJECTIONS
            "elementProperties": {
                "pageObjectId": 'g86bb6be987_0_102',
                "size": {
                    'width': {'magnitude': 454, 'unit': 'PT'},
                    'height': {'magnitude': 280, 'unit': 'PT'}},
                "transform": {
                    "scaleX": 1,
                    "scaleY": 1,
                    "translateX": 125,
                    "translateY": 87,
                    "unit": "PT"}},
            "spreadsheetId": global_dict['working sheetID'],
            "chartId": 413967291,
            "linkingMode": 'LINKED'}},
        {"createSheetsChart": {  # CHART SHOWING RELEVANT STATS ON EXPECTED RESPONSE AND PROJECTIONS
            "elementProperties": {
                "pageObjectId": 'g86bb6be987_0_114',
                "size": {
                    'width': {'magnitude': 454, 'unit': 'PT'},
                    'height': {'magnitude': 280, 'unit': 'PT'}},
                "transform": {
                    "scaleX": 1,
                    "scaleY": 1,
                    "translateX": 125,
                    "translateY": 87,
                    "unit": "PT"}},
            "spreadsheetId": global_dict['working sheetID'],
            "chartId": 773332605,
            "linkingMode": 'LINKED'}},
        {"createSheetsChart": {  # CHART SHOWING RELEVANT STATS ON EXPECTED RESPONSE AND PROJECTIONS
            "elementProperties": {
                "pageObjectId": 'g86bb6be987_0_191',
                "size": {
                    'width': {'magnitude': 454, 'unit': 'PT'},
                    'height': {'magnitude': 280, 'unit': 'PT'}},
                "transform": {
                    "scaleX": 1,
                    "scaleY": 1,
                    "translateX": 125,
                    "translateY": 87,
                    "unit": "PT"}},
            "spreadsheetId": global_dict['working sheetID'],
            "chartId": 937941452,
            "linkingMode": 'LINKED'}},
        {"createSheetsChart": {  # CHART SHOWING RELEVANT STATS ON EXPECTED RESPONSE AND PROJECTIONS
            "elementProperties": {
                "pageObjectId": 'g86bb6be987_0_203',
                "size": {
                    'width': {'magnitude': 454, 'unit': 'PT'},
                    'height': {'magnitude': 280, 'unit': 'PT'}},
                "transform": {
                    "scaleX": 1,
                    "scaleY": 1,
                    "translateX": 125,
                    "translateY": 87,
                    "unit": "PT"}},
            "spreadsheetId": global_dict['working sheetID'],
            "chartId": 577111194,
            "linkingMode": 'LINKED'}}]

    if has_years:
        yearly_reqs = [
            {"createImage": {  # CREATES CHART SHOWING NUMBER OVER DONORS BY YEAR
                "elementProperties": {
                    'pageObjectId': 'g8634961b8a_0_162',
                    'size': {
                        'width': {'magnitude': 350, 'unit': 'PT'},
                        'height': {'magnitude': 220, 'unit': 'PT'}},
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": 20,
                        "translateY": 103.5,
                        "unit": "PT"}},
                "url": str(get_file_url(global_dict['donors_by_year.jpg']))}},
            {"createImage": {  # CREATES CHART SHOWING SUM OF DONATIONS BY YEAR
                "elementProperties": {
                    'pageObjectId': 'g8634961b8a_0_162',
                    'size': {
                        'width': {'magnitude': 350, 'unit': 'PT'},
                        'height': {'magnitude': 220, 'unit': 'PT'}},
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": 350,
                        "translateY": 103.5,
                        "unit": "PT"}},
                "url": str(get_file_url(global_dict['gifts_by_year.jpg']))}},
            {"createImage": {  # CREATES CHART SHOWING NUMBER OVER DONORS BY YEAR, HIGHLIGHTING THOSE WHO GAVE OVER $5000
                "elementProperties": {
                    'pageObjectId': 'g14ae5f3b2f1f785e_0',
                    'size': {
                        'width': {'magnitude': 350, 'unit': 'PT'},
                        'height': {'magnitude': 220, 'unit': 'PT'}},
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": 20,
                        "translateY": 103.5,
                        "unit": "PT"}},
                "url": str(get_file_url(global_dict['pareto_donors.jpg']))}},
            {"createImage": {  # CREATES TABLE BELOW DONORS_BY_YEAR SHOWING RETENTION RATE FOR EACH YEAR POSSIBLE
                "elementProperties": {
                    'pageObjectId': 'g8634961b8a_0_162',
                    'size': {
                        'width': {'magnitude': 240, 'unit': 'PT'},
                        'height': {'magnitude': 93.75, 'unit': 'PT'}},
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": 64,
                        "translateY": 305,
                        "unit": "PT"}},
                "url": str(get_file_url(global_dict['donor_retention.jpg']))}},
            {"createImage": {  # CREATES TABLE BELOW GIFTS_BY_YEAR SHOWING AVERAGE DONATION PER DONOR FOR EACH YEAR POSSIBLE
                "elementProperties": {
                    'pageObjectId': 'g8634961b8a_0_162',
                    'size': {
                        'width': {'magnitude': 240, 'unit': 'PT'},
                        'height': {'magnitude': 93.75, 'unit': 'PT'}},
                    "transform": {
                        "scaleX": 1.1667,
                        "scaleY": 1.1667,
                        "translateX": 384,
                        "translateY": 297.14,
                        "unit": "PT"}},
                "url": str(get_file_url(global_dict['avg_donation.jpg']))}},
            {"createImage": {  # CREATES CHART SHOWING SUM OF DONATIONS BY YEAR, HIGHLIGHTING THE AMOUNT DERIVED FROM GIFTS OVER $5000
                "elementProperties": {
                    'pageObjectId': 'g14ae5f3b2f1f785e_0',
                    'size': {
                        'width': {'magnitude': 350, 'unit': 'PT'},
                        'height': {'magnitude': 220, 'unit': 'PT'}},
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": 350,
                        "translateY": 103.5,
                        "unit": "PT"}},
                "url": str(get_file_url(global_dict['pareto_gifts.jpg']))}},
            {"createImage": {  # CREATES CHART THE PROPORTION OF THE FILE THAT CAN BE REACHED VIA EMAIL VS ONLY DIRECT MAIL
                "elementProperties": {
                    'pageObjectId': 'g863e3227ca_0_583',
                    'size': {
                        'width': {'magnitude': 85.68, 'unit': 'PT'},
                        'height': {'magnitude': 51.84, 'unit': 'PT'}},
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": 68,
                        "translateY": 70,
                        "unit": "PT"}},
                "url": str(get_file_url(global_dict['pie_chart.jpg']))}},
            {"createImage": {  # CREATES CHART SHOWING RETENTION OF DONORS IN A SINGLE YEAR OVER THE NEXT 3 YEARS
                "elementProperties": {
                    'pageObjectId': 'g8922566a02_0_1',
                    'size': {
                        'width': {'magnitude': 350, 'unit': 'PT'},
                        'height': {'magnitude': 220, 'unit': 'PT'}},
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": 20,
                        "translateY": 103.5,
                        "unit": "PT"}},
                "url": str(get_file_url(global_dict['single_class_retention.jpg']))}},
            {"createImage": {  # CREATES CHART SHOWING HOW MANY DONORS ORG. COULD HAVE HAD GIVEN A MODEST INCREASE IN RETENTION
                "elementProperties": {
                    'pageObjectId': 'g8922566a02_0_65',
                    'size': {
                        'width': {'magnitude': 350, 'unit': 'PT'},
                        'height': {'magnitude': 220, 'unit': 'PT'}},
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": 20,
                        "translateY": 110,
                        "unit": "PT"}},
                "url": str(get_file_url(global_dict['unrealized_donors0.jpg']))}},
            {"createImage": {  # CREATES CHART SHOWING HOW MUCH MORE REVENUE ORG. COULD HAVE RECEIVED GIVEN A MODEST INCREASE IN RETENTION
                "elementProperties": {
                    'pageObjectId': 'g8922566a02_0_65',
                    'size': {
                        'width': {'magnitude': 350, 'unit': 'PT'},
                        'height': {'magnitude': 220, 'unit': 'PT'}},
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": 350,
                        "translateY": 110,
                        "unit": "PT"}},
                "url": str(get_file_url(global_dict['unrealized_gifts0.jpg']))}},
            {"createImage": {  # CREATES CHART SHOWING HOW MANY DONORS WHO GAVE OVER $5000 ORG. COULD HAVE HAD GIVEN A MODEST INCREASE IN RETENTION
                "elementProperties": {
                    'pageObjectId': 'g8922566a02_0_82',
                    'size': {
                        'width': {'magnitude': 350, 'unit': 'PT'},
                        'height': {'magnitude': 220, 'unit': 'PT'}},
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": 20,
                        "translateY": 110,
                        "unit": "PT"}},
                "url": str(get_file_url(global_dict['unrealized_donors4999.jpg']))}},
            {"createImage": {  # CREATES CHART SHOWING HOW MUCH MORE REVENUE FROM OVER $5000 DONATIONS ORG. COULD HAVE RECEIVED GIVEN A MODEST INCREASE IN RETENTION
                "elementProperties": {
                    'pageObjectId': 'g8922566a02_0_82',
                    'size': {
                        'width': {'magnitude': 350, 'unit': 'PT'},
                        'height': {'magnitude': 220, 'unit': 'PT'}},
                    "transform": {
                        "scaleX": 1,
                        "scaleY": 1,
                        "translateX": 350,
                        "translateY": 110,
                        "unit": "PT"}},
                "url": str(get_file_url(global_dict['unrealized_gifts4999.jpg']))}},
            {"replaceAllText": {  # UPDATES TEXT BOX WITH CORRECT AVERAGE MAJOR GIFT VALUE
                "replaceText": '{:,}'.format(int(yoy_obj.avg_major_gift())),
                "pageObjectIds": ['g14ae5f3b2f1f785e_0', 'g8922566a02_0_82'],
                "containsText": {"text": '{avg-mjr-gift}'}}},
            {"replaceAllText": {  # UPDATES TEXT BOX WITH CORRECT AVERAGE RETENTION RATE VALUE
                "replaceText": str(round(np.mean(yoy_obj.donor_retention(no_show=True)), 1)),
                "pageObjectIds": ['g8922566a02_0_65'],
                "containsText": {"text": '{avg-retention}'}}},
            {"replaceAllText": {  # UPDATES TEXT BOX WITH CORRECT AVERAGE RETENTION RATE TIMES 1.1 (A 10% INCREASE) VALUE FOR WHAT-IF ANALYSIS
                "replaceText": str(round(1.1*np.mean(yoy_obj.donor_retention(no_show=True)), 1)),
                "pageObjectIds": ['g8922566a02_0_65'],
                "containsText": {"text": '{avg-retention*1.1}'}}}]

        # CHECKS IF FILE PROVIDES ENOUGH YEARS TO DO MULTIYEAR RETENTION ANALYSIS
        if len(yoy_obj.years_labels) >= 7:
            yearly_reqs += [
                {"createImage": {  # CREATES CHART SHOWING RETENTION OVER THE NEXT 3 YEARS OF 4 DIFFERENT INITIAL YEARS
                    "elementProperties": {
                        'pageObjectId': 'g8b04417e30_0_1',
                        'size': {
                            'width': {'magnitude': 350, 'unit': 'PT'},
                            'height': {'magnitude': 220, 'unit': 'PT'}},
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": 350,
                            "translateY": 103.5,
                            "unit": "PT"}},
                    "url": str(get_file_url(global_dict['multi_class_retention.jpg']))}},
                {"createImage": {  # CREATES CHART TO BE OVERLAYED THE ABOVE CHART TO CLARIFY WHICH BAR CORRESPONDS TO THE SINGLE YEAR RETENTION CHART.
                    "elementProperties": {
                        'pageObjectId': 'g8922566a02_0_1',
                        'size': {
                            'width': {'magnitude': 350, 'unit': 'PT'},
                            'height': {'magnitude': 220, 'unit': 'PT'}},
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": 350,
                            "translateY": 103.5,
                            "unit": "PT"}},
                    "url": str(get_file_url(global_dict['multi_class_retention_single.jpg']))}},
                {"createImage": {  # CREATES CHART OF SINGLE YEAR RETENTION ON NEXT SLIDE.
                    "elementProperties": {
                        'pageObjectId': 'g8b04417e30_0_1',
                        'size': {
                            'width': {'magnitude': 350, 'unit': 'PT'},
                            'height': {'magnitude': 220, 'unit': 'PT'}},
                        "transform": {
                            "scaleX": 1,
                            "scaleY": 1,
                            "translateX": 20,
                            "translateY": 103.5,
                            "unit": "PT"}},
                    "url": str(get_file_url(global_dict['single_class_retention.jpg']))}}]
    else:
        yearly_reqs = []

    zorder_reqs = [  # ADJUSTS Z-ORDER OF ALL ELEMENTS TO CORRECFT FOR SLIDES API AUTOMATICALLY PLACING ALL NEW OBJECTS IN THE FRONT.
        {"updatePageElementsZOrder": {
            "pageElementObjectIds": ["g895c6cfab1_0_4"],
            "operation": 'BRING_TO_FRONT'}},
        {"updatePageElementsZOrder": {
            "pageElementObjectIds": ["g895c6cfab1_0_5"],
            "operation": 'BRING_TO_FRONT'}},
        {"updatePageElementsZOrder": {
            "pageElementObjectIds": ["g895c6cfab1_0_6"],
            "operation": 'BRING_TO_FRONT'}},
        {"updatePageElementsZOrder": {
            "pageElementObjectIds": ["g895c6cfab1_0_7"],
            "operation": 'BRING_TO_FRONT'}},
        {"updatePageElementsZOrder": {
            "pageElementObjectIds": ["g8922566a02_0_81",
                                     "g8922566a02_0_80",
                                     "g8922566a02_0_70"],
            "operation": 'BRING_TO_FRONT'}},
        {"updatePageElementsZOrder": {
            "pageElementObjectIds": ["g895c6cfab1_0_22"],
            "operation": 'BRING_TO_FRONT'}},
        {"updatePageElementsZOrder": {
            "pageElementObjectIds": ["g8922566a02_0_97",
                                     "g8922566a02_0_98",
                                     "g8922566a02_0_87"],
            "operation": 'BRING_TO_FRONT'}},
        {"updatePageElementsZOrder": {
            "pageElementObjectIds": ['g8b69c33044_0_0'],
            "operation": 'BRING_TO_FRONT'}}]

    # COMBINES REQUESTS ACCORDING TO AGES AND YEARS CONDITION
    reqs = general_reqs + age_reqs + yearly_reqs + zorder_reqs

    total_counts_placeholders = ['{n-donor-ids}', '{n-emails}', '{n-mail-addr}', '{n-no-solicit}',
                                 '{n-mjr-prosp}', '{n-legacy-prosp}', '{n-legacy-mmbers}']
    # UPDATES TEXT FOR COUNTS OF VARIOUS CATEGORIES DENOTED ABOVE IN SHORTHAND
    for i, j in zip(total_counts_placeholders, rfm_obj.total_counts()):
        reqs.append({"replaceAllText": {"replaceText": str(j),
                                        "pageObjectIds": ['g863622c97f_4_58'],
                                        "containsText": {"text": i}
                                        }})
    rfm_placeholders = {'total_rfm': ['{total-best}', '{total-high}', '{total-med}', '{total-low}'],
                        'mjr_prosp': ['{best-mjr-prosp}', '{high-mjr-prosp}', '{med-mjr-prosp}', '{low-mjr-prosp}'],
                        'legacy_prosp': ['{best-legacy-prosp}', '{high-legacy-prosp}', '{med-legacy-prosp}', '{low-legacy-prosp}'],
                        'legacy_member': ['{best-legacy-mmbr}', '{high-legacy-mmbr}', '{med-legacy-mmbr}', '{low-legacy-mmbr}'],
                        'total_all': ['{total-all}', '{total-mjr-prosp}', '{total-legacy-prosp}', '{total-legacy-mmbr}']}
    # UPDATES CELLS OF TABLE WITH COUNTS OF VARIOUS CATEGORIES DENOTED ABOVE IN SHORTHAND
    for i in rfm_obj.rfm_table().keys():
        for j, k in zip(rfm_placeholders[i], rfm_obj.rfm_table()[i]):
            reqs.append({"replaceAllText": {"replaceText": '{:,}'.format(k),
                                            "pageObjectIds": ['g863e3227ca_0_583'],
                                            "containsText": {"text": j}
                                            }})
    pyramid_placeholders = {'direct_mail': ['{dir-mail-best}', '{dir-mail-high}', '{dir-mail-med}', '{dir-mail-low}'],
                            'email': ['{email-best}', '{email-high}', '{email-med}', '{email-low}'],
                            'totals': ['{total-best}', '{total-high}', '{total-med}', '{total-low}'],
                            'total_categories': ['{dir-mail-total}', '{email-total}', '{total-all}'],
                            'total_best_high': ['{total-best-high}'],
                            'total_not_low': ['{total-not-low}']}
    # UPDATES TEXT IN THE PYRAMID RFM SLIDE (POSSIBLY DEPRECATED DEPENDING ON JEFF'S PREFERENCES)
    for i in rfm_obj.pyramid_table().keys():
        for j, k in zip(pyramid_placeholders[i], rfm_obj.pyramid_table()[i]):
            reqs.append({"replaceAllText": {"replaceText": '{:,}'.format(k),
                                            "pageObjectIds": ['g863e3227ca_0_423', 'g863e3227ca_0_593', 'g889f3efaf4_0_70', 'g889f3efaf4_0_110', 'g8a4fd4bf87_1_46', 'g8a4fd4bf87_1_2', 'g889f3efaf4_0_193'],
                                            "containsText": {"text": j}
                                            }})
    if has_years:
        unrealized_vals0 = yoy_obj.unrealized_potential(gift_floor=0)
        unrealized_vals4999 = yoy_obj.unrealized_potential(gift_floor=4999)
        unrealized_vals = {'{unrealized-amount-all}': '{:,}'.format(int(unrealized_vals0['unrealized_gifts'])),
                           '{unrealized-amount-5k+}': '{:,}'.format(int(unrealized_vals4999['unrealized_gifts'])),
                           '{pct-from-5k+}': str(round(100*unrealized_vals4999['unrealized_gifts']/unrealized_vals0['unrealized_gifts'], 1))+'%',
                           '{pct-of-donors}': str(round(100*sum(unrealized_vals4999['what_if_dif'])/sum(unrealized_vals0['n_donors']), 1))+'%',
                           '{earliest-year}': unrealized_vals0['earliest_year']}
        for i, j in unrealized_vals.items():  # UPDATES TEXT CONCERNING THE WHAT-IF ANALYSIS
            reqs.append({"replaceAllText": {"replaceText": str(j),
                                            "pageObjectIds": ['g8922566a02_0_65', 'g8922566a02_0_82'],
                                            "containsText": {"text": i}
                                            }})
        retention_placeholders = [
            '{{year[-1]}}', '{{year[-4]}}', '{{num_retained}}', '{{rate%}}']
        # UPDATES TEXT CONCERNING OVER $5000 DONORS' RETENTION OVER A FEW YEARS
        for i, j in zip(retention_placeholders, yoy_obj.over5k_retention()):
            reqs.append({"replaceAllText": {"replaceText": str(j),
                                            "pageObjectIds": ['g14ae5f3b2f1f785e_0'],
                                            "containsText": {"text": i}
                                            }})

    # UPDATES TEXT OF THE BEQUEST POTENTIAL CALCULATION SLIDE
    for i, j in zip(rfm_obj.bequest_potential(), ['{low-end}', '{high-end}']):
        reqs.append({"replaceAllText": {"replaceText": '$'+i,
                                        "pageObjectIds": ['g863e3227ca_0_593'],
                                        "containsText": {"text": j}
                                        }})

    response_placeholders = ['{n-email}', '{n-direct-mail}']
    response_audience_maj = rfm_obj.response_breakdown(gtype='major', no_show=True)[
        0]  # UPDATES TEXT OF RESPONSE BREAKDOWN FOR MAJOR GIFTS
    for i, j in zip(response_audience_maj.values(), response_placeholders):
        reqs.append({"replaceAllText": {"replaceText": '{:,}'.format(i),
                                        "pageObjectIds": ['g889f3efaf4_0_70'],
                                        "containsText": {"text": j}
                                        }})

    response_audience_legacy = rfm_obj.response_breakdown(gtype='legacy', no_show=True)[
        0]  # UPDATES TEXT OF RESPONSE BREAKDOWN FOR LEGACY GIFTS
    for i, j in zip(response_audience_legacy.values(), response_placeholders):
        reqs.append({"replaceAllText": {"replaceText": '{:,}'.format(i),
                                        "pageObjectIds": ['g889f3efaf4_0_110'],
                                        "containsText": {"text": j}
                                        }})

    # UPDATES TEXT OF RESPONSE BREAKDOWN FOR MAJOR GIFTS (INCLUDING MEDIUM RFM PEOPLE)
    response_audience_maj_med = rfm_obj.response_breakdown(
        gtype='major', medium=True, no_show=True)[0]
    for i, j in zip(response_audience_maj_med.values(), response_placeholders):
        reqs.append({"replaceAllText": {"replaceText": '{:,}'.format(i),
                                        "pageObjectIds": ['g8a4fd4bf87_1_2'],
                                        "containsText": {"text": j}
                                        }})

    # UPDATES TEXT OF RESPONSE BREAKDOWN FOR LEGACY GIFTS (INCLUDING MEDIUM RFM PEOPLE)
    response_audience_legacy_med = rfm_obj.response_breakdown(
        gtype='legacy', medium=True, no_show=True)[0]
    for i, j in zip(response_audience_maj_med.values(), response_placeholders):
        reqs.append({"replaceAllText": {"replaceText": '{:,}'.format(i),
                                        "pageObjectIds": ['g8a4fd4bf87_1_46'],
                                        "containsText": {"text": j}
                                        }})

    response_projections_placeholders_major = [
        '{deferred}', '{immediate}', '{already-gave}']
    response_projections_placeholders_legacy = [
        '{none}', '{deferred}', '{immediate}', '{already-gave}']
    response_projections_major = rfm_obj.response_breakdown(
        gtype='major', medium=False, no_show=True)[1]
    total_major = rfm_obj.response_breakdown(
        gtype='major', medium=False, no_show=True)[2]
    reqs.append({"replaceAllText": {"replaceText": '{:,}'.format(int(total_major)),
                                    "pageObjectIds": ['g889f3efaf4_0_70'],
                                    "containsText": {"text": '{total}'}
                                    }})
    # UPDATES TEXT OF RESPONSE BREAKDOWN FOR MAJOR GIFTS
    for i, j in zip(response_projections_major.values(), response_projections_placeholders_major):
        reqs.append({"replaceAllText": {"replaceText": '{:,}'.format(int(i)),
                                        "pageObjectIds": ['g889f3efaf4_0_70'],
                                        "containsText": {"text": j}
                                        }})

    response_projections_legacy = rfm_obj.response_breakdown(
        gtype='legacy', medium=False, no_show=True)[1]
    total_legacy = rfm_obj.response_breakdown(
        gtype='legacy', medium=False, no_show=True)[2]
    reqs.append({"replaceAllText": {"replaceText": '{:,}'.format(int(total_legacy)),
                                    "pageObjectIds": ['g889f3efaf4_0_110'],
                                    "containsText": {"text": '{total}'}
                                    }})
    # UPDATES TEXT OF RESPONSE BREAKDOWN FOR LEGACY GIFTS
    for i, j in zip(response_projections_legacy.values(), response_projections_placeholders_legacy):
        reqs.append({"replaceAllText": {"replaceText": '{:,}'.format(int(i)),
                                        "pageObjectIds": ['g889f3efaf4_0_110'],
                                        "containsText": {"text": j}
                                        }})

    response_projections_major_med = rfm_obj.response_breakdown(
        gtype='major', medium=True, no_show=True)[1]
    total_major_med = rfm_obj.response_breakdown(
        gtype='major', medium=True, no_show=True)[2]
    reqs.append({"replaceAllText": {"replaceText": '{:,}'.format(int(total_major_med)),
                                    "pageObjectIds": ['g8a4fd4bf87_1_2'],
                                    "containsText": {"text": '{total}'}
                                    }})
    # UPDATES TEXT OF RESPONSE BREAKDOWN FOR MAJOR GIFTS (INCLUDING MEDIUM RFM PEOPLE)
    for i, j in zip(response_projections_major_med.values(), response_projections_placeholders_major):
        reqs.append({"replaceAllText": {"replaceText": '{:,}'.format(int(i)),
                                        "pageObjectIds": ['g8a4fd4bf87_1_2'],
                                        "containsText": {"text": j}
                                        }})

    response_projections_legacy_med = rfm_obj.response_breakdown(
        gtype='legacy', medium=True, no_show=True)[1]
    total_legacy_med = rfm_obj.response_breakdown(
        gtype='legacy', medium=True, no_show=True)[2]
    reqs.append({"replaceAllText": {"replaceText": '{:,}'.format(int(total_legacy_med)),
                                    "pageObjectIds": ['g8a4fd4bf87_1_46'],
                                    "containsText": {"text": '{total}'}
                                    }})
    # UPDATES TEXT OF RESPONSE BREAKDOWN FOR LEGACY GIFTS (INCLUDING MEDIUM RFM PEOPLE)
    for i, j in zip(response_projections_legacy_med.values(), response_projections_placeholders_legacy):
        reqs.append({"replaceAllText": {"replaceText": '{:,}'.format(int(i)),
                                        "pageObjectIds": ['g8a4fd4bf87_1_46'],
                                        "containsText": {"text": j}
                                        }})
    marketing_mix_placeholders = [
        '{email-total}', '{dir-mail-not-low}', '{total-audience}']
    marketing_mix = list(rfm_obj.response_breakdown(
        no_show=True, medium=True)[3].values())
    marketing_mix.append(sum(marketing_mix))
    for i, j in zip(marketing_mix, marketing_mix_placeholders):
        reqs.append({"replaceAllText": {"replaceText": '{:,}'.format(int(i)),
                                        "pageObjectIds": ['g8a4fd4bf87_3_25'],
                                        "containsText": {"text": j}
                                        }})

    presentation = slides_service.presentations()
    # EXECUTES ABOVE REQUESTS
    presentation.batchUpdate(
        body={"requests": reqs}, presentationId=presID).execute()
    print('CHANGES APPLIED')


# SENDS UPDATE REQUEST TO SHEETS API TO INPUT RELEVANT VALUES IN THE SPREADSHEET TOOL
def batch_update_sheet(sheetID):
    print('APPLYING CHANGES TO SHEET')
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
             "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        "creds.json", scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheetID)
    sheet1 = sheet.sheet1
    sheet2 = sheet.get_worksheet(2)
    sheet3 = sheet.get_worksheet(4)
    table_sheet = sheet.get_worksheet(1)
    if has_ages:
        age_dist = [[round(i, 2)/100]
                    for i in rfm_obj.age_distribution(no_show=True)]
        sheet1.update('B3:B8', age_dist)
        rfm_by_age = rfm_obj.countby_age_rfm()
        age_percentile = rfm_obj.countby_age_percentile()
        table_sheet.update(
            'A2:C30', [rfm_by_age.columns.values.tolist()] + rfm_by_age.values.tolist())
        table_sheet.update('E2:G9', [
                           age_percentile.columns.values.tolist()] + age_percentile.values.tolist())

    if 'email on file' in rfm_obj.rfm_data.columns:
        rfm_email = rfm_obj.countby_email_rfm()
        table_sheet.update(
            'E12:G16', [rfm_email.columns.values.tolist()] + rfm_email.values.tolist())

    if 'email on file' in rfm_obj.rfm_data.columns and 'physical address on file' in rfm_obj.rfm_data.columns:
        rfm_mail = rfm_obj.countby_physical_mail_rfm()
        table_sheet.update(
            'E19:H23', [rfm_mail.columns.values.tolist()] + rfm_mail.values.tolist())

    if 'legacy_member' in rfm_obj.rfm_data.columns:
        rfm_legacy = rfm_obj.countby_legacy_rfm()
        table_sheet.update(
            'I2:K6', [rfm_legacy.columns.values.tolist()] + rfm_legacy.values.tolist())

    if 'managed_prospect' in rfm_obj.rfm_data.columns:
        rfm_prospect = rfm_obj.countby_managed_prospect_rfm()
        table_sheet.update(
            'I9:K14', [rfm_prospect.columns.values.tolist()] + rfm_prospect.values.tolist())

    if 'dns_mail' in rfm_obj.rfm_data.columns:
        rfm_dns_mail = rfm_obj.countby_dns_mail()
        table_sheet.update(
            'I23:K27', [rfm_dns_mail.columns.values.tolist()] + rfm_dns_mail.values.tolist())

    if 'dns_email' in rfm_obj.rfm_data.columns:
        rfm_dns_email = rfm_obj.countby_dns_email()
        table_sheet.update('I30:K34', [
                           rfm_dns_email.columns.values.tolist()] + rfm_dns_email.values.tolist())

    rfm_counts = rfm_obj.countby_rfm()
    table_sheet.update(
        'E26:G30', [rfm_counts.columns.values.tolist()] + rfm_counts.values.tolist())

    response_data = [[int(i)] for i in rfm_obj.response_breakdown(
        no_show=True, medium=True)[3].values()]
    sheet2.update('B3:B4', response_data)

    rfm_table = rfm_obj.brief_rfm()
    rfm_table.to_csv('providence_table.csv')
    # sheet3.update([rfm_table.columns.values.tolist()] + rfm_table.values.tolist())

    print('CHANGES APPLIED')


# GETS URL FOR IMAGE OF A CHART THAT NEEDS TO BE PLACED IN THE PRESENTATION
def get_file_url(img_fileID):
    url = drive_service.files()\
        .get(fileId=img_fileID, fields='thumbnailLink')\
        .execute()['thumbnailLink'][:-5]
    return url


# UPLOADS ALL FILES IN A FOLDER (/TEMP PLOTS BY DEFAULT) TO A SPECIFIED FOLDER IN GOOGLE DRIVE
def upload_multiple_files_to_folder(drive_folder_id, folder_path='/temp plots/'):
    for file in os.listdir(os.getcwd()+folder_path):
        upload_image_to_folder(os.getcwd()+folder_path +
                               file, drive_folder_id, file)
        os.remove(os.getcwd()+folder_path+file)
    os.rmdir(os.getcwd()+folder_path)


# DUPLICATES THE PRESENTATION TEMPLATE AND RENAMES IT AND RETURNS THE DUPLICATES ID
def duplicate_pres(presID, folder_id, new_name):
    global global_dict
    body = {'name': new_name,
            'parents': [folder_id]}
    drive = drive_service.files().copy(fileId=presID, body=body).execute()
    new_presID = drive.get('id')
    global_dict['working presID'] = drive.get('id')
    print('NEW PRESENTATION CREATED')
    return new_presID


# DUPLICATES GOOGLE SHEET TO HELP JEFF WITH CALCULATIONS AND CHART ADJUSTMENTS
def duplicate_sheet(sheetID, folder_id, new_name):
    global global_dict
    body = {'name': new_name,
            'parents': [folder_id]}
    drive = drive_service.files().copy(fileId=sheetID, body=body).execute()
    new_sheetID = drive.get('id')
    global_dict['working sheetID'] = drive.get('id')
    print('NEW SHEETS CREATED')
    service_account = 'oppanalysis@opportunity-analysis-285715.iam.gserviceaccount.com'
    share_sheet(global_dict['working sheetID'], email=service_account)
    return new_sheetID


# SHARES THE PRESENTATION WITH THE SPECIFIED EMAIL ADDRESS
def share_pres(presID, email="lwarner@imarketsmart.com"):
    reqs = {"role": "writer",
            "type": "user",
            "emailAddress": email}
    perm = drive_service.permissions().create(body=reqs, fileId=presID).execute()
    print(f'SHARED PRESENTATION WITH {email}')


# SHARES THE SHEET WITH SPECIFIED EMAIL ADDRESS. USED TO SHARE WITH SERVICE ACCOUNT.
def share_sheet(sheetID, email="lwarner@imarketsmart.com"):
    reqs = {"role": "writer",
            "type": "user",
            "emailAddress": email}
    perm = drive_service.permissions().create(body=reqs, fileId=sheetID).execute()
    print(f'SHARED SHEET WITH {email}')


def share_folder(folderID, email='lwarner@imarketsmart.com'):
    reqs = {"role": "writer",
            "type": "user",
            "emailAddress": email}
    perm = drive_service.permissions().create(body=reqs, fileId=folderID).execute()
    print(f'SHARED FOLDER WITH {email}')

# data = get_sheet_info(analysis_sheet)
# with open('data/data2.json', 'w', encoding='utf-8') as f:
#     json.dump(data, f, ensure_ascii=False, indent=4)


if __name__ == '__main__':  # RUNS NECESSARY FUNCTIONS FOR A NEW PRESENTATION TO BE GENERATED CORRECTLY
    parse_args()
    print(org_name)
    rfm_obj = oa.RFM(df)
    yoy_obj = oa.Yearly(df)
    has_ages = rfm_obj.has_ages
    has_years = yoy_obj.has_years
    print(has_ages)
    print(has_years)
    folderid = create_folder(f'{org_name} Opp Analysis')
    duplicate_pres(template, folderid, f'{org_name} - Opp Analysis')
    duplicate_sheet(analysis_sheet, folderid,
                    f'{org_name} - Opp Analysis Sheet {sheet_rand_id}')
    rfm_obj.save_plots()
    yoy_obj.save_plots()
    upload_multiple_files_to_folder(folderid)
    batch_update_sheet(global_dict['working sheetID'])
    batch_update_pres(global_dict['working presID'])
    share_folder(folderid, recipient)
    share_pres(global_dict['working presID'], recipient)
    share_sheet(global_dict['working sheetID'], recipient)
    print('DONE')

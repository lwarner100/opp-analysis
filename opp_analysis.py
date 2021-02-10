import pandas as pd
import numpy as np
from matplotlib import pyplot as plt
from matplotlib import font_manager as fm
from matplotlib import patheffects as path_effects
from argparse import ArgumentParser
import seaborn as sns
import os
from datetime import date

prop = fm.FontProperties(fname='fonts1/HelveticaNeueLt.ttf')
prop_bold = fm.FontProperties(fname='fonts1/HelveticaNeueMed.ttf')
prop_legend = fm.FontProperties(fname='fonts1/HelveticaNeueLt.ttf', size=8)



def parse_args():
    global df
    parser = ArgumentParser(description="Visualization Parser")
    parser.add_argument('file_path', action='store',
                        help='The file path of the csv')
    args = parser.parse_args()
    file_path = str(args.file_path)
    df = pd.read_csv(file_path)


class RFM:  # CLASS OF OBJECT FROM DATAFRAME TO ALLOW FOR GENERATION OF ALL NECESSARY AND METRICS AND VISUALS RELATED TO RFM SCORE AND MORE
    def __init__(self, data):
        self.data = data
        self.rfm_data = self.RFMify()
        self.years_labels = list(
            filter(lambda x: str(x).startswith('20'), data.columns))
        self.has_years = len(self.years_labels) > 2
        if self.has_years:
            self.years_ints = [int(i) for i in self.years_labels]
            self.max_yr_is_this_yr = date.today().year == max(self.years_ints)
            self.years_labels = self.years_labels[:-
                                                  1] if self.max_yr_is_this_yr else self.years_labels
            self.avg_mjr_gift = sum([self.data[self.data[year] >= 5000][year].sum() for year in self.years_labels])\
                / sum([len(self.data[self.data[year] >= 5000][year]) for year in self.years_labels])
        try:
            self.age_proportion = (self.rfm_data[(self.rfm_data.age != 0) & (
                self.rfm_data.age.notna())].donor_id.count()/len(self.rfm_data))
            self.over50_age_proportion = (self.rfm_data[self.rfm_data.age >= 50].donor_id.count(
            )/len(self.rfm_data[self.rfm_data.age != 0]))
            # self.has_ages = self.age_proportion > 0.05
            self.has_ages = True
        except:
            self.age_proportion = 0
            self.has_ages = False

    def RFMify(self):  # CONVERTS RAW CSV DATAFRAME TO DATAFRAME WITH RFM INFORMATION APPENDED TO EACH DONOR
        rfm_dict = {}
        # donor_id,age,email on file,address on file,dns_mail,dns_email,legacy_prospect,gift_sum,days_since_last_gift,gift_frequency,2013,2014,2015,2016,2017,2018,2019,2020
        variations = {'donor_id': ['conid', 'id', 'constituent code', 'donor id'],
                      'age': ['AGE'],
                      'gift_sum': ['gift sum', 'gift_sum'],
                      'email_on_file': ['email on file', 'email on file?', 'Valid_Email'],
                      'address_on_file': ['address on file', 'address on file?', 'physical address on file', 'physical address on file?', 'send mail', 'send mail?', 'Valid_Address'],
                      'dns_mail': ['dns mail', 'do not solicit mail', 'do not solicit by mail'],
                      'dns_email': ['dns email', 'do not solicit email', 'do not solicit by email'],
                      'dns': ['Solicit_Code_Description', 'Do Not Solicit'],
                      'legacy_member': ['legacy society', 'legacy society member', 'legacy member', 'Legacy Society Member', 'legacy_society_member'],
                      'managed_prospect': ['assigned/managed prospect', 'assigned/managed_prospect', 'managed prospect', 'Assigned Major Donor', 'assigned_major_donor'],
                      'legacy_prospect': ['legacy prospect', 'Legacy Society', 'Assigned Legacy Donor', 'assigned_legacy_donor'],
                      'gift_frequency': ['gift frequency'],
                      'days_since_last_gift': ['days since last gift', 'gift recency']}

        def frequency_score(freq):
            # return [1 if (i == 1)
            # else 2 if (i >= 2) and (i < 5)
            # else 3 if (i >= 5) and (i < 15)
            # else 4 if (i >= 15) and (i < 20)
            # else 5 if (i >= 20)
            # else 0 if (i==0) or (np.isnan(i))
            # else None for i in freq]
            return [1 if (i == 1)
                    else 2 if (i >= 2) and (i < 5)
                    else 3 if (i >= 5) and (i < 7)
                    else 4 if (i >= 7) and (i < 9)
                    else 5 if (i >= 9)
                    else 0 if (i == 0) or (np.isnan(i))
                    else None for i in freq]

        def monetary_score(gifts):
            return [1 if (i < 100) and (i != 0)
                    else 2 if (i >= 100) and (i < 1000)
                    else 3 if (i >= 1000) and (i < 5000)
                    else 4 if (i >= 5000) and (i < 10000)
                    else 5 if (i >= 10000)
                    else 0 if (i == 0) or (np.isnan(i))
                    else None for i in gifts]

        def recency_score(rec):
            return [5 if (i < 365) and (i != 0)
                    else 4 if (i >= 365) and (i < 730)
                    else 3 if (i >= 730) and (i < 1095)
                    else 2 if (i >= 1095) and (i < 1460)
                    else 1 if (i >= 1460)
                    else 0 if (i == 0) or (np.isnan(i))
                    else None for i in rec]

        def normalize(rfm):
            maximum = rfm_dict['max']
            minimum = rfm_dict['min']
            return (rfm - minimum)/(maximum - minimum)

        def rfm_group(rfm):
            return ['Low' if (i <= 0.25)
                    else 'Medium' if (i > 0.25) and (i <= 0.66)
                    else 'High' if (i > 0.66) and (i <= 0.99)
                    else 'Best' if (i > 0.99)
                    else None for i in rfm]

        def age_group(age):
            if (age == 0):
                return 'No age given'
            elif (age < 50) and (age > 0):
                return 'Under 50'
            elif (age >= 50) and (age < 60):
                return '50-59'
            elif (age >= 60) and (age < 70):
                return '60-69'
            elif (age >= 70) and (age < 80):
                return '70-79'
            elif (age >= 80) and (age < 90):
                return '80-89'
            elif (age >= 90):
                return 'Over 90'

        def map_variations(cols):
            new_cols = []
            keys = variations.keys()
            for col in cols:
                found_variation = None
                if str(col).lower() in keys:
                    new_cols.append(str(col).lower())
                    continue
                else:
                    for key in keys:
                        if str(col).lower() in variations[key]:
                            new_cols.append(key)
                            found_variation = True
                            break
                        else:
                            found_variation = False
                    if not found_variation:
                        new_cols.append(col)
            return new_cols

        self.data.columns = map_variations(self.data.columns)
        rfm_data = pd.DataFrame()
        rfm_data['donor_id'] = self.data.donor_id
        try:
            rfm_data['age'] = self.data.age
            rfm_data['age_group'] = self.data.age.apply(age_group)
        except:
            pass
        # donor_id,age,email on file,address on file,dns_mail,dns_email,legacy_prospect,gift_sum,days_since_last_gift,gift_frequency,2013,2014,2015,2016,2017,2018,2019,2020
        rfm_data['gift_sum'] = self.data.gift_sum
        try:
            rfm_data['email on file'] = self.data['email_on_file']
        except:
            try:
                rfm_data['email on file'] = self.data['Email on file']
            except:
                try:
                    rfm_data['email on file'] = self.data['email on file']
                except:
                    pass
        try:
            rfm_data['physical address on file'] = self.data['address_on_file']
        except:
            try:
                rfm_data['physical address on file'] = self.data['address on file']
            except:
                pass
        rfm_data['frequency_score'] = frequency_score(self.data.gift_frequency)
        rfm_data['recency_score'] = recency_score(
            self.data.days_since_last_gift)
        rfm_data['monetary_score'] = monetary_score(self.data.gift_sum)
        rfm_data['rfm_score'] = ((1.3 * rfm_data.recency_score) +
                                 rfm_data.monetary_score + rfm_data.frequency_score)/3
        try:
            rfm_data['legacy_member'] = self.data['legacy_member']
        except:
            try:
                rfm_data['legacy_member'] = self.data['Legacy Society']
            except:
                try:
                    rfm_data['legacy_member'] = self.data['Legacy Society Member']
                except:
                    pass
        try:
            rfm_data['managed_prospect'] = self.data['assigned_major_donor']
        except:
            try:
                rfm_data['managed_prospect'] = self.data['managed_prospect']
            except:
                pass
        try:
            rfm_data['legacy_prospect'] = self.data['assigned_legacy_donor']
        except:
            try:
                rfm_data['legacy_prospect'] = self.data['legacy_prospect']
            except:
                pass
        try:
            # INPUT FINALIZED DO NOT SOLICIT COLUMN HERE
            rfm_data['dns_email'] = self.data['dns_email']
        except:
            pass
        try:
            rfm_data['dns_mail'] = self.data['dns_mail']  # SAME AS ABOVE
        except:
            pass
        rfm_dict['max'] = rfm_data.rfm_score.max()
        rfm_dict['min'] = rfm_data.rfm_score.min()
        rfm_data['rfm_normed'] = rfm_data.rfm_score.apply(normalize)
        rfm_data['rfm_group'] = rfm_group(rfm_data.rfm_normed)
        # print(rfm_data[rfm_data['dns_mail'].notnull()])
        return rfm_data

    def brief_rfm(self):
        data = self.rfm_data
        new_cols = ['donor_id', 'age', 'rfm_group', 'rfm_score', 'email on file',
                    'physical address on file',
                    'dns_email', 'dns_mail', 'gift_sum',
                    'rfm_normed', 'legacy_member', 'managed_prospect', 'legacy_prospect']
        data = data.drop(['age_group', 'frequency_score',
                          'recency_score', 'monetary_score'], axis=1, inplace=False)
        data = data.reindex(new_cols, axis=1)
        data = data.fillna('')
        return data

    # GENERATES BAR CHART SHOWING AGE DISTRIBUTION OF THE ORG'S DONORS
    def age_distribution(self, legacy_only=False, no_show=False, save=False):
        ages = self.rfm_data.groupby('age_group').donor_id.count() if not legacy_only\
            else self.rfm_data[(self.rfm_data.legacy_member == 'Yes') | (self.rfm_data.legacy_member == 'Y')]\
                   .groupby('age_group').donor_id.count()
        try:
            ages = ages.drop('No age given')
        except:
            pass
        age_labels = ['Under 50', '50-59',
                      '60-69', '70-79', '80-89', 'Over 90']
        age_counts = {i: 0 for i in age_labels}
        for i in ages.index:
            age_counts[i] = ages[i]
        total_donors = sum(age_counts.values())
        age_percentages = [x*100 / total_donors for x in age_counts.values()]
        max_y = max(age_percentages)
        fig = plt.figure(figsize=(6, 3.7))
        ax = plt.subplot()
        ax.grid(axis='x', alpha=0.5)
        ax.set_axisbelow(True)
        ax.set_yticks(range(len(age_percentages)))
        ax.set_yticklabels(age_labels, fontproperties=prop)
        ax.spines['left'].set_color('#404041')
        ax.set_xticks(range(0, int(max_y)+10, 5))
        ax.set_xticklabels(
            [str(x)+'%' for x in range(0, int(max_y)+10, 5)], fontproperties=prop)
        ax.tick_params(axis='y', colors='#404041')
        ax.tick_params(axis='x', colors='#404041')
        plt.barh(range(len(age_percentages)), age_percentages, color='#00a7e0')
        sns.despine(top=True, bottom=True, right=True)
        for index, value in enumerate(age_percentages):
            value = round(value, 1)
            if value >= 10:
                plt.text(value - max_y/10, index - 0.075,
                         str(value)+'%', color='w', fontproperties=prop)
            elif value < 4:
                plt.text(value + max_y/450, index - 0.075, str(value) +
                         '%', color='#404041', fontproperties=prop)
            else:
                plt.text(value - max_y/13, index - 0.075,
                         str(value) + '%', color='w', fontproperties=prop)
        if save:
            if legacy_only:
                plt.savefig('temp plots/age_dist_legacy.jpg',
                            dpi=300, bbox_inches='tight', pad_inches=0)
            else:
                plt.savefig('temp plots/age_dist.jpg', dpi=300,
                            bbox_inches='tight', pad_inches=0)
            plt.close()
        elif not no_show and not save:
            plt.show()
        return age_percentages

    # GENERATES BAR CHART SHOWING AGE DISTRIBUTION OF THE ORG'S DONORS COMPARED TO MARKETSMARTS HISTORICAL AVERAGES
    def age_distribution_compare(self, legacy_only=False, save=False):
        from matplotlib.cbook import get_sample_data
        ages = self.rfm_data.groupby('age_group').donor_id.count() if not legacy_only\
            else self.rfm_data[(self.rfm_data.legacy_member == 'Yes') | (self.rfm_data.legacy_member == 'Y')]\
                   .groupby('age_group').donor_id.count()
        try:
            ages = ages.drop('No age given')
        except:
            pass
        age_labels = ['Under 50', '50-59',
                      '60-69', '70-79', '80-89', 'Over 90']
        age_counts = {i: 0 for i in age_labels}
        for i in ages.index:
            age_counts[i] = ages[i]
        total_donors = sum(age_counts.values())
        age_percentages = [x*100 / total_donors for x in age_counts.values()]
        marketsmart_age_all = [36177, 47232, 103690, 92793, 37009, 4776]
        marketsmart_age_legacy_only = [902, 1557, 3964, 3902, 1656, 248]
        marketsmart_values = [round(100*i/sum(marketsmart_age_all), 1) for i in marketsmart_age_all] if not legacy_only\
            else [round(100*i/sum(marketsmart_age_legacy_only), 1) for i in marketsmart_age_legacy_only]
        age_percentages = [x*100 / total_donors for x in age_counts.values()]
        max_y = max(age_percentages) if max(age_percentages) > max(
            marketsmart_values) else max(marketsmart_values)
        age_labels = ['Under 50', '50-59',
                      '60-69', '70-79', '80-89', 'Over 90']
        fig = plt.figure(figsize=(6, 3.7))
        ax = plt.subplot()
        ax.grid(axis='x', alpha=0.5)
        ax.set_axisbelow(True)
        ax.set_yticks([i-0.4 for i in range(2*len(age_percentages))[::2]])
        ax.set_yticklabels(age_labels, fontproperties=prop)
        ax.spines['left'].set_color('#404041')
        ax.set_xticks(range(0, int(max_y)+10, 5))
        ax.set_xticklabels(
            [str(x)+'%' for x in range(0, int(max_y)+10, 5)], fontproperties=prop)
        ax.tick_params(axis='y', colors='#404041')
        ax.tick_params(axis='x', colors='#404041')
        plt.barh(range(2*len(age_percentages))
                 [::2], age_percentages, color='#00a7e0', label='Your Data')
        plt.barh([i-0.8 for i in range(2*len(age_percentages))[::2]],
                 marketsmart_values, color='#ffce01', label='MarketSmart\'s Data')
        im = plt.imread(get_sample_data(
            os.getcwd()+'/images/marketsmart-logo-alt.png'))
        plt.legend(prop=prop_legend, handlelength=0.8,
                   labelspacing=0.25, framealpha=1.0)
        sns.despine(top=True, bottom=True, right=True)
        for index, value in enumerate(age_percentages):
            index *= 2
            value = round(value, 1)
            if value >= 10:
                plt.text(value - max_y/10, index - 0.15, str(value) +
                         '%', color='w', fontproperties=prop)
            elif value < 4:
                plt.text(value + max_y/450, index - 0.15, str(value) +
                         '%', color='#404041', fontproperties=prop)
            else:
                plt.text(value - max_y/13, index - 0.15, str(value) +
                         '%', color='w', fontproperties=prop)
        for index, value in enumerate(marketsmart_values):
            index *= 2
            index += -0.8
            value = round(value, 1)
            plt.text(value + 0.15, index - 0.15, str(value) +
                     '%', color='#404041', fontproperties=prop)
        for index, value in enumerate(marketsmart_values):
            if value > 5:
                newax = fig.add_axes(
                    [0.07875 + (value-(value*0.26))/max_y, index/8.25+0.1475, 0.04, 0.04])
                newax.imshow(im)
                newax.axis('off')
        if save:
            if legacy_only:
                plt.savefig('temp plots/age_dist_compare_legacy.jpg',
                            dpi=500, bbox_inches='tight', pad_inches=0)
            else:
                plt.savefig('temp plots/age_dist_compare.jpg',
                            dpi=500, bbox_inches='tight', pad_inches=0)
            plt.close()
        else:
            plt.show()

    # RETURNS A DATAFRAME OF NUMBER OF DONORS IN EACH AGE GROUP AND RFM GROUP
    def countby_age_rfm(self):
        table = self.rfm_data.groupby(['rfm_group', 'age_group'])\
                    .size().reset_index()
        table.rename(columns={'age_group': 'age_group',
                              0: 'count'}, inplace=True)
        return table

    def countby_rfm(self):
        table = self.rfm_data.groupby('rfm_group').size().reset_index()
        table.rename(columns={'age_group': 'age_group',
                              0: 'count'}, inplace=True)
        return table

    def countby_dns_email(self):
        table = self.rfm_data[self.rfm_data.dns_email == 'Y'].groupby(
            ['dns_email', 'rfm_group']).size().reset_index()
        table.rename(columns={'age_group': 'age_group',
                              0: 'count'}, inplace=True)
        return table

    def countby_dns_mail(self):
        table = self.rfm_data[self.rfm_data.dns_mail == 'Y'].groupby(
            ['dns_mail', 'rfm_group']).size().reset_index()
        table.rename(columns={'age_group': 'age_group',
                              0: 'count'}, inplace=True)
        return table

    # RETURNS A DATAFRAME OF NUMBER OF DONORS WITH AN EMAIL ON FILE BY RFM GROUP
    def countby_email_rfm(self):
        table = self.rfm_data[(self.rfm_data['email on file'] == 'Y')
                              | (self.rfm_data['email on file'] == 'Yes')]\
            .groupby(['email on file', 'rfm_group']).size().reset_index()
        table.rename(columns={0: 'count'}, inplace=True)
        return table

    # RETURNS A DATAFRAME OF NUMBER OF DONORS WITH A PHYSICAL ADDRESS BUT NO EMAIL BY RFM GROUP
    def countby_physical_mail_rfm(self):
        table = self.rfm_data[((self.rfm_data['physical address on file'] == 'Y'))
                              & (self.rfm_data['email on file'].isna())]\
            .groupby(['physical address on file', 'rfm_group']).size().reset_index()

        table.rename(columns={0: 'count'}, inplace=True)
        return table

    # RETURNS A DATAFRAME OF NUMBER OF DONORS IN LEGACY SOCIETY BY RFM GROUP
    def countby_legacy_rfm(self):
        table = self.rfm_data[(self.rfm_data['legacy_member'] == 'Y')
                              | (self.rfm_data['legacy_member'] == 'Yes')]\
            .groupby(['legacy_member', 'rfm_group']).size().reset_index()
        table.rename(columns={0: 'count'}, inplace=True)
        return table

    # RETURNS A DATAFRAME OF NUMBER OF DONORS IN THAT ARE MANAGED PROSPECTS BY RFM GROUP
    def countby_managed_prospect_rfm(self):
        table = self.rfm_data[(self.rfm_data['managed_prospect'] == 'Y')
                              | (self.rfm_data['managed_prospect'] == 'Yes')]\
            .groupby(['managed_prospect', 'rfm_group']).size().reset_index()
        table.rename(columns={0: 'count'}, inplace=True)
        return table

    # RETURNS A DATAFRAME OF NUMBER OF DONORS WITH A PHYSICAL ADDRESS BY EMAIL ON FILE (Y/N) AND RFM GROUP
    def countby_physical_and_email_rfm(self):
        table = self.rfm_data[((self.rfm_data['physical address on file'] == 'Y')
                               | (self.rfm_data['physical address on file'] == 'Yes'))]\
            .groupby(['physical address on file', 'email on file', 'rfm_group']).size().reset_index()
        table.rename(columns={0: 'count'}, inplace=True)
        return table

    # RETURNS A DATAFRAME OF NUMBER OF DONORS IN AN AGE GROUP AND THE PERCENTAGE OF TOTAL IN THAT GROUP
    def countby_age_percentile(self):
        table = self.rfm_data.groupby(['age_group']).size().reset_index()
        table.rename(columns={'age_group': 'age_group',
                              0: 'count'}, inplace=True)
        table['percentile'] = round(table['count']/table['count'].sum()*100, 2)
        return table

    # GENERATES A PIE CHART SHOWING THE SHARE OF THE FILE THAT ONLY HAS PHYSICAL ADDRESS VS. EMAIL ONLY OR BOTH
    def print_online_piechart(self, save=False):
        fig, ax = plt.subplots()
        color_palette_list = ['#009ACD', '#ADD8E6', '#63D1F4', '#0EBFE9',
                              '#C1F0F6', '#0099CC']
        plt.rcParams['font.sans-serif'] = 'Arial'
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['text.color'] = '#404041'
        plt.rcParams['axes.labelcolor'] = '#404041'
        plt.rcParams['xtick.color'] = '#404041'
        plt.rcParams['ytick.color'] = '#404041'
        plt.rcParams['font.size'] = 28
        labels = ['Print', 'Email']
        mail_only = self.pyramid_table()['total_categories'][0]
        email_or_both = self.pyramid_table()['total_categories'][1]
        total = mail_only+email_or_both
        print(mail_only)
        print(email_or_both)
        percentages = [float(mail_only/total)*100,
                       float(email_or_both/total)*100]
        explode = (0.1, 0)
        ax.pie(percentages, labels=labels,
               colors=['#004f6e', '#00a7e0'], startangle=-60, labeldistance=1.15)
        ax.axis('equal')
        my_circle = plt.Circle((0, 0), 0.65, color='white')
        p = plt.gcf()
        p.gca().add_artist(my_circle)
        if save:
            plt.savefig('temp plots/pie_chart.jpg', dpi=150,
                        bbox_inches='tight', pad_inches=0)
            plt.close()
        else:
            plt.show()

    def total_counts(self):  # RETURNS A LIST OF NUMBER OF TOTAL DONORS IN RELEVANT CATEGORIES TO BE APPLIED TO A TEXT BOX IN THE SLIDES

        n_donor_ids = self.rfm_data.donor_id.count()
        n_emails = self.rfm_data[self.rfm_data['email on file']
                                 == 'Y'].donor_id.count()
        n_mail_addr = self.rfm_data[self.rfm_data['physical address on file']
                                    == 'Y'].donor_id.count()
        n_no_solicit = len(np.union1d(self.rfm_data[self.rfm_data.dns_email == 'Y'].donor_id, self.rfm_data[self.rfm_data.dns_mail == 'Y'].donor_id))\
            if {'dns_email', 'dns_mail'}.issubset(set(self.rfm_data.columns)) else 0
        n_mjr_prosp = self.rfm_data[self.rfm_data['managed_prospect'] == 'Y'].donor_id.count()\
            if 'managed_prospect' in self.rfm_data.columns else 0
        n_legacy_prosp = self.rfm_data[self.rfm_data['legacy_prospect'] == 'Y'].donor_id.count()\
            if 'legacy_prospect' in self.rfm_data.columns else 0
        n_legacy_mmbr = self.rfm_data[self.rfm_data['legacy_member'] == 'Y'].donor_id.count()\
            if 'managed_prospect' in self.rfm_data.columns else 0
        results = [n_donor_ids, n_emails, n_mail_addr, n_no_solicit,
                   n_mjr_prosp, n_legacy_prosp, n_legacy_mmbr]
        results = ['{:,}'.format(i) for i in results]
        return results

    def rfm_table(self):  # RETURNS A DICTIONARY THAT FILLS A TABLE IN THE SLIDES WITH RELEVANT INFORMATION
        total_rfm = self.rfm_data.groupby('rfm_group').donor_id.count().reindex(
            ['Best', 'High', 'Medium', 'Low']).fillna(0).to_list()
        mjr_prosp = self.rfm_data[self.rfm_data['managed_prospect'] == 'Y'].groupby('rfm_group').size().reindex(['Best', 'High', 'Medium', 'Low']).fillna(0).to_list()\
            if 'managed_prospect' in self.rfm_data.columns else [0, 0, 0, 0]
        legacy_prosp = self.rfm_data[self.rfm_data['legacy_prospect'] == 'Y'].groupby('rfm_group').size().reindex(['Best', 'High', 'Medium', 'Low']).fillna(0).to_list()\
            if 'legacy_prospect' in self.rfm_data.columns else [0, 0, 0, 0]
        legacy_member = self.rfm_data[self.rfm_data['legacy_member'] == 'Y'].groupby('rfm_group').size().reindex(['Best', 'High', 'Medium', 'Low']).fillna(0).to_list()\
            if 'legacy_member' in self.rfm_data.columns else [0, 0, 0, 0]
        legacy_member = [int(i) if type(i) == float and not np.isnan(
            i) else 0 for i in legacy_member]
        totals = [sum(total_rfm), sum(mjr_prosp), sum(
            legacy_prosp), sum(legacy_member)]
        results = {'total_rfm': total_rfm, 'mjr_prosp': mjr_prosp,
                   'legacy_prosp': legacy_prosp, 'legacy_member': legacy_member, 'total_all': totals}
        return results

    # RETURNS A DICTIONARY OF RELEVANT INFORMATION USED IN VARIOUS PARTS OF THE SLIDES
    def pyramid_table(self):
        if {'dns_email', 'dns_mail'}.issubset(set(self.rfm_data.columns)):
            email = self.rfm_data[(self.rfm_data['email on file'] == 'Y') & (self.rfm_data['dns_email'].isna())]\
                .groupby('rfm_group').donor_id.count().reindex(['Best', 'High', 'Medium', 'Low']).fillna(0)
            direct_mail = self.rfm_data[(self.rfm_data['physical address on file'] == 'Y') & (self.rfm_data['email on file'] == 'N') & (self.rfm_data['dns_mail'].isna())]\
                .groupby('rfm_group').donor_id.count().reindex(['Best', 'High', 'Medium', 'Low']).fillna(0)
        else:
            email = self.rfm_data[(self.rfm_data['email on file'] == 'Y')].groupby(
                'rfm_group').donor_id.count().reindex(['Best', 'High', 'Medium', 'Low']).fillna(0)
            direct_mail = self.rfm_data[(self.rfm_data['physical address on file'] == 'Y') & (self.rfm_data['email on file'] == 'N')]\
                .groupby('rfm_group').donor_id.count().reindex(['Best', 'High', 'Medium', 'Low']).fillna(0)
        print(self.rfm_data)
        print(self.rfm_data[(self.rfm_data['physical address on file'] == 'Y') & (
            self.rfm_data['email on file'] == 'N')])
        print(direct_mail)
        asdadada
        totals = [sum([i, j]) for i, j in zip(direct_mail, email)]
        # SUM OF BEST AND HIGH RFM DONORS
        total_best_high = [totals[0]+totals[1]]
        total_not_low = [totals[0]+totals[1]+totals[2]]
        total_categories = [sum(direct_mail), sum(email), sum(totals)]
        results = {'direct_mail': direct_mail, 'email': email, 'totals': totals,
                   'total_categories': total_categories, 'total_best_high': total_best_high, 'total_not_low': total_not_low}
        return results

    # RETURNS A LIST CONTAINING THE LOW-HIGH ESTIMATE OF BEQUEST POTENTIAL TO BE USED IN THE SLIDES
    def bequest_potential(self):
        n_donors = sum(self.pyramid_table()['totals'][:2])
        return ['{:,}'.format(1520*n_donors), '{:,}'.format(6270*n_donors)]

    # GENERATES CHARTS SHOWING THE EXPECTED RESPONSES OF A SURVEY BLAST BASED ON MARKETSMART DATA
    def response_breakdown(self, save=False, medium=False, gtype='major', no_show=False):
        label = '_medium' if medium else '' if not medium else None
        if {'dns_email', 'dns_mail'}.issubset(set(self.rfm_data.columns)):
            email = self.rfm_data[(self.rfm_data['email on file'] == 'Y') & (self.rfm_data['dns_email'].isna())]\
                .groupby('rfm_group').donor_id.count().reindex(['Best', 'High', 'Medium', 'Low']).fillna(0)
            direct_mail = self.rfm_data[(self.rfm_data['physical address on file'] == 'Y') & (self.rfm_data['email on file'].isna()) & (self.rfm_data['dns_mail'].isna())]\
                .groupby('rfm_group').donor_id.count().reindex(['Best', 'High', 'Medium', 'Low']).fillna(0)
        else:
            email = self.rfm_data[(self.rfm_data['email on file'] == 'Y')].groupby(
                'rfm_group').donor_id.count().reindex(['Best', 'High', 'Medium', 'Low']).fillna(0)
            direct_mail = self.rfm_data[(self.rfm_data['physical address on file'] == 'Y') & (self.rfm_data['email on file'].isna())]\
                .groupby('rfm_group').donor_id.count().reindex(['Best', 'High', 'Medium', 'Low']).fillna(0)
        audience = {'email': email[0]+email[1]+email[2], 'direct_mail': direct_mail[0]+direct_mail[1]+direct_mail[2]} if medium\
            else {'email': email[0]+email[1], 'direct_mail': direct_mail[0]+direct_mail[1]}
        n_responses_breakdown = {'email': int(
            audience['email']*0.05), 'direct_mail': int(audience['direct_mail']*0.04)}
        n_responses = int(audience['email']*0.05 +
                          audience['direct_mail']*0.04)
        response_distribution = {'already_gave': int(n_responses*0.0388), 'immediate': int(n_responses*0.0451), 'deferred': int(n_responses*0.4857)} if gtype == 'major'\
            else {'already_gave': int(n_responses*0.0388), 'immediate': int(n_responses*0.0451), 'deferred': int(n_responses*0.4857), 'none': int(n_responses*0.4304)} if gtype == 'legacy'\
            else None
        labels = ['Somewhat Likely', 'Likely', 'Already given'] if gtype == 'major'\
            else ['No Interest', 'Deferred Interest', 'Immediate Interest', 'Gift Disclosed'] if gtype == 'legacy'\
            else None
        try:
            avg_major_gift = self.avg_mjr_gift
        except:
            avg_major_gift = 30000
        avg_legacy_gift = 78630
        projections = {'deferred': avg_major_gift*0.15*int(n_responses*0.4857), 'immediate': avg_major_gift*0.35*int(n_responses*0.0451), 'already_gave': avg_major_gift*int(n_responses*0.0388)} if gtype == 'major'\
            else {'none': avg_legacy_gift*0.05*int(n_responses*0.4304), 'deferred': avg_legacy_gift*0.15*int(n_responses*0.4857), 'immediate': avg_legacy_gift*0.35*int(n_responses*0.0451), 'already_gave': avg_legacy_gift*int(n_responses*0.0388)} if gtype == 'legacy'\
            else None
        total = sum(projections.values())
        max_y = max(response_distribution.values())
        intervals = [1, 2, 3, 4, 5, 10, 15, 20, 25, 30, 40,
                     50, 100, 150, 200, 250, 300, 400, 500, 750, 1000]
        interval_index = np.argmin([abs(i-(max_y/8)) for i in intervals])
        plt.figure()
        ax = plt.subplot()
        ax.grid(axis='x', alpha=0.5)
        ax.set_axisbelow(True)
        ax.set_yticks(range(len(labels)))
        ax.set_yticklabels(labels[::-1], fontproperties=prop)
        ax.spines['left'].set_color('#404041')
        ax.set_xticks(range(0, int(max_y)+10, intervals[interval_index]))
        ax.set_xticklabels([str(x) for x in range(
            0, int(max_y)+10, intervals[interval_index])], fontproperties=prop)
        ax.tick_params(axis='y', colors='#404041')
        ax.tick_params(axis='x', colors='#404041')
        plt.xlabel('Responses', fontproperties=prop)
        plt.barh(range(len(labels)), response_distribution.values(),
                 color='#00a7e0', height=0.6)
        sns.despine(top=True, bottom=True, right=True)
        for index, value in enumerate(response_distribution.values()):
            value = round(value, 1)
            if value >= max_y/8:
                plt.text(value-(max_y*(len(str(value)+f' ({round(100*value/n_responses,1)}%)')))/60,
                         index - 0.0375, str(value), color='w', fontproperties=prop_bold)
                plt.text(value-(max_y*(len(f' ({round(100*value/n_responses,1)}%)')))/60,
                         index - 0.0375, f' ({round(100*value/n_responses,1)}%)', color='w', fontproperties=prop)
            elif value < max_y/8:
                plt.text(value + max_y/400, index - 0.0375, str(value),
                         color='#404041', fontproperties=prop_bold)
                plt.text(value + len(str(value))*max_y/60 + max_y/400, index - 0.0375,
                         f' ({round(100*value/n_responses,1)}%)', color='#404041', fontproperties=prop)
            else:
                plt.text(value-(max_y*(len(str(value)+f' ({round(100*value/n_responses,1)}%)')))/60,
                         index - 0.0375, str(value)+f' ({round(100*value/n_responses,1)}%)', color='w', fontproperties=prop)
        if save:
            if gtype == 'major':
                plt.savefig(
                    f'temp plots/response_breakdown_major{label}.jpg', dpi=300, bbox_inches='tight', pad_inches=0)
            elif gtype == 'legacy':
                plt.savefig(
                    f'temp plots/response_breakdown_legacy{label}.jpg', dpi=300, bbox_inches='tight', pad_inches=0)
            plt.close()
        elif not save and not no_show:
            plt.show()
        return n_responses_breakdown, projections, total, audience

    def save_plots(self):  # SAVES ALL RELEVANT CHARTS/IMAGES TO A TEMPORARY FOLDER UNTIL THEY CAN BE UPLOADED TO GOOGLE DRIVE
        cd = os.getcwd()
        if os.path.exists(cd+'/temp plots'):
            for file in os.listdir(cd+'/temp plots/'):
                os.remove(cd+'/temp plots/'+str(file))
            os.rmdir(cd+'/temp plots')
        os.mkdir(cd+'/temp plots')
        if self.has_ages:
            self.age_distribution(save=True)
            self.age_distribution_compare(save=True)
            if 'legacy_member' in self.rfm_data.columns:
                self.age_distribution(legacy_only=True, save=True)
                self.age_distribution_compare(legacy_only=True, save=True)
        self.print_online_piechart(save=True)
        self.response_breakdown(save=True, medium=True, gtype='legacy')
        self.response_breakdown(save=True, medium=True, gtype='major')
        self.response_breakdown(save=True, medium=False, gtype='legacy')
        self.response_breakdown(save=True, medium=False, gtype='major')
        plt.close(fig='all')


class Yearly:  # CLASS OF OBJECT FROM DATAFRAME TO ALLOW FOR GENERATION OF ALL NECESSARY YEAR-BASED METRICS AND VISUALS

    def __init__(self, data):
        self.data = data
        self.years_labels = list(
            filter(lambda x: str(x).startswith('20'), data.columns))
        self.has_years = len(self.years_labels) > 2
        if self.has_years:
            self.years_ints = [int(i) for i in self.years_labels]
            self.max_yr_is_this_yr = date.today().year == max(self.years_ints)
            self.years_labels = self.years_labels[:-
                                                  1] if self.max_yr_is_this_yr else self.years_labels
        self.class_dict = {}

    # GENERATES BAR CHART SHOWING AMOUNT OF GIFTS BY YEAR
    def gifts_by_year(self, save=False):
        years = self.years_labels
        data = self.data[years].sum(axis=0)
        gift_sums = [i/1000000 for i in data]
        max_y = max(gift_sums)
        gift_intervals = [0.25, 0.5, 0.75, 1, 2, 2.5, 4, 5, 7.5, 10]
        interval_index = np.argmin([abs(i-(max_y/6)) for i in gift_intervals])
        gift_sum_labels = [
            '$'+str(i)+'M' for i in np.arange(0, int(1.3*max_y), gift_intervals[interval_index])]
        change = [str(round(((gift_sums[i+1]/gift_sums[i])-1) *
                            100, 1))+'%' for i in range(len(years)-1)]
        change_percent = ['+'+i if i[0] != '-' else i for i in change]
        plt.figure()
        plt.ylim(0, int(max_y*1.3))
        ax1 = plt.subplot(label='gifts by years')
        ax1.grid(axis='y', alpha=0.5)
        ax1.set_axisbelow(True)
        ax1.set_yticks(np.arange(0, int(1.3*max_y),
                                 gift_intervals[interval_index]))
        ax1.set_yticklabels(gift_sum_labels, fontproperties=prop)
        ax1.spines['bottom'].set_color('#404041')
        ax1.spines['left'].set_color('#404041')
        ax1.set_xticks(range(len(years)))
        ax1.set_xticklabels(years, fontproperties=prop)
        ax1.tick_params(axis='y', colors='#404041')
        ax1.tick_params(axis='x', colors='#404041')
        plt.ylabel('Gifts', fontproperties=prop, color='#404041')
        plt.rcParams['font.sans-serif'] = 'Arial'
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['font.size'] = 10
        sns.despine(top=True, right=True, left=True)
        plt.bar(years, gift_sums, color='#00a7e0')
        for index, value in enumerate(change_percent):
            if len(value) > 6:
                plt.text(index+0.575+(0.05*(7-len(years))), max_y /
                         18, value, color='w', fontproperties=prop)
            elif len(value) > 5:
                plt.text(index+0.65+(0.05*(7-len(years))), max_y /
                         18, value, color='w', fontproperties=prop)
            else:
                plt.text(index+0.7+(0.05*(7-len(years))), max_y /
                         18, value, color='w', fontproperties=prop)
        for index, value in enumerate(gift_sums):
            plt.text(index-0.35+(0.05*(7-len(years))), value+(max_y/80), '$' +
                     str(round(value, 1))+'M', color='#404041', fontproperties=prop)
        if save:
            plt.savefig('temp plots/gifts_by_year.jpg', dpi=300,
                        bbox_inches='tight', pad_inches=0)
            plt.close()
        else:
            plt.show()

    # GENERATES BAR CHART SHOWING NUMBER OF DONORS BY YEAR
    def donors_by_year(self, save=False):
        years = self.years_labels
        data = self.data[self.years_labels].count().to_list()
        n_donors = data
        max_y = max(n_donors)
        print(data)
        asdasd
        donor_intervals = [100, 250, 500, 1000, 2000, 2500, 4000,
                           5000, 7500, 10000, 15000, 20000, 25000, 50000, 100000]
        interval_index = np.argmin([abs(i-(max_y/6)) for i in donor_intervals])
        n_donor_labels = ['{:,}'.format(i) for i in range(
            0, int(1.2*max(n_donors)), donor_intervals[interval_index])]
        change = [str(round(((n_donors[i]/n_donors[i-1])-1)*100, 1)
                      )+'%' for i in range(1, len(years))]
        change_percent = ['+'+i if i[0] != '-' else i for i in change]
        plt.figure()
        plt.ylim(0, max_y*1.2)
        ax2 = plt.subplot(label='donors by years')
        ax2.grid(axis='y', alpha=0.5)
        ax2.set_axisbelow(True)
        ax2.set_yticks(np.arange(0, int(1.2*max(n_donors)),
                                 donor_intervals[interval_index]))
        ax2.set_yticklabels(n_donor_labels, fontproperties=prop)
        ax2.spines['bottom'].set_color('#404041')
        ax2.spines['left'].set_color('#404041')
        ax2.set_xticks(range(len(years)))
        ax2.set_xticklabels(years, fontproperties=prop)
        ax2.tick_params(axis='y', colors='#404041')
        ax2.tick_params(axis='x', colors='#404041')
        plt.ylabel('Donors', fontproperties=prop, color='#404041')
        plt.rcParams['font.sans-serif'] = 'Arial'
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['font.size'] = 10
        sns.despine(top=True, right=True, left=True)
        plt.bar(years, n_donors, color='#00a7e0')
        for index, value in enumerate(change_percent):
            if len(value) > 6:
                plt.text(index+0.575+(0.05*(7-len(years))), max_y /
                         18, value, color='w', fontproperties=prop)
            elif len(value) > 5:
                plt.text(index+0.65+(0.05*(7-len(years))), max_y /
                         18, value, color='w', fontproperties=prop)
            else:
                plt.text(index+0.7+(0.05*(7-len(years))), max_y /
                         18, value, color='w', fontproperties=prop)
        for index, value in enumerate(n_donors):
            if len('{:,}'.format(value)) >= 6:
                plt.text(index-0.325+(0.05*(7-len(years))), value+(max_y/80),
                         '{:,}'.format(value), color='#404041', fontproperties=prop)
            else:
                plt.text(index-0.25+(0.05*(7-len(years))), value+(max_y/80),
                         '{:,}'.format(value), color='#404041', fontproperties=prop)
        if save:
            plt.savefig('temp plots/donors_by_year.jpg', dpi=300,
                        bbox_inches='tight', pad_inches=0)
            plt.close()
        else:
            plt.show()

    # GENERATES TABLE SHOWING RETENTION RATES FOR EACH YEAR
    def donor_retention(self, save=False, no_show=False):
        from table_renderer import render_mpl_table
        years = self.years_labels
        raw_data = {year: self.data[self.data[year] > 0].donor_id.sort_values(
        ).to_numpy() for year, year in zip(years, years)}
        unique_ids = pd.Series(raw_data)
        retention_rate_floats = [round((len(np.intersect1d(unique_ids[i+1], unique_ids[i]))
                                        / len(unique_ids[i]))*100, 1) for i in range(len(unique_ids)-1)]
        retention_rates = [str(i)+'%' for i in retention_rate_floats]
        retention_dict = {years[i+1]: retention_rates[j]
                          for i, j in zip(range(len(unique_ids)-1), range(len(unique_ids)-1))}
        self.class_dict = retention_dict
        df = pd.DataFrame(data=retention_dict, index=[''])
        render_mpl_table(df, header_columns=0, col_width=1.25)
        if save:
            plt.savefig('temp plots/donor_retention.jpg', dpi=300,
                        bbox_inches='tight', pad_inches=0)
            plt.close()
        elif not no_show and not save:
            plt.show()
        return retention_rate_floats

    # GENERATES TABLE SHOWING AVERAGE DONATION FOR EACH YEAR
    def avg_donation(self, save=False):
        from table_renderer import render_mpl_table
        years = self.years_labels
        raw_data = {year: '${:,}'.format(int(
            self.data[self.data[year] > 0][year].mean())) for year, year in zip(years, years)}
        avg_donations = pd.Series(raw_data)
        df = pd.DataFrame(data=raw_data, index=[''])
        render_mpl_table(df, header_columns=0, col_width=1.25)
        if save:
            plt.savefig('temp plots/avg_donation.jpg', dpi=300,
                        bbox_inches='tight', pad_inches=0)
            plt.close()
        else:
            plt.show()

    # GENERATES CHART SHOWING RETENTION OF A SINGLE YEAR'S DONORS OVER THE NEXT 3 YEARS
    def single_class_retention(self, save=False, gift_size_cutoff=0):
        years = self.years_labels
        raw_data = {year: self.data[self.data[year] > gift_size_cutoff].donor_id.sort_values(
        ).to_numpy() for year, year in zip(years, years)}
        data = pd.Series(raw_data)
        unique_ids = data
        pct_retained_donors_4yrs_ago = [
            100*len(np.intersect1d(unique_ids[-4], unique_ids[i-4]))/len(unique_ids[-4]) for i in range(4)]
        # pct_retained_donors_4yrs_ago = [100*len(np.intersect1d(unique_ids[-3],unique_ids[i-3]))/len(unique_ids[-3]) for i in range(3)]
        y = pct_retained_donors_4yrs_ago
        year_labels = [years[-4], years[-3], years[-2], years[-1]]
        # year_labels = [years[-3],years[-2],years[-1]]
        max_y = max(y)
        pct_intervals = [1, 2, 2.5, 4, 5, 10, 12.5, 15, 20]
        interval_index = np.argmin([abs(i-(max_y/6)) for i in pct_intervals])
        pct_labels = [str(i)+'%' for i in range(0, int(max_y),
                                                pct_intervals[interval_index])]
        plt.figure()
        plt.ylim(0, 1.2*max_y)
        ax2 = plt.subplot(label='donors by years')
        ax2.grid(axis='y', alpha=0.5)
        ax2.set_axisbelow(True)
        ax2.set_yticks(np.arange(0, int(max_y), pct_intervals[interval_index]))
        ax2.set_yticklabels(pct_labels, fontproperties=prop)
        ax2.spines['bottom'].set_color('#404041')
        ax2.spines['left'].set_color('#404041')
        ax2.set_xticks(range(len(year_labels)))
        ax2.set_xticklabels(year_labels, fontproperties=prop)
        ax2.tick_params(axis='y', colors='#404041')
        ax2.tick_params(axis='x', colors='#404041')
        plt.ylabel('Donors', fontproperties=prop, color='#404041')
        plt.rcParams['font.sans-serif'] = 'Arial'
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['font.size'] = 10
        sns.despine(top=True, right=True, left=True)
        plt.title(f'%  Retained of {years[-3]} Donors',
                  fontproperties=prop, color='#404041')
        plt.bar(range(4), y, color='#00a7e0', width=0.5, label=years[-4])
        # plt.bar(range(3),y,color = '#00a7e0',width=0.5,label=years[-3])
        plt.plot(range(4), y, 'o-', color='#404041')
        # plt.plot(range(3),y,'o-',color = '#404041')
        width = ax2.patch.get_width()
        for index, value in enumerate(y):
            plt.text(index, value+(max_y/40), str(round(value, 1)) +
                     '%', color='#404041', fontproperties=prop)
        if save:
            plt.savefig('temp plots/single_class_retention.jpg',
                        dpi=300, bbox_inches='tight', pad_inches=0)
            plt.close()
        else:
            plt.show()

    # GENERATES CHART SHOWING RETENTION OF A MULTIPLE YEARS'S DONORS OVER THE NEXT 3 YEARS
    def multi_class_retention(self, save=False, gift_size_cutoff=0, single_only=False):
        years = self.years_labels
        raw_data = {year: self.data[self.data[year] > gift_size_cutoff].donor_id.sort_values(
        ).to_numpy() for year, year in zip(years, years)}
        data = pd.Series(raw_data)
        unique_ids = data
        pct_retained_donors_4yrs_ago = [
            100*len(np.intersect1d(unique_ids[-4], unique_ids[i-4]))/len(unique_ids[-4]) for i in range(4)]
        pct_retained_donors_5yrs_ago = [
            100*len(np.intersect1d(unique_ids[-5], unique_ids[i-5]))/len(unique_ids[-5]) for i in range(4)]
        pct_retained_donors_6yrs_ago = [
            100*len(np.intersect1d(unique_ids[-6], unique_ids[i-6]))/len(unique_ids[-6]) for i in range(4)]
        pct_retained_donors_7yrs_ago = [
            100*len(np.intersect1d(unique_ids[-7], unique_ids[i-7]))/len(unique_ids[-7]) for i in range(4)]
        y = pct_retained_donors_4yrs_ago
        y2 = pct_retained_donors_5yrs_ago
        y3 = pct_retained_donors_6yrs_ago
        y4 = pct_retained_donors_7yrs_ago
        # year_labels = [years[i-3] for i in range(3)]
        year_labels = [years[i-4] for i in range(4)]
        year_labels = ['Initial Year', 'After 1 year',
                       'After 2 years', 'After 3 years']
        max_y = max(y+y2+y3+y4)
        pct_intervals = [1, 2, 2.5, 3, 4, 5, 7.5, 10, 12.5, 15, 20]
        interval_index = np.argmin([abs(i-(max_y/6)) for i in pct_intervals])
        pct_labels = [str(i)+'%' for i in range(0, int(max_y),
                                                pct_intervals[interval_index])]
        plt.figure()
        plt.ylim((0, max_y))
        ax2 = plt.subplot(label='donors by years')
        ax2.grid(axis='y', alpha=0.5)
        ax2.set_axisbelow(True)
        ax2.set_yticks(np.arange(0, int(max_y), pct_intervals[interval_index]))
        ax2.set_yticklabels(pct_labels, fontproperties=prop)
        ax2.spines['bottom'].set_color('#404041')
        ax2.spines['left'].set_color('#404041')
        ax2.set_xticks(range(len(year_labels)))
        ax2.set_xticklabels(year_labels, fontproperties=prop)
        ax2.tick_params(axis='y', colors='#404041')
        ax2.tick_params(axis='x', colors='#404041')
        plt.ylabel('% Retained', fontproperties=prop, color='#404041')
        plt.rcParams['font.sans-serif'] = 'Arial'
        plt.rcParams['font.family'] = 'sans-serif'
        plt.rcParams['font.size'] = 10
        sns.despine(top=True, right=True, left=True)
        plt.title(
            f'%  Retained of {years[-7]}-{years[-4]} Donors', fontproperties=prop, color='#404041')
        # plt.title(f'%  Retained of {years[-7]}-{years[-3]} Donors',fontproperties=prop,color='#404041')
        y_avg = [np.mean([i, j, k, l]) for i, j, k, l in zip(y, y2, y3, y4)]
        if single_only:
            plt.bar(np.arange(4)+0.225, y, color='#00a7e0',
                    width=0.15, label=years[-4])
            plt.bar(np.arange(4)+0.075, y2, color='none', width=0.15)
            plt.bar(np.arange(4)-0.075, y3, color='none', width=0.15)
            plt.bar(np.arange(4)-0.225, y4, color='none', width=0.15)
        elif not single_only:
            plt.bar(np.arange(4)+0.225, y, color='#00a7e0',
                    width=0.15, label=years[-4])
            plt.plot(np.arange(4), y_avg, 'o-', color='#404041')
            plt.bar(np.arange(4)+0.075, y2, color='#ffce01',
                    width=0.15, label=years[-5])
            plt.bar(np.arange(4)-0.075, y3, color='#006d9e',
                    width=0.15, label=years[-6])
            plt.bar(np.arange(4)-0.225, y4, color='#004f6e',
                    width=0.15, label=years[-7])
            for index, value in enumerate(y_avg):
                text = plt.text(index+0.05, value+(max_y/30), str(round(value, 1)
                                                                  )+'%', color='#404041', fontproperties=prop)
                text.set_path_effects([path_effects.Stroke(
                    linewidth=2, foreground='white'), path_effects.Normal()])
        plt.legend(prop=prop, labelspacing=0.25,
                   framealpha=1.0, handlelength=0.8)
        if save:
            if single_only:
                plt.savefig('temp plots/multi_class_retention_single.jpg',
                            dpi=300, bbox_inches='tight', pad_inches=0)
            else:
                plt.savefig('temp plots/multi_class_retention.jpg',
                            dpi=300, bbox_inches='tight', pad_inches=0)
            plt.close()
        else:
            plt.show()

    # GENERATES BAR CHART SHOWING UNREALIZED NUMBER OF DONORS/GIFTS BY YEAR COMPARED TO A MARGINALLY BETTER RETETNION RATE
    def unrealized_potential(self, gift_floor=0, donors=False, gifts=False, save=False):
        years = self.years_labels[1:]
        gift_data = self.data[years].sum()

        def calc_under_4999():
            avg_gifts = [self.data[(self.data[year] > 0) & (
                self.data[year] < 5000)][year].mean() for year in years]
            raw_data = {year: self.data[(self.data[year] > 0) & (
                self.data[year] < 5000)].donor_id.sort_values().to_numpy() for year, year in zip(years, years)}
            unique_ids = pd.Series(raw_data)
            n_donors = [len(unique_ids[i]) for i in range(unique_ids.size)]
            n_new_donors = [len(np.setdiff1d(unique_ids[i+1], unique_ids[i]))
                            for i in range(len(unique_ids)-1)]
            retention_rates = [len(np.intersect1d(
                unique_ids[i+1], unique_ids[i]))/len(unique_ids[i]) for i in range(len(unique_ids)-1)]
            what_if_donors = [len(unique_ids[0])]
            for i in range(unique_ids.size-1):
                next_item = n_new_donors[i] + \
                    (what_if_donors[i]*(1.1*retention_rates[i]))
                what_if_donors.append(next_item)
            what_if_dif = [what_if_donors[i]-n_donors[i]
                           for i in range(len(n_donors))]
            what_if_gifts = [
                gift_data[i] + (what_if_dif[i]*avg_gifts[i]) for i in range(len(gift_data))]
            what_if_gifts_dif = [what_if_gifts[i] - gift_data[i]
                                 for i in range(len(gift_data))]
            unrealized_gifts = sum([what_if_gifts[i]-gift_data[i]
                                    for i in range(len(gift_data))])
            return [unrealized_gifts, np.array(what_if_dif[1:]), np.array(n_donors), years[0],
                    np.array(what_if_donors), np.array(what_if_gifts), np.array(what_if_gifts_dif)]

        def calc_over_4999():
            avg_gifts = [self.data[self.data[year] > 4999][year].mean()
                         for year in years]
            raw_data = {year: self.data[self.data[year] > 4999].donor_id.sort_values(
            ).to_numpy() for year, year in zip(years, years) if year != 2013}
            unique_ids = pd.Series(raw_data)

            n_donors = [len(unique_ids[i]) for i in range(unique_ids.size)]
            n_new_donors = [len(np.setdiff1d(unique_ids[i+1], unique_ids[i]))
                            for i in range(len(unique_ids)-1)]
            retention_rates = [len(np.intersect1d(
                unique_ids[i+1], unique_ids[i]))/len(unique_ids[i]) for i in range(len(unique_ids)-1)]
            what_if_donors = [len(unique_ids[0])]
            for i in range(unique_ids.size-1):
                next_item = n_new_donors[i] + \
                    (what_if_donors[i]*(1.1*retention_rates[i]))
                what_if_donors.append(next_item)
            what_if_dif = [what_if_donors[i]-n_donors[i]
                           for i in range(len(n_donors))]
            what_if_gifts = [
                gift_data[i] + (what_if_dif[i]*avg_gifts[i]) for i in range(len(gift_data))]
            what_if_gifts_dif = [what_if_gifts[i] - gift_data[i]
                                 for i in range(len(gift_data))]
            unrealized_gifts = sum([what_if_gifts[i]-gift_data[i]
                                    for i in range(len(gift_data))])
            return [unrealized_gifts, np.array(what_if_dif[1:]), np.array(n_donors), years[0],
                    np.array(what_if_donors), np.array(what_if_gifts), np.array(what_if_gifts_dif)]

        if gift_floor == 0:
            what_if_donors = calc_over_4999()[4]+calc_under_4999()[4]
            what_if_dif = calc_over_4999()[1]+calc_under_4999()[1]
            what_if_gifts = calc_over_4999()[5]+calc_under_4999()[6]
            n_donors = calc_over_4999()[2]+calc_under_4999()[2]
            what_if_gifts_dif = calc_over_4999()[6]+calc_under_4999()[6]
            result_dict = dict(unrealized_gifts=calc_under_4999()[0]+calc_over_4999()[0],
                               what_if_dif=what_if_dif,
                               n_donors=n_donors,
                               earliest_year=min([int(calc_under_4999()[3]), int(calc_over_4999()[3])]))

        elif gift_floor == 4999:
            what_if_donors = calc_over_4999()[4]
            what_if_dif = calc_over_4999()[1]
            what_if_gifts = calc_over_4999()[5]
            n_donors = calc_over_4999()[2]
            what_if_gifts_dif = calc_over_4999()[6]
            result_dict = dict(unrealized_gifts=calc_over_4999()[0],
                               what_if_dif=what_if_dif,
                               n_donors=n_donors,
                               earliest_year=int(calc_over_4999()[3]))
        if donors or save:
            max_y = max(what_if_donors)
            donor_intervals = [1, 2, 4, 5, 8, 10, 15, 20, 25, 30, 40, 50, 60, 80, 100, 250, 500, 1000, 2000,
                               2500, 4000, 5000, 7500, 10000, 15000, 20000, 25000, 50000, 100000]
            interval_index = np.argmin([abs(i-(max_y/6))
                                        for i in donor_intervals])
            n_donor_labels = ['{:,}'.format(i) for i in range(
                0, int(1.2*max(n_donors)), donor_intervals[interval_index])]
            plt.figure()
            plt.ylim(0, max_y*1.2)
            ax2 = plt.subplot(label='unrealized donors by years')
            ax2.grid(axis='y', alpha=0.5)
            ax2.set_axisbelow(True)
            ax2.set_yticks(np.arange(0, int(1.2*max(n_donors)),
                                     donor_intervals[interval_index]))
            ax2.set_yticklabels(n_donor_labels, fontproperties=prop)
            ax2.spines['bottom'].set_color('#404041')
            ax2.spines['left'].set_color('#404041')
            ax2.set_xticks(range(len(years)))
            ax2.set_xticklabels(years, fontproperties=prop)
            ax2.tick_params(axis='y', colors='#404041')
            ax2.tick_params(axis='x', colors='#404041')
            plt.ylabel('Donors', fontproperties=prop, color='#404041')
            sns.despine(top=True, right=True, left=True)
            plt.bar(years, what_if_donors, edgecolor='#00a7e0', color='none')
            plt.bar(years, n_donors, color='#00a7e0')
            for index, value in enumerate(n_donors):
                if value/max_y > 0.045:
                    if len('{:,}'.format(value)) >= 6:
                        plt.text(index-0.325+(0.05*(7-len(years))), value-(max_y/25),
                                 '{:,}'.format(value), color='w', fontproperties=prop)
                    else:
                        plt.text(index-0.25+(0.05*(7-len(years))), value-(max_y/25),
                                 '{:,}'.format(value), color='w', fontproperties=prop)
            for index, value in enumerate(what_if_dif):
                if len('{:,}'.format(int(value))) >= 5:
                    plt.text(index+1-0.325+(0.05*(7-len(years))), n_donors[index+1]+value+(
                        max_y/80), '{:,}'.format(int(value)), color='#404041', fontproperties=prop)
                else:
                    plt.text(index+0.8+(0.05*(7-len(years))), n_donors[index+1]+value+(
                        max_y/80), '{:,}'.format(int(value)), color='#404041', fontproperties=prop)
            if save:
                plt.savefig(
                    f'temp plots/unrealized_donors{gift_floor}.jpg', dpi=300, bbox_inches='tight', pad_inches=0)
                plt.close()
            else:
                plt.show()
        if gifts or save:
            what_if_gift_sums = [i/1000000 for i in what_if_gifts]
            what_if_gifts_dif_sums = [i/1000000 for i in what_if_gifts_dif]
            gift_sums = [i/1000000 for i in gift_data]
            max_y = max(what_if_gift_sums)
            gift_intervals = [0.25, 0.5, 0.75, 1, 2, 2.5, 4, 5, 7.5, 10]
            interval_index = np.argmin([abs(i-(max_y/6))
                                        for i in gift_intervals])
            gift_sum_labels = [
                '$'+str(i)+'M' for i in np.arange(0, int(1.2*max_y), gift_intervals[interval_index])]
            plt.figure()
            plt.ylim(0, int(max_y*1.2))
            ax1 = plt.subplot(label='unrealized gifts by years')
            ax1.grid(axis='y', alpha=0.5)
            ax1.set_axisbelow(True)
            ax1.set_yticks(np.arange(0, int(1.2*max_y),
                                     gift_intervals[interval_index]))
            ax1.set_yticklabels(gift_sum_labels, fontproperties=prop)
            ax1.spines['bottom'].set_color('#404041')
            ax1.spines['left'].set_color('#404041')
            ax1.set_xticks(range(len(years)))
            ax1.set_xticklabels(years, fontproperties=prop)
            ax1.tick_params(axis='y', colors='#404041')
            ax1.tick_params(axis='x', colors='#404041')
            plt.ylabel('Gifts', fontproperties=prop, color='#404041')
            sns.despine(top=True, right=True, left=True)
            plt.bar(years, what_if_gift_sums,
                    edgecolor='#00a7e0', color='none')
            plt.bar(years, gift_sums, color='#00a7e0')
            for index, value in enumerate(what_if_gift_sums[1:]):
                if value/max_y > 0.045:
                    plt.text(index+0.675+(0.05*(7-len(years))), value+(max_y/80), '$'+str(round(
                        what_if_gifts_dif_sums[index+1], 1))+'M', color='#404041', fontproperties=prop)
            for index, value in enumerate(gift_sums):
                if value/max_y > 0.045:
                    plt.text(index-0.325+(0.05*(7-len(years))), value-(max_y/25),
                             '$'+str(round(value, 1))+'M', color='w', fontproperties=prop)
            if save:
                plt.savefig(
                    f'temp plots/unrealized_gifts{gift_floor}.jpg', dpi=300, bbox_inches='tight', pad_inches=0)
                plt.close()
            else:
                plt.show()
        return result_dict

    # RETURNS SPECIFIC DATA ABOUT RETENTION OF $5K+ DONORS TO BE APPLIED TO A TEXT BOX
    def over5k_retention(self):
        years = self.years_labels[1:]
        raw_data = {year: self.data[self.data[year] >= 5000].donor_id.sort_values(
        ).to_numpy() for year, year in zip(years, years)}
        data = pd.Series(raw_data)
        unique_ids = data
        year4_retained_number = len(
            np.intersect1d(unique_ids[-1], unique_ids[-4]))
        # year4_retained_number = len(np.intersect1d(unique_ids[-1],unique_ids[-3]))
        year4_retained_rate = str(
            round((year4_retained_number/len(unique_ids[-4]))*100, 1))+'%'
        # year4_retained_rate = str(round((year4_retained_number/len(unique_ids[-3]))*100,1))+'%'
        retention_rates = [str(round((len(np.intersect1d(unique_ids[i+1], unique_ids[i]))
                                      / len(unique_ids[i]))*100, 1))+'%' for i in range(len(unique_ids)-1)]
        return [years[-1], years[-4], year4_retained_number, year4_retained_rate]
        # return [years[-1],years[-3],year4_retained_number,year4_retained_rate]

    def avg_major_gift(self):  # RETURNS THE AVERAGE MAJOR GIFT OF THE DATASET
        years = self.years_labels
        major_gifts = [self.data[self.data[year] >= 5000][year]
                       for year in years]
        avg_mjr_gift = sum([self.data[self.data[year] >= 5000][year].sum() for year in years])\
            / sum([len(self.data[self.data[year] >= 5000][year]) for year in years])
        return avg_mjr_gift

    # GENERATES BAR CHART SHOWING AMOUNT OF GIFTS BY YEAR BUT HIGHLIGHTING IN YELLOW THE AMOUNT THAT CAME FROM $5K+ DONORS
    def pareto_gifts(self, save=False):
        years = self.years_labels
        raw_best = {year: self.data[self.data[year] >= 5000]
                    [year].sum() for year, year in zip(years, years)}
        raw_else = {year: self.data[self.data[year] < 5000]
                    [year].sum() for year, year in zip(years, years)}
        gift_sum_best = pd.Series(raw_best)
        gift_sum_else = pd.Series(raw_else)
        gift_sum_best = [i/1000000 for i in gift_sum_best]
        gift_sum_else = [i/1000000 for i in gift_sum_else]
        gift_sum_total = [i+j for i, j in zip(gift_sum_best, gift_sum_else)]
        max_y = max(gift_sum_total)
        gift_intervals = [0.25, 0.5, 0.75, 1, 2, 2.5, 4, 5, 7.5, 10]
        interval_index = np.argmin([abs(i-(max_y/6)) for i in gift_intervals])
        gift_sum_labels = [
            '$'+str(i)+'M' for i in np.arange(0, int(1.2*max_y), gift_intervals[interval_index])]
        plt.figure()
        plt.ylim(0, max_y*1.2)
        ax4 = plt.subplot(label='pareto gifts')
        ax4.grid(axis='y', alpha=0.5)
        ax4.set_axisbelow(True)
        ax4.set_yticks(np.arange(0, int(1.2*max_y),
                                 gift_intervals[interval_index]))
        ax4.set_yticklabels(gift_sum_labels, fontproperties=prop)
        ax4.spines['bottom'].set_color('#404041')
        ax4.spines['left'].set_color('#404041')
        ax4.set_xticks(range(len(years)))
        ax4.set_xticklabels(years, fontproperties=prop)
        ax4.tick_params(axis='y', colors='#404041')
        ax4.tick_params(axis='x', colors='#404041')
        plt.ylabel('Gifts', fontproperties=prop, color='#404041')
        sns.despine(top=True, right=True, left=True)
        plt.bar(years, gift_sum_best, bottom=gift_sum_else,
                color='#ffce01', label='Sum of Gifts >$5000')
        plt.bar(years, gift_sum_else, color='#00a7e0',
                label='Sum of Gifts <$5000')
        for index, value in enumerate(gift_sum_best):
            if value/max_y > 0.05:
                plt.text(index-0.35+(0.05*(7-len(years))), gift_sum_else[index]+gift_sum_best[index]-(
                    max_y/20), '$'+str(round(value, 1))+'M', color='#404041', fontproperties=prop)
            else:
                plt.text(index-0.35+(0.05*(7-len(years))), gift_sum_else[index]+gift_sum_best[index]+(
                    max_y/80), '$'+str(round(value, 1))+'M', color='#404041', fontproperties=prop)
        for index, value in enumerate(gift_sum_else):
            if value/max_y > 0.035:
                plt.text(index-0.35+(0.05*(7-len(years))), value-(max_y/25),
                         '$'+str(round(value, 1))+'M', color='w', fontproperties=prop)
        plt.legend(prop=prop, labelspacing=0.25, framealpha=1.0,
                   handlelength=0.8, bbox_to_anchor=(1.05, 1.125))
        if save:
            plt.savefig('temp plots/pareto_gifts.jpg', dpi=300,
                        bbox_inches='tight', pad_inches=0)
            plt.close()
        else:
            plt.show()

    # GENERATES BAR CHART SHOWING NUMBER OF DONORS BY YEAR BUT HIGHLIGHTING IN YELLOW THOSE WHO ARE $5K+ DONORS
    def pareto_donors(self, save=False):
        years = self.years_labels
        raw_best = {year: self.data[self.data[year] >= 5000]
                    [year].count() for year, year in zip(years, years)}
        raw_else = {year: self.data[self.data[year] < 5000]
                    [year].count() for year, year in zip(years, years)}
        n_donors_best = pd.Series(raw_best)
        n_donors_else = pd.Series(raw_else)
        n_donors_total = [i+j for i, j in zip(n_donors_best, n_donors_else)]
        max_y = max(n_donors_total)
        donor_intervals = [100, 250, 500, 1000, 2000, 4000,
                           5000, 7500, 10000, 15000, 25000, 50000, 100000]
        interval_index = np.argmin([abs(i-(max_y/6)) for i in donor_intervals])
        n_donor_labels = ['{:,}'.format(i) for i in range(
            0, int(1.2*max_y), donor_intervals[interval_index])]
        plt.figure()
        plt.ylim(0, max_y*1.2)
        ax3 = plt.subplot(label='pareto donors')
        ax3.grid(axis='y', alpha=0.5)
        ax3.set_axisbelow(True)
        ax3.set_yticks(np.arange(0, int(1.2*max_y),
                                 donor_intervals[interval_index]))
        ax3.set_yticklabels(n_donor_labels, fontproperties=prop)
        ax3.spines['bottom'].set_color('#404041')
        ax3.spines['left'].set_color('#404041')
        ax3.set_xticks(range(len(years)))
        ax3.set_xticklabels(years, fontproperties=prop)
        ax3.tick_params(axis='y', colors='#404041')
        ax3.tick_params(axis='x', colors='#404041')
        plt.ylabel('Donors', fontproperties=prop, color='#404041')
        sns.despine(top=True, right=True, left=True)
        plt.bar(years, n_donors_best, bottom=n_donors_else,
                color='#ffce01', label='>$5000 Donors')
        plt.bar(years, n_donors_else, color='#00a7e0', label='<$5000 Donors')
        for index, value in enumerate(n_donors_best):
            if len(str(value)) >= 3:
                plt.text(index-0.25+(0.05*(7-len(years))), n_donors_total[index]+(
                    max_y/80), '{:,}'.format(value), color='#404041', fontproperties=prop)
            elif len(str(value)) >= 5:
                plt.text(index-0.35+(0.05*(7-len(years))), n_donors_total[index]+(
                    max_y/80), '{:,}'.format(value), color='#404041', fontproperties=prop)
            else:
                plt.text(index-0.15+(0.05*(7-len(years))), n_donors_total[index]+(
                    max_y/80), '{:,}'.format(value), color='#404041', fontproperties=prop)
        for index, value in enumerate(n_donors_else):
            if len(str(value)) >= 5:
                plt.text(index-0.325+(0.05*(7-len(years))), value-(max_y/20),
                         '{:,}'.format(value), color='w', fontproperties=prop)
            elif len(str(value)) >= 3 and len(str(value)) < 5:
                plt.text(index-0.25+(0.05*(7-len(years))), value-(max_y/20),
                         '{:,}'.format(value), color='w', fontproperties=prop)
            else:
                plt.text(index-0.15+(0.05*(7-len(years))), value-(max_y/20),
                         '{:,}'.format(value), color='w', fontproperties=prop)
        plt.legend(prop=prop, labelspacing=0.25, framealpha=1.0,
                   handlelength=0.8, bbox_to_anchor=(1.05, 1.125))
        if save:
            plt.savefig('temp plots/pareto_donors.jpg', dpi=300,
                        bbox_inches='tight', pad_inches=0)
            plt.close()
        else:
            plt.show()

    def save_plots(self):  # SAVES ALL RELEVANT CHARTS/IMAGES TO A TEMPORARY FOLDER UNTIL THEY CAN BE UPLOADED TO GOOGLE DRIVE
        cd = os.getcwd()
        if os.path.exists(cd+'/temp plots'):
            pass
        else:
            os.mkdir(dir+'/temp plots')
        if self.has_years:
            self.donors_by_year(save=True)
            self.gifts_by_year(save=True)
            self.donor_retention(save=True)
            self.avg_donation(save=True)
            self.pareto_donors(save=True)
            self.pareto_gifts(save=True)
            self.unrealized_potential(gift_floor=0, save=True)
            self.unrealized_potential(gift_floor=4999, save=True)
            self.single_class_retention(save=True)
            if len(self.years_labels) >= 7:
                self.multi_class_retention(save=True)
                self.multi_class_retention(save=True, single_only=True)
        plt.close(fig='all')


if __name__ == '__main__':
    rfm_obj = RFM(df)
    yoy_obj = Yearly(df)
    yoy_obj.donors_by_year(save=True)

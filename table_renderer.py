import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import six
from matplotlib import font_manager as fm


prop = fm.FontProperties(fname='fonts1/HelveticaNeueLt.ttf')

def render_mpl_table(data, col_width=2.0, row_height=0.4, font_size=14,
                     header_color='#00a7e0', row_colors=['#f1f1f2', '#f3f3f3'], edge_color='w',
                     bbox=[0, 0, 1, 1], header_columns=0,
                     ax=None, **kwargs):
    if ax is None:
        size = (np.array(data.shape[::-1]) + np.array([0, 1])) * np.array([col_width, row_height])
        fig, ax = plt.subplots(figsize=size)
        ax.axis('off')

    mpl_table = ax.table(cellText=data.values, bbox=bbox, colLabels=data.columns, **kwargs)

    mpl_table.auto_set_font_size(True)
    mpl_table.set_fontsize(font_size)

    for k, cell in six.iteritems(mpl_table._cells):
        cell.set_edgecolor(edge_color)
        if k[0] == 0 or k[1] < header_columns:
            cell.set_text_props(color='w',fontproperties=prop,size=font_size)
            cell.set_facecolor(header_color)
        else:
            cell.set_text_props(color='#404041',fontproperties=prop,size=font_size)
            cell.set_facecolor(row_colors[k[0]%len(row_colors) ])
    return ax
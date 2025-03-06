import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import streamlit as st
import matplotlib.font_manager
from streamlit_option_menu import option_menu


def get_available_fonttypeface():

    # Get a list of all available font names
    return sorted(set(f.name for f in matplotlib.font_manager.fontManager.ttflist))

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))


def add_bar_slide(df, prs, grouping):

    if grouping == 'Clustered':
        grouptype = XL_CHART_TYPE.COLUMN_CLUSTERED
        layoutname = 'BarChartClustered'
    if grouping == 'Stacked':
        grouptype = XL_CHART_TYPE.COLUMN_STACKED
        layoutname = 'BarChartStacked'

    # get the index number of the Chart Placeholder from the slide named BarChart
    for layout in prs.slide_layouts:
        if layout.name == layoutname:
            slide = prs.slides.add_slide(layout)
            for placeholder in layout.placeholders:
                if 'Chart Placeholder' in placeholder.name:
                    idx = placeholder.placeholder_format.idx

    # Define chart data
    chart_data = CategoryChartData()
    chart_data.categories = df[df.columns.tolist()[0]]
    
    df_headers = df.columns.tolist()[1:]
    for df_header in df_headers:
        chart_data.add_series(df_header, df[df_header])
    
    # Insert chart to the placeholder
    chart_placeholder = slide.placeholders[idx]
    chart_frame = chart_placeholder.insert_chart(grouptype, chart_data)
    _chart = chart_frame.chart



    # -------------- chart details -------------------
    # chart title
    available_fonts = sorted(set(f.name for f in matplotlib.font_manager.fontManager.ttflist)) 

    cht_title_dict = {}
    with st.expander('CHART TITLE'):
        col_title1, col_title2, col_title3, col_title4 = st.columns([0.4, 0.3, 0.15, 0.15])
        with col_title1:
            cht_title = st.text_input('Chart Title', key='cht_title', placeholder='Input Chart Title Here')
            if cht_title in [None, '']:
                cht_title_dict['text'] = 'No Chart Title'
            else:
                cht_title_dict['text'] = cht_title
        with col_title2:
            cht_title_dict['font_name'] = st.selectbox('Font Name', options=available_fonts, key='cht_title_font_name')
        with col_title3:
            cht_title_dict['font_size'] = st.number_input('Font Size', min_value=1, max_value=100, value=10, step=1, key='cht_title_font_size')
        with col_title4:
            cht_title_font_color = st.color_picker('Font Color', key='cht_title_font_color')
            rgb_color = hex_to_rgb(cht_title_font_color)
            cht_title_dict['font_color'] = rgb_color                

    # category axis
    cht_category_axis_dict = {}
    with st.expander('CATEGORY AXIS'):
        col_cat1, col_cat2, col_cat3, col_cat4 = st.columns([0.4, 0.3, 0.15, 0.15])
        with col_cat1:
            cat_label = st.text_input('Category Label', key='cat_label')
            if cat_label in [None, '']:
                cht_category_axis_dict['text'] = df.columns.tolist()[0]
            else:   
                cht_category_axis_dict['text'] = cat_label
        with col_cat2:
            cht_category_axis_dict['font_name'] = st.selectbox('Font Name', options=available_fonts, key='cht_category_font_name')
        with col_cat3:
            cht_category_axis_dict['font_size'] = st.number_input('Font Size', min_value=1, max_value=100, value=10, step=1, key='cht_category_font_size')
        with col_cat4:
                cht_cat_axis_ftcolor = st.color_picker('Font Color', key='cht_category_font_color')
                rgb_color = hex_to_rgb(cht_cat_axis_ftcolor)
                cht_category_axis_dict['font_color'] = rgb_color
    
    

    chart_details(df, _chart, cht_title_dict, cht_category_axis_dict)

    return


def chart_details(df, _chart, cht_title_dict, cht_category_axis_dict):

    # chart title
    _chart.has_title = True
    _chart.chart_title.text_frame.text = cht_title_dict['text']
    try:
        _chart.chart_title.text_frame.paragraphs[0].font.size = Pt(cht_title_dict['font_size'])
        _chart.chart_title.text_frame.paragraphs[0].font.name = cht_title_dict['font_name']
        _chart.chart_title.text_frame.paragraphs[0].font.color.rgb = RGBColor(*cht_title_dict['font_color'])
    except:
        pass
    
    # category axis
    _chart.category_axis.axis_title.text_frame.text = df.columns.tolist()[0]
    try:
        _chart.category_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(cht_category_axis_dict['font_size'])
        _chart.category_axis.axis_title.text_frame.paragraphs[0].font.name = cht_category_axis_dict['font_name']
        _chart.category_axis.axis_title.text_frame.paragraphs[0].font.color.rgb = RGBColor(*cht_category_axis_dict['font_color'])
    except:
        pass
    
    # if value_axis == True:
    #     # value axis
    #     _chart.value_axis.axis_title.text_frame.text = "Count"
    #     try:
    #         _chart.value_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(18)
    #         _chart.value_axis.axis_title.text_frame.paragraphs[0].font.name = 'Arial'
    #     except:
    #         pass
    
    # # legend
    # _chart.has_legend = True
    # _chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    # _chart.legend.include_in_layout = False
    # _chart.legend.font.size = Pt(8)
    # _chart.legend.font.bold = True
    # _chart.legend.font.name = 'Arial'

    # # data labels
    # for series in _chart.series:
    #     series.has_data_labels = True  # Enable data labels
    #     series.data_labels.show_value = True  # Show values on the labels
    #     # series.data_labels.show_category_name = True
    #     series.data_labels.font.size = Pt(8)
    #     series.data_labels.font.bold = True
    #     series.data_labels.font.name = 'Arial'
    #     # series.data_labels.show_series_name = False
    #     # series.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    
    # # data table
    # _chart.has_data_table = True
    # # _chart.data_table.has_border_horizontal = True  # Add horizontal borders
    # # _chart.data_table.has_border_vertical = True  # Add vertical borders
    # # _chart.data_table.has_border_outline = True  # Add outline border


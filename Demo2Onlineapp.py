import streamlit as st
import pandas as pd
import base64
import io
import numpy as np
import re
from PIL import Image
import matplotlib.pyplot as plt
# import spacy
import logging
import warnings
from nltk.corpus import stopwords
import nltk
import os

# Load data function
def load_data(file):
    if file:
        data = pd.read_excel(file)
        return data
    return None

# Data preprocessing function (You can include your data preprocessing here)

# Function to create separate Excel sheets by Entity
def create_entity_sheets(data, writer):
    # Define a format with text wrap
    wrap_format = writer.book.add_format({'text_wrap': True})

    for Entity in data['Entity'].unique():
        entity_df = data[data['Entity'] == Entity]
        entity_df.to_excel(writer, sheet_name=Entity, index=False)
        worksheet = writer.sheets[Entity]
        worksheet.set_column(1, 4, 48, cell_format=wrap_format)
        # Calculate column widths based on the maximum content length in each column except columns 1 to 4
        max_col_widths = [
            max(len(str(value)) for value in entity_df[column])
            for column in entity_df.columns[5:]  # Exclude columns 1 to 4
        ]

        # Set the column widths dynamically for columns 5 onwards
        for col_num, max_width in enumerate(max_col_widths):
            worksheet.set_column(col_num + 5, col_num + 5, max_width + 2)  # Adding extra padding for readability       

# Streamlit app with a sidebar layout
st.set_page_config(layout="wide")

# Custom CSS for title bar position
title_bar_style = """
    <style>
        .title h1 {
            margin-top: -10px; /* Adjust this value to move the title bar up or down */
        }
    </style>
"""

st.markdown(title_bar_style, unsafe_allow_html=True)

st.title("Meltwater Data Insights Dashboard")

# Sidebar for file upload and download options
st.sidebar.title("Upload a file for tables")

# File Upload Section
file = st.sidebar.file_uploader("Upload Data File (Excel or CSV)", type=["xlsx", "csv"])

if file:
    st.sidebar.write("File Uploaded Successfully!")

    # Load data
    data = load_data(file)

    if data is not None:
        # Data Preview Section (optional)
        # st.write("## Data Preview")
        # st.write(data)

        # Data preprocessing
        data.drop(columns=data.columns[10:], axis=1, inplace=True)
        data = data.rename({'Influencer': 'Journalist'}, axis=1)
        data.drop_duplicates(subset=['Date', 'Entity', 'Headline', 'Publication Name'], keep='first', inplace=True)
        finaldata = data
        finaldata['Date'] = pd.to_datetime(finaldata['Date']).dt.normalize()

        # Share of Voice (SOV) Calculation
        En_sov = pd.crosstab(finaldata['Entity'], columns='News Count', values=finaldata['Entity'], aggfunc='count').round(0)
        En_sov.sort_values('News Count', ascending=False)
        En_sov['% '] = ((En_sov['News Count'] / En_sov['News Count'].sum()) * 100).round(2)
        Sov_table = En_sov.sort_values(by='News Count', ascending=False)
        Sov_table.loc['Total'] = Sov_table.sum(numeric_only=True, axis=0)
        Entity_SOV1 = Sov_table.round()

        # Additional DataFrames
        sov_dt = pd.crosstab((finaldata['Date'].dt.to_period('M')), finaldata['Entity'], margins=True, margins_name='Total')
        pub_table = pd.crosstab(finaldata['Publication Name'], finaldata['Entity'])
        pub_table['Total'] = pub_table.sum(axis=1)
        pubs_table = pub_table.sort_values('Total', ascending=False).round()
        pubs_table.loc['GrandTotal'] = pubs_table.sum(numeric_only=True, axis=0)

        PP = pd.crosstab(finaldata['Publication Name'], finaldata['Publication Type'])
        PP['Total'] = PP.sum(axis=1)
        PP_table = PP.sort_values('Total', ascending=False).round()
        PP_table.loc['GrandTotal'] = PP_table.sum(numeric_only=True, axis=0)

        PT_Entity = pd.crosstab(finaldata['Publication Type'], finaldata['Entity'])
        PT_Entity['Total'] = PT_Entity.sum(axis=1)
        PType_Entity = PT_Entity.sort_values('Total', ascending=False).round()
        PType_Entity.loc['GrandTotal'] = PType_Entity.sum(numeric_only=True, axis=0)

        # Journalist Table
        finaldata['Journalist'] = finaldata['Journalist'].str.split(',')
        finaldata = finaldata.explode('Journalist')
        jr_tab = pd.crosstab(finaldata['Journalist'], finaldata['Entity'])
        jr_tab = jr_tab.reset_index(level=0)
        newdata = finaldata[['Journalist', 'Publication Name']]
        Journalist_Table = pd.merge(jr_tab, newdata, how='inner', left_on=['Journalist'], right_on=['Journalist'])
        Journalist_Table.drop_duplicates(subset=['Journalist'], keep='first', inplace=True)
        valid_columns = Journalist_Table.select_dtypes(include='number').columns
        Journalist_Table['Total'] = Journalist_Table[valid_columns].sum(axis=1)
        Journalist_Table = Journalist_Table.sort_values('Total', ascending=False).round()
        Jour_table = Journalist_Table.reset_index(drop=True)
        Jour_table.loc['GrandTotal'] = Jour_table.sum(numeric_only=True, axis=0)
        Jour_table.insert(1, 'Publication Name', Jour_table.pop('Publication Name'))

        # Function to classify news exclusivity and topic
        def classify_exclusivity(row):
            entity_name = row['Entity']
            if entity_name.lower() in row['Headline'].lower():
                return "Exclusive"
            else:
                return "Not Exclusive"

        finaldata['Exclusivity'] = finaldata.apply(classify_exclusivity, axis=1)

        # Define a dictionary of keywords for each entity
        entity_keywords = {
            'Amazon': ['Amazon', 'Amazons', 'amazon'],
            # Add other entities and keywords here
        }

        def qualify_entity(row):
            entity_name = row['Entity']
            text = row['Headline']
            if entity_name in entity_keywords:
                keywords = entity_keywords[entity_name]
                if any(keyword in text for keyword in keywords):
                    return "Qualified"
            return "Not Qualified"

        finaldata['Qualification'] = finaldata.apply(qualify_entity, axis=1)

        # Topic classification
        topic_mapping = {
            'Merger': ['merger', 'merges'],
            'Acquire': ['acquire', 'acquisition', 'acquires'],
            'Partnership': ['partnership', 'tie-up'],
            'Business Strategy': ['launch', 'campaign', 'IPO', 'sales'],
            'Investment and Funding': ['invest', 'funding'],
            'Employee Engagement': ['layoff', 'hiring'],
            'Financial Performance': ['profit', 'loss', 'revenue'],
            'Business Expansion': ['expansion', 'opens'],
            'Leadership': ['ceo'],
            'Stock Related': ['stock', 'shares'],
            'Awards & Recognition': ['award'],
            'Legal & Regulatory': ['penalty', 'scam'],
        }

        def classify_topic(headline):
            for topic, words in topic_mapping.items():
                if any(word in headline.lower() for word in words):
                    return topic
            return 'Other'

        finaldata['Topic'] = finaldata['Headline'].apply(classify_topic)

        dfs = [Entity_SOV1, sov_dt, pubs_table, Jour_table, PType_Entity, PP_table]
        comments = ['SOV Table', 'Month-on-Month Table', 'Publication Table', 'Journalist Table', 'Pub Type and Pub Name Table', 'PubType Entity Table']

        # Sidebar for download options
        st.sidebar.write("## Download Options")
        download_formats = st.sidebar.selectbox("Select format:", ["Excel", "CSV", "Excel (Entity Sheets)"])

        if st.sidebar.button("Download Data"):
            if download_formats == "Excel":
                # Download all DataFrames as a single Excel file
                excel_io = io.BytesIO()
                with pd.ExcelWriter(excel_io, engine="xlsxwriter") as writer:
                    for df, comment in zip(dfs, comments):
                        df.to_excel(writer, sheet_name=comment, index=False)
                excel_io.seek(0)
                b64_data = base64.b64encode(excel_io.read()).decode()
                href_data = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_data}" download="data.xlsx">Download Excel</a>'
                st.sidebar.markdown(href_data, unsafe_allow_html=True)

            elif download_formats == "CSV":
                # Download all DataFrames as CSV
                csv_io = io.StringIO()
                for df in dfs:
                    df.to_csv(csv_io, index=False)
                csv_io.seek(0)
                b64_data = base64.b64encode(csv_io.read().encode()).decode()
                href_data = f'<a href="data:text/csv;base64,{b64_data}" download="data.csv">Download CSV</a>'
                st.sidebar.markdown(href_data, unsafe_allow_html=True)

            elif download_formats == "Excel (Entity Sheets)":
                # Download DataFrames as Excel with separate sheets by entity
                excel_io = io.BytesIO()
                with pd.ExcelWriter(excel_io, engine="xlsxwriter") as writer:
                    create_entity_sheets(finaldata, writer)
                excel_io.seek(0)
                b64_data = base64.b64encode(excel_io.read()).decode()
                href_data = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_data}" download="entity_sheets.xlsx">Download Entity Sheets</a>'
                st.sidebar.markdown(href_data, unsafe_allow_html=True)

else:
    st.sidebar.write("No file uploaded yet.")

import streamlit as st
import pandas as pd
import base64
import io
# import streamlit as st
# import pandas as pd
import numpy as np
# import base64
import re
from wordcloud import WordCloud
from PIL import Image
# from fuzzywuzzy import fuzz
import matplotlib.pyplot as plt
# import gensim
import spacy
# import pyLDAvis.gensim_models
# from gensim.utils import simple_preprocess
# from gensim.models import CoherenceModel
from pprint import pprint
import logging
import warnings
from nltk.corpus import stopwords
# import gensim.corpora as corpora
from io import BytesIO
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

    for Entity in finaldata['Entity'].unique():
        entity_df = finaldata[finaldata['Entity'] == Entity]
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
            
            
# Function to save multiple DataFrames in a single Excel sheet
def multiple_dfs(df_list, sheets, file_name, spaces, comments):
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    row = 2
    for dataframe, comment in zip(df_list, comments):
        pd.Series(comment).to_excel(writer, sheet_name=sheets, startrow=row,
                                    startcol=1, index=False, header=False)
        dataframe.to_excel(writer, sheet_name=sheets, startrow=row + 1, startcol=0)
        row = row + len(dataframe.index) + spaces + 2
    writer.close()
     
    
def top_10_dfs(df_list, file_name, comments, top_11_flags):
    writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
    row = 2
    for dataframe, comment, top_11_flag in zip(df_list, comments, top_11_flags):
        if top_11_flag:
            top_df = dataframe.head(50)  # Select the top 11 rows for specific DataFrames
        else:
            top_df = dataframe  # Leave other DataFrames unchanged

        top_df.to_excel(writer, sheet_name="Top 10 Data", startrow=row, index=True)
        row += len(top_df) + 2  # Move the starting row down by len(top_df) + 2 rows

    # Create a "Report" sheet with all the DataFrames
    for dataframe, comment in zip(df_list, comments):
        dataframe.to_excel(writer, sheet_name="Report", startrow=row, index=True, header=True)
        row += len(dataframe) + 2  # Move the starting row down by len(dataframe) + 2 rows

    writer.close()

    
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

# Modify the paths according to your specific directory
download_path = r"C:\Users\akshay.annaldasula"

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
        # Data Preview Section
#         st.write("## Data Preview")
#         st.write(data)

        # Entity SOV Section
        # st.sidebar.write("## Entity Share of Voice (SOV)")
        # Include your Entity SOV code here

        # Data preprocessing
        data.drop(columns=data.columns[10:], axis=1, inplace=True)
        data = data.rename({'Influencer': 'Journalist'}, axis=1)
        data.drop_duplicates(subset=['Date', 'Entity', 'Headline', 'Publication Name'], keep='first', inplace=True, ignore_index=True)
        data.drop_duplicates(subset=['Date', 'Entity', 'Opening Text', 'Publication Name'], keep='first', inplace=True, ignore_index=True)
        data.drop_duplicates(subset=['Date', 'Entity', 'Hit Sentence', 'Publication Name'], keep='first', inplace=True, ignore_index=True)
        finaldata = data
        En_sov = pd.crosstab(finaldata['Entity'], columns='News Count', values=finaldata['Entity'], aggfunc='count').round(0)
        En_sov.sort_values('News Count', ascending=False)
        En_sov['% '] = ((En_sov['News Count'] / En_sov['News Count'].sum()) * 100).round(2)
        Sov_table = En_sov.sort_values(by='News Count', ascending=False)
        Sov_table.loc['Total'] = Sov_table.sum(numeric_only=True, axis=0)
        Entity_SOV1 = Sov_table.round()

        # st.sidebar.write(Entity_SOV1)

        finaldata['Date'] = pd.to_datetime(finaldata['Date']).dt.normalize()
        sov_dt = pd.crosstab((finaldata['Date'].dt.to_period('M')),finaldata['Entity'],margins = True ,margins_name='Total')

        pub_table = pd.crosstab(finaldata['Publication Name'],finaldata['Entity'])
        pub_table['Total']= pub_table.sum(axis=1)
        pubs_table=pub_table.sort_values('Total',ascending=False).round()
        pubs_table.loc['GrandTotal']= pubs_table.sum(numeric_only=True,axis=0)

        PP = pd.crosstab(finaldata['Publication Name'],finaldata['Publication Type'])
        PP['Total']= PP.sum(axis=1)
        PP_table=PP.sort_values('Total',ascending=False).round()
        PP_table.loc['GrandTotal']= PP_table.sum(numeric_only=True,axis=0)

        PT_Entity = pd.crosstab(finaldata['Publication Type'],finaldata['Entity'])
        PT_Entity['Total']= PT_Entity.sum(axis=1)
        PType_Entity=PT_Entity.sort_values('Total',ascending=False).round()
        PType_Entity.loc['GrandTotal']= PType_Entity.sum(numeric_only=True,axis=0)

        ppe = pd.crosstab(columns=finaldata['Entity'],index=[finaldata["Publication Type"],finaldata["Publication Name"]],margins=True,margins_name='Total')
        ppe1 = ppe.reset_index()
        ppe1.set_index("Publication Type", inplace = True)

        finaldata['Journalist']=finaldata['Journalist'].str.split(',')
        finaldata = finaldata.explode('Journalist')
        jr_tab=pd.crosstab(finaldata['Journalist'],finaldata['Entity'])
        jr_tab = jr_tab.reset_index(level=0)
        newdata = finaldata[['Journalist','Publication Name']]
        Journalist_Table = pd.merge(jr_tab, newdata, how='inner',
                  left_on=['Journalist'],
                  right_on=['Journalist'])

        Journalist_Table.drop_duplicates(subset=['Journalist'], keep='first', inplace=True, ignore_index=True)
        valid_columns = Journalist_Table.select_dtypes(include='number').columns
        Journalist_Table['Total'] = Journalist_Table[valid_columns].sum(axis=1)

        Journalist_Table.sort_values('Total',ascending=False)
        Jour_table=Journalist_Table.sort_values('Total',ascending=False).round()
        bn_row = Jour_table.loc[Jour_table['Journalist'] == 'Bureau News']
        Jour_table = Jour_table[Jour_table['Journalist'] != 'Bureau News']
        Jour_table = pd.concat([Jour_table, bn_row], ignore_index=True)
        Jour_table.loc['GrandTotal'] = Jour_table.sum(numeric_only=True, axis=0)
        Jour_table.insert(1, 'Publication Name', Jour_table.pop('Publication Name'))
        
        # Remove square brackets and single quotes from the 'Journalist' column
        #data['Journalist'] = data['Journalist'].str.strip("[]'")
        
        # Remove square brackets and single quotes from the 'Journalist' column
        data['Journalist'] = data['Journalist'].str.replace(r"^\['(.+)'\]$", r"\1", regex=True)
        
        # Define a function to classify news as "Exclusive" or "Not Exclusive" for the current entity
        def classify_exclusivity(row):
                
            entity_name = finaldata['Entity'].iloc[0]  # Get the entity name for the current sheet
            # Check if the entity name is mentioned in either 'Headline' or 'Similar_Headline'
            if entity_name.lower() in row['Headline'].lower() or entity_name.lower() in row['Headline'].lower():
                
                return "Exclusive"
            else:
                
                return "Not Exclusive"
                    
        # Apply the classify_exclusivity function to each row in the current entity's data
        finaldata['Exclusivity'] = finaldata.apply(classify_exclusivity, axis=1) 
        
        # Define a dictionary of keywords for each entity
        entity_keywords = {
                        'Amazon': ['Amazon','Amazons','amazon'],
#                           'LTTS': ['LTTS', 'ltts'],
#                           'KPIT': ['KPIT', 'kpit'],
#                          'Cyient': ['Cyient', 'cyient'], 
            }
            
        # Define a function to qualify entity based on keyword matching
        def qualify_entity(row):    
            
            entity_name = row['Entity']
            text = row['Headline']   
                
            if entity_name in entity_keywords:
                keywords = entity_keywords[entity_name]
                # Check if at least one keyword appears in the text
                if any(keyword in text for keyword in keywords):
                    
                    return "Qualified"
                
            return "Not Qualified"
            
        # Apply the qualify_entity function to each row in the current entity's data
        finaldata['Qualification'] = finaldata.apply(qualify_entity, axis=1)
        
        # Define a dictionary to map predefined words to topics
        topic_mapping = {
              'Merger': ['merger', 'merges'],
                
              'Acquire': ['acquire', 'acquisition', 'acquires'],
                
              'Partnership': ['partnership', 'tieup', 'tie-up','mou','ties up','ties-up','joint venture'],
                
               'Business Strategy': ['launch', 'launches', 'launched', 'announces','announced', 'announcement','IPO','campaign','launch','launches','ipo','sales','sells','introduces','announces','introduce','introduced','unveil',
                                    'unveils','unveiled','rebrands','changes name','bags','lays foundation','hikes','revises','brand ambassador','enters','ambassador','signs','onboards','stake','stakes','to induct','forays','deal'],
                
               'Investment and Funding': ['invests', 'investment','invested','funding', 'raises','invest','secures'],
                
              'Employee Engagement': ['layoff', 'lay-off', 'laid off', 'hire', 'hiring','hired','appointment','re-appoints','reappoints','steps down','resigns','resigned','new chairman','new ceo','layoffs','lay offs'],
                
              'Financial Performence': ['quarterly results', 'profit', 'losses', 'revenue','q1','q2','q3','q4'],
                
               'Business Expansion': ['expansion', 'expands', 'inaugration', 'inaugrates','to open','opens','setup','set up','to expand','inaugurates'], 
                
               'Leadership': ['in conversation', 'speaking to', 'speaking with','ceo'], 
                
               'Stock Related': ['buy', 'target', 'stock','shares' ,'stocks','trade spotlight','short call','nse'], 
                
                'Awards & Recognition': ['award', 'awards'],
                
                'Legal & Regulatory': ['penalty', 'fraud','scam','illegal'],
            
            'Sale - Offers - Discounts' : ['sale','offers','discount','discounts','discounted']
        }
            
        # Define a function to classify headlines into topics
        def classify_topic(headline):
            
            lowercase_headline = headline.lower()
            for topic, words in topic_mapping.items():
                for word in words:
                    if word in lowercase_headline:
                        return topic
            return 'Other'  # If none of the predefined words are found, assign 'Other'
            
                      
        # Apply the classify_topic function to each row in the dataframe
        finaldata['Topic'] = finaldata['Headline'].apply(classify_topic)
        
        


        dfs = [Entity_SOV1, sov_dt, pubs_table, Jour_table, PType_Entity, PP_table, ppe1]

        comments = ['SOV Table', 'Month-on-Month Table', 'Publication Table', 'Journalist Table','Pub Type and Pub Name Table', 'Pub Type and Entity Table', 'PubType PubName and Entity Table']

        # Sidebar for download options
        st.sidebar.write("## Download Options")
        download_formats = st.sidebar.selectbox("Select format:", ["Excel", "CSV", "Excel (Entity Sheets)"])
        file_name_data = st.sidebar.text_input("Enter file name for all DataFrames", "entitydata.xlsx")

        if st.sidebar.button("Download Data"):
            if download_formats == "Excel":
                # Create a link to download the Excel file for data
                excel_path = os.path.join(download_path, "data.xlsx")
                with pd.ExcelWriter(excel_path, engine="xlsxwriter", mode="xlsx") as writer:
                    data.to_excel(writer, index=False)

                st.sidebar.write(f"Excel file saved at {excel_path}")
#                 excel_io_data = io.BytesIO()
#                 with pd.ExcelWriter(excel_io_data, engine="xlsxwriter", mode="xlsx") as writer:
#                     data.to_excel(writer, index=False)
#                 excel_io_data.seek(0)
#                 b64_data = base64.b64encode(excel_io_data.read()).decode()
#                 href_data = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_data}" download="data.xlsx">Download Data Excel</a>'
#                 st.sidebar.markdown(href_data, unsafe_allow_html=True)

            elif download_formats == "CSV":
                # Create a link to download the CSV file for data
#                 csv_io_data = io.StringIO()
                csv_path = os.path.join(download_path, "data.csv")
                data.to_csv(csv_path, index=False)
                st.sidebar.write(f"CSV file saved at {csv_path}")
                
#                 data.to_csv(csv_io_data, index=False)
#                 csv_io_data.seek(0)
#                 b64_data = base64.b64encode(csv_io_data.read().encode()).decode()
#                 href_data = f'<a href="data:text/csv;base64,{b64_data}" download="data.csv">Download Data CSV</a>'
#                 st.sidebar.markdown(href_data, unsafe_allow_html=True)

            elif download_formats == "Excel (Entity Sheets)":
                # Create a link to download separate Excel sheets by Entity
#                 excel_io_sheets = io.BytesIO()
                excel_path_sheets = os.path.join(download_path, file_name_data)
                with pd.ExcelWriter(excel_path_sheets, mode="w", date_format='yyyy-mm-dd', datetime_format='yyyy-mm-dd') as writer:
                    create_entity_sheets(data, writer)

                st.sidebar.write(f"Excel sheets saved at {excel_path_sheets}")
#                 with pd.ExcelWriter(excel_io_sheets, engine="xlsxwriter", mode="xlsx" , date_format='yyyy-mm-dd', datetime_format='yyyy-mm-dd') as writer:
#                     create_entity_sheets(data, writer)
#                 excel_io_sheets.seek(0)
#                 b64_sheets = base64.b64encode(excel_io_sheets.read()).decode()
#                 href_sheets = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_sheets}" download="{file_name_data}">Download Entity Sheets Excel</a>'
#                 st.sidebar.markdown(href_sheets, unsafe_allow_html=True)

        # Download selected DataFrame
        st.sidebar.write("## Download Selected DataFrame")
        
                # Create a dropdown to select the DataFrame to download
        dataframes_to_download = {
            "Entity_SOV1": Entity_SOV1,
            "Data": data,
            "Finaldata": finaldata,
            "Month-on-Month":sov_dt,
            "Publication Table":pubs_table,
            "Journalist Table":Jour_table,
            "Publication Type and Name Table":PP_table,
            "Publication Type Table with Entity":PType_Entity,
            "Publication type,Publication Name and Entity Table":ppe1,
            "Entity-wise Sheets": finaldata  # Add this option to download entity-wise sheets
        }
        
        selected_dataframe = st.sidebar.selectbox("Select DataFrame:", list(dataframes_to_download.keys()))

        if st.sidebar.button("Download Selected DataFrame"):
            if selected_dataframe in dataframes_to_download:
                # Create a link to download the selected DataFrame in Excel
                selected_df = dataframes_to_download[selected_dataframe]
                excel_io_selected = io.BytesIO()
                with pd.ExcelWriter(excel_io_selected, engine="xlsxwriter", mode="xlsx") as writer:
                    selected_df.to_excel(writer, index=True)
                excel_io_selected.seek(0)
                b64_selected = base64.b64encode(excel_io_selected.read()).decode()
                href_selected = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_selected}" download="{selected_dataframe}.xlsx">Download {selected_dataframe} Excel</a>'
                st.sidebar.markdown(href_selected, unsafe_allow_html=True)
                 
        # Download All DataFrames as a Single Excel Sheet
        st.sidebar.write("## Download All DataFrames as a Single Excel Sheet")
        file_name_all = st.sidebar.text_input("Enter file name for all DataFrames", "all_dataframes.xlsx")
#         download_options = st.sidebar.selectbox("Select Download Option:", [ "Complete Dataframes"])
        
        if st.sidebar.button("Download All DataFrames"):
            # List of DataFrames to save
            dfs = [Entity_SOV1, sov_dt, pubs_table, Jour_table, PType_Entity, PP_table, ppe1]
            comments = ['SOV Table', 'Month-on-Month Table', 'Publication Table', 'Journalist Table',
                        'Pub Type and Entity Table', 'Pub Type and Pub Name Table',
                        'PubType PubName and Entity Table']
            
            excel_path_all = os.path.join(download_path, file_name_all)
            multiple_dfs(dfs, 'Tables', excel_path_all, 2, comments)
            st.sidebar.write(f"All DataFrames saved at {excel_path_all}")
            
#             # Create a link to download all DataFrames as a single Excel sheet with separation
#             excel_io_all = io.BytesIO()
#             multiple_dfs(dfs, 'Tables', excel_io_all, 2, comments)
#             excel_io_all.seek(0)
#             b64_all = base64.b64encode(excel_io_all.read()).decode()
#             href_all = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_all}" download="{file_name_all}">Download All DataFrames Excel</a>'
#             st.sidebar.markdown(href_all, unsafe_allow_html=True)
            
            
        # Download Top 10 DataFrames as a Single Excel Sheet
        st.sidebar.write("## Download Top N DataFrames as a Single Excel Sheet")
        file_name_topn = st.sidebar.text_input("Enter file name for all DataFrames", "top_dataframes.xlsx")
        # Slider to select the range of dataframes
        selected_range = st.sidebar.slider("Select start range:", 10, 50, 10)
        
        if st.sidebar.button("Download Top DataFrames"):
            # List of DataFrames to save
            selected_dfs = [Entity_SOV1, sov_dt, pubs_table, Jour_table, PType_Entity, PP_table, ppe1]
            comments_selected = ['SOV Table', 'Month-on-Month Table', 'Publication Table', 'Journalist Table',
                        'Pub Type and Entity Table', 'Pub Type and Pub Name Table',
                        'PubType PubName and Entity Table']
            top_n_flags = [False, False, True, True, True, True, True]
            
            # Create a link to download all DataFrames as a single Excel sheet with two sheets
            selected_dfs = [df.head(selected_range) for df in selected_dfs]
            comments_selected = comments_selected[:selected_range]
            top_n_flags = top_n_flags[:selected_range]

            excel_path_topn = os.path.join(download_path, file_name_topn)
            top_10_dfs(selected_dfs, excel_path_topn, comments_selected, top_n_flags)
            st.sidebar.write(f"Selected DataFrames saved at {excel_path_topn}")
            
#             excel_io_all = io.BytesIO()
#             top_10_dfs(dfs1, excel_io_all, commentss, top_11_flags)  # Save the top 10 rows in the first sheet                      
#             excel_io_all.seek(0)
#             b64_all = base64.b64encode(excel_io_all.read()).decode()
#             href_all = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64_all}" download="file_name_top10">Download Top10 DataFrames Excel</a>'
#             st.sidebar.markdown(href_all, unsafe_allow_html=True)
            
    else:
        st.sidebar.write("Please upload a file.")
        
            # Preview selected DataFrame in the main content area
    st.write("## Preview Selected DataFrame")
    selected_dataframe = st.selectbox("Select DataFrame to Preview:", list(dataframes_to_download.keys()))
    st.dataframe(dataframes_to_download[selected_dataframe])
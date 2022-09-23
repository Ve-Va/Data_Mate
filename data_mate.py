# streamlit app to work with excel files...
import base64
from io import BytesIO, StringIO, TextIOWrapper
from openpyxl import Workbook
from requests import options
import streamlit as st
import pandas as pd
import numpy as np
import datetime
import openpyxl
import xlrd
import xlsxwriter as xw
from xfilios.excel import ExcelHandler


# This app is to work with excel files...
# to fulfill most of the data manipulation needs and to make it easy to work with excel files...

# This is the main function in which we will build the actual app

# Wide Screen
st.set_page_config(layout="wide")

st.title("Data Mate")
st.subheader("To fulfill most of the excel/csv data manipulation needs and to make it easy to work with excel files...")
st.text("This app is in development phase...\nDeveloper: Venkat Varun Gundapuneedi")

# File Uploader for Excel or CSV files
today = datetime.date.today().strftime("%m%d")

# function to generate the excel file
# and to download the file
def link_to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0'})
    worksheet.set_column('A:A', None, format1)
    writer.save()
    processed_data = output.getvalue()
    b64 = base64.b64encode(processed_data).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="output_{today}.xlsx">Download Excel File</a>'
    return href

# function for csv file
def link_to_csv(df):
    output = BytesIO()
    df.to_csv(output, index=False)
    processed_data = output.getvalue()
    b64 = base64.b64encode(processed_data).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="output_{today}.csv">Download CSV File</a>'
    return href
                
# Option to check and keep only needed columns
st.subheader("Check and keep only needed columns and download the file as csv or excel")
uploaded_file = st.file_uploader("Choose a file", type = ['xlsx', 'csv'])
if uploaded_file is not None:
    if uploaded_file.name.endswith('.xlsx'):
        df = pd.read_excel(uploaded_file)
    elif uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    st.dataframe(df)
    if st.checkbox("Check and keep only needed columns"):
        df = df[st.multiselect("Select columns to keep", df.columns)]
        st.dataframe(df)
        # radio button to select the file type to download
        if st.checkbox("Download the file with needed columns"):
            file_type = st.radio("Select the file type to download", ('Excel', 'CSV'))
            if file_type == 'Excel':
                st.markdown(link_to_excel(df), unsafe_allow_html=True)
            elif file_type == 'CSV':
                st.markdown(link_to_csv(df), unsafe_allow_html=True)

    if st.checkbox("Rename the columns"):
        new_names = st.multiselect("Select columns to rename", df.columns)
        for i in new_names:
            new_name = st.text_input(f"Enter the new name for {i}", i)
            df.rename(columns = {i: new_name}, inplace = True)
        st.dataframe(df)
        # radio button to select the file type to download
        if st.checkbox("Download the file with renamed columns"):
            file_type = st.radio("Select the file type to download", ('Excel', 'CSV'))
            if file_type == 'Excel':
                st.markdown(link_to_excel(df), unsafe_allow_html=True)
            elif file_type == 'CSV':
                st.markdown(link_to_csv(df), unsafe_allow_html=True)
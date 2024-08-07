# Importing useful libraries
import streamlit as st
from scripts import utils
import pandas as pd
import os
import openpyxl

def common_elements(lists):
    # Convert the first list to a set
    common_set = set(lists[0])

    # Find intersection with each subsequent list
    for lst in lists[1:]:
        common_set.intersection_update(lst)

    # Return the result as a list
    return list(common_set)

# Creating a title and icon to the webpage
st.set_page_config(
    page_title="Month Wise Comparision",
    page_icon="📚"
)

st.header("Month wise revenue comparision Simple App")

st.subheader("Data Uploading Section")

number_of_months = st.number_input("Enter how many months you want to compare:", min_value=2, max_value=12, value=2, step=1)


comprehensive_reports = st.file_uploader(
    'Please Upload the comprehensive report',
    type = ['xlsx'],
    accept_multiple_files=True
)

# Create a radio button for navigation
needed_affiliate = st.selectbox("Navigate to:", [
    '',
    'All Inbox(Jason Jacobs)', 
    'Flatiron Media', 
    'Flex Marketing Group, LLC', 
    'Fluent',
    'Interest Media', 
    'Master Affiliate CD1', 
    'Master Affiliate CD3', 
    'Pure Ads Digital', 
    'Push Crew', 
    'Pushnami LLC', 
    'Revenue from JS Midpath for Website1', 
    'SBG Group', 
    'Survey Club', 
    'Tiburon Media', 
    'Union Street Enterprises', 
    'Unique Rewards', 
    'Webclients FrontStreetMarktng', 
    'What If Holdings', 
    'Zeeto Media'
])


if st.button('Generate Report'):

    if comprehensive_reports:

        if len(comprehensive_reports) == number_of_months:

            if needed_affiliate != '':
                
                with st.spinner('Generating Report...'):

                    month_data_frames = []

                    for file in comprehensive_reports:

                        temp_dfs = utils.generate_report(file)

                        # Getting the monthly data for February
                        feb_dfs = {}
                        for sheet_name,sheet_df in temp_dfs.items():
                            feb_dfs[sheet_name] = sheet_df

                        month_data_frames.append(feb_dfs)

                    # Let us get the comparision dfs
                    final_df = pd.concat([month_df[needed_affiliate] for month_df in month_data_frames],axis=0)
                    final_df.drop(
                        columns = 'Date',
                        inplace = True
                    )
                    final_df.dropna(inplace = True)
                    final_df.sort_values('Publisher',inplace=True)
                    final_df.to_excel('report_2.xlsx',index=False)

                    excel_path = 'report_2.xlsx'
                    # Load the Excel file with openpyxl
                    workbook = openpyxl.load_workbook(excel_path)

                    # Apply formatting to all sheets
                    for sheet_name in workbook.sheetnames:
                        sheet = workbook[sheet_name]
                        utils.format_sheet(sheet)

                    workbook.save(excel_path)

                    # Read the file into a binary stream
                    with open(excel_path, "rb") as file:
                        downloadable_data = file.read()

                    st.success("Finished generating")

                    # Providing a download button
                    st.download_button(
                        label="Download Excel File",
                        data=downloadable_data,
                        file_name="report.xlsx",
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )

                    os.remove(excel_path)

            else:

                st.warning('Please first select the affiliate!!!')

        else:
            st.warning(f"Please input {number_of_months} months comprehensive reports before!!!")

        

    else:

        st.warning("Upload a file first please")
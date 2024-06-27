
# Import libraries
import streamlit as st
import pandas as pd
from io import BytesIO
import base64
# import locale
# locale.setlocale( locale.LC_ALL, 'en_ZA.ANSI' )
# st.set_page_config(layout="centered")

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
    processed_data = output.getvalue()
    return processed_data

def get_table_download_link(df, filename="transformed_data.xlsx"):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    val = to_excel(df)
    b64 = base64.b64encode(val).decode()  # Some strings <-> bytes conversions necessary here
    return f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">Download Excel file</a>'

st.title('Rep Sell Out & Stock on Hand')

# Date_End = st.date_input("Week ending: ")
# Date_Start = Date_End - dt.timedelta(days=6)

# if Date_End.day < 10:
#     Day = '0'+str(Date_End.day)
# else:
#     Day = str(Date_End.day)

# Month = Date_End.month

# Year = str(Date_End.year)
# Short_Date_Dict = {1:'Jan', 2:'Feb', 3:'Mar',4:'Apr',5:'May',6:'Jun',7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}
# Long_Date_Dict = {1:'January', 2:'February', 3:'March',4:'April',5:'May',6:'June',7:'July',8:'August',9:'September',10:'October',11:'November',12:'December'}
# Country_Dict = {'AO':'Angola', 'MW':'Malawi', 'MZ':'Mozambique', 'NG':'Nigeria', 'UG':'Uganda', 'ZA':'South Africa', 'ZM':'Zambia', 'ZW':'Zimbabwe'}



st.write("")
st.markdown("This will handle the following reps: **Bernie, Lee, Ryan**")
st.markdown("Please make sure the sheets in your file are named correctly")

# Streamlit file uploader
# data_file = st.file_uploader('Rep File',type=['csv','txt','xlsx','xls'])
# if data_file:    
#     df = pd.read_excel(data_file)


uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

# Check if a file has been uploaded
if uploaded_file is not None:
    sheet_names = ['Bernie', 'Ryan', 'Lee']

    def transform_data(df, sheet_name):
        # Save the current header
        old_header = df.columns.tolist()

        # Insert the old header as the first row
        df.loc[-1] = old_header  # Add old header as a row at index -1
        df.index = df.index + 1  # Shift index
        df = df.sort_index()     # Sort index to move the new row to the top

        # Create new header with 'Unnamed:' prefix
        new_header = ['Unnamed: ' + str(i) for i in range(len(df.columns))]
        df.columns = new_header

        # Concatenate the first 4 rows with a delimiter '|'
        new_header = df.iloc[0:4].apply(lambda x: ' | '.join(x.dropna().astype(str)), axis=0)

        # Drop the first 4 rows and set new header
        df.columns = new_header
        df = df.iloc[4:].reset_index(drop=True)

        # Keep the first 4 columns
        id_vars = new_header[:4]

        # Unpivot the remaining columns
        melted_df = pd.melt(df, id_vars=id_vars, var_name='Variable', value_name='Value')

        filterdf_SOH = melted_df[~melted_df['Variable'].str.contains('Sell Out', na=False)]

        # Resetting index for filterdf_SOH
        filterdf_SOH = filterdf_SOH.reset_index(drop=True)

        filterdf_Sales = melted_df[~melted_df['Variable'].str.contains('Week', na=False)]

        # Resetting index for filterdf_Sales
        filterdf_Sales = filterdf_Sales.reset_index(drop=True)

        # Add 'Sales' from df2 to df1 using .loc
        filterdf = filterdf_SOH
        filterdf.loc[:,'Sell Out'] = filterdf_Sales['Value']

        filterdf = filterdf[~filterdf['Variable'].str.contains('Notes', na=False)]

        # Rename columns
        df = filterdf.rename(columns={
            'Unnamed: 0 | 365 Code': '365 Code',
            'Unnamed: 1 | Product Description': 'Product Description',
            'Unnamed: 2 | Category': 'Category',
            'Unnamed: 3 | Date SOH was Collected: | Sub-Cat' : 'Sub-Cat',
            'Value' : 'Stock on Hand'
        })

        # Split 'Variable' based on '|'
        df[['Retailer', 'Date SOH was Collected', 'Week No.', 'Dummy']] = df['Variable'].str.split('|', expand=True)

        # Drop 'Dummy' and 'Variable' columns
        df = df.drop(['Dummy', 'Variable'], axis=1)

        # Convert 'Sell Out' and 'Stock on Hand' column to integer
        df['Sell Out'] = pd.to_numeric(df['Sell Out'], errors='coerce').fillna(0).astype(int)
        df['Stock on Hand'] = pd.to_numeric(df['Stock on Hand'], errors='coerce').fillna(0).astype(int)

        # Strip spaces from 'Retailer' column
        df['Retailer'] = df['Retailer'].str.strip()

        # Remove dots and subsequent numbers, and then strip spaces from 'Retailer' column
        df["Retailer"] = df["Retailer"].str.replace(r"\.*\d+", "", regex=True)

        # Convert 'Date SOH was Collected' column to date type
        df['Date SOH was Collected'] = pd.to_datetime(df['Date SOH was Collected']).dt.date

        # Add rep column
        df['Rep'] = sheet_name

        # Filter out records where 'Retailer' column contains 'Unnamed'
        df = df[~df['Retailer'].str.contains('Unnamed')]

        return df

    # Read and transform data from each sheet
    transformed_dfs = []
    for sheet in sheet_names:
        df = pd.read_excel(uploaded_file, sheet_name=sheet)
        transformed_df = transform_data(df, sheet)
        transformed_dfs.append(transformed_df)

    # Concatenate all transformed DataFrames
    final_df = pd.concat(transformed_dfs, ignore_index=True)

    # Display the final DataFrame
    st.write("Transformed Data:")
    st.dataframe(final_df)

    st.markdown(get_table_download_link(final_df), unsafe_allow_html=True)

# # Ackermans
# if option == 'Ackermans':

#     if Date_End.month < 10:
#         Month = '0'+str(Date_End.month)
#     else:
#         Month = str(Date_End.month)

#     Units_Sold = 'Sales: ' + Year + '/' + str(Month) + '/' + Day
#     CSOH = 'CSOH: ' + Year + '/' + str(Month) + '/' + Day



#     try:
#         # Get retailers map
#         df_ackermans_retailers_map = df_map
#         df_ackermans_retailers_map = df_ackermans_retailers_map.rename(columns={'Style Code': 'SKU No.'})
#         df_ackermans_retailers_map_final = df_ackermans_retailers_map[['SKU No.','Product Description','SMD Product Code']]
        
#         # Get retailer data
#         df_ackermans_data = df_data
#         df_ackermans_data['SKU No.'] = df_ackermans_data['Style'].astype(str).str.split(' ').str[0]
        
#         # Merge with retailer map
#         df_ackermans_data['SKU No.'] = df_ackermans_data['SKU No.'].astype(int)
#         df_ackermans_merged = df_ackermans_data.merge(df_ackermans_retailers_map_final, how='left', on='SKU No.')

#         # Find missing data
#         missing_model_ackermans = df_ackermans_merged['SMD Product Code'].isnull()
#         df_ackermans_missing_model = df_ackermans_merged[missing_model_ackermans]
#         df_missing = df_ackermans_missing_model[['SKU No.','Style']]
#         df_missing_unique = df_missing.drop_duplicates()
#         st.write("The following products are missing the SMD code on the map: ")
#         st.table(df_missing_unique)
#         st.write(" ")

#     except:
#         st.markdown("**Retailer map column headings:** Style Code, Product Description, SMD Product Code")
#         st.markdown("**Retailer data column headings:** Store, Style, Closing Stock Units, Nett Sale Units, Nett Sale Value")
#         st.markdown("Column headings are **case sensitive.** Please make sure they are correct") 

        
#     try:
#         # Set date columns
#         df_ackermans_merged['Start Date'] = Date_End

#         # Add retailer column and store column
#         df_ackermans_merged['Forecast Group'] = 'Ackermans'

#         # Rename columns
#         df_ackermans_merged = df_ackermans_merged.rename(columns={'Closing Stock Units': 'SOH Qty'})
#         df_ackermans_merged = df_ackermans_merged.rename(columns={'Nett Sale Units': 'Sales Qty'})
#         df_ackermans_merged = df_ackermans_merged.rename(columns={'Nett Sale Value': 'Total Amt'})
#         df_ackermans_merged = df_ackermans_merged.rename(columns={'SMD Product Code': 'Product Code'})
#         df_ackermans_merged = df_ackermans_merged.rename(columns={'Store': 'Store Name'})

#         # Don't change these headings. Rather change the ones above
#         final_df_ackermans = df_ackermans_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
#         final_df_ackermans_p = df_ackermans_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
#         final_df_ackermans_s = df_ackermans_merged[['Store Name','Total Amt']]

#         # Show final df
#         df_stats(final_df_ackermans, final_df_ackermans_p, final_df_ackermans_s)

#         # Output to .xlsx
#         st.write('Please ensure that no products are missing before downloading!')
#         st.markdown(get_table_download_link(final_df_ackermans), unsafe_allow_html=True)

#     except:
#         st.write('Check data')


# # Bradlows/Russels
# elif option == 'Bradlows/Russels':
#     try:
#         # Get retailers map
#         df_br_retailers_map = df_map
#         df_br_retailers_map = df_br_retailers_map.rename(columns={'Article Number':'SKU No. B&R'})
#         df_br_retailers_map = df_br_retailers_map[['SKU No. B&R','Product Code','Product Description','RSP']]

#         # Get retailer data
#         df_br_data = df_data
#         df_br_data.columns = df_br_data.iloc[1]
#         df_br_data = df_br_data.iloc[2:]

#         # Fill sales qty
#         df_br_data['Sales Qty*'].fillna(0,inplace=True)

#         # Drop result rows
#         df_br_data.drop(df_br_data[df_br_data['Article'] == 'Result'].index, inplace = True) 
#         df_br_data.drop(df_br_data[df_br_data['Site'] == 'Result'].index, inplace = True) 
#         df_br_data.drop(df_br_data[df_br_data['Cluster'] == 'Overall Result'].index, inplace = True) 

#         # Get SKU No. column
#         df_br_data['SKU No. B&R'] = df_br_data['Article'].astype(float)

#         # Site columns
#         df_br_data['Store Name'] = df_br_data['Site'] + ' - ' + df_br_data['Site Name'] 

#         # Consolidate
#         df_br_data_new = df_br_data[['Cluster','SKU No. B&R','Description','Store Name','Sales Qty*','Valuated Stock Qty(Total)']]

#         # Merge with retailer map
#         df_br_data_merged = df_br_data_new.merge(df_br_retailers_map, how='left', on='SKU No. B&R',indicator=True)

#         # Find missing data
#         missing_model_br = df_br_data_merged['Product Code'].isnull()
#         df_br_missing_model = df_br_data_merged[missing_model_br]
#         df_missing = df_br_missing_model[['SKU No. B&R','Description']]
#         df_missing_unique = df_missing.drop_duplicates()
#         st.write("The following products are missing the SMD code on the map: ")
#         st.table(df_missing_unique)
#         st.write(" ")

#         missing_rsp_br = df_br_data_merged['RSP'].isnull()
#         df_br_missing_rsp = df_br_data_merged[missing_rsp_br]
#         df_missing_2 = df_br_missing_rsp[['SKU No. B&R','Description']]
#         df_missing_unique_2 = df_missing_2.drop_duplicates()
#         st.write("The following products are missing the RSP on the map: ")
#         st.table(df_missing_unique_2)
        
#     except:
#         st.markdown("**Retailer map column headings:** Article Number, Product Code, Product Description & RSP")
#         st.markdown("**Retailer data column headings:** Cluster, Article, Description, Site, Site Name, Valuated Stock Qty(Total), Sales Qty*")
#         st.markdown("Column headings are **case sensitive.** Please make sure they are correct") 

#     try:
#         # Set date columns
#         df_br_data_merged['Start Date'] = Date_Start

#         # Total amount column
#         df_br_data_merged['Total Amt'] = df_br_data_merged['Sales Qty*'] * df_br_data_merged['RSP']

#         # Tidy columns
#         df_br_data_merged['Forecast Group'] = 'Bradlows/Russels'
#         df_br_data_merged['Store Name']= df_br_data_merged['Store Name'].str.title() 

#         # Rename columns
#         df_br_data_merged = df_br_data_merged.rename(columns={'Sales Qty*': 'Sales Qty'})
#         df_br_data_merged = df_br_data_merged.rename(columns={'SKU No. B&R': 'SKU No.'})
#         df_br_data_merged = df_br_data_merged.rename(columns={'Valuated Stock Qty(Total)': 'SOH Qty'})

#         # Don't change these headings. Rather change the ones above
#         final_df_br = df_br_data_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
#         final_df_br_p = df_br_data_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
#         final_df_br_s = df_br_data_merged[['Store Name','Total Amt']]

#         # Show final df
#         df_stats(final_df_br,final_df_br_p,final_df_br_s)

#         # Output to .xlsx
#         st.write('Please ensure that no products are missing before downloading!')
#         st.markdown(get_table_download_link(final_df_br), unsafe_allow_html=True)

#     except:
#         st.write('Check data')


# else:
#     st.write('Retailer not selected yet')
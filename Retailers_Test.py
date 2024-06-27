# Import libraries
import streamlit as st
import numpy as np
import pandas as pd
import base64
from io import BytesIO
import datetime as dt
# import locale
# locale.setlocale( locale.LC_ALL, 'en_ZA.UTF-8' )

def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter') # pylint: disable=abstract-class-instantiated
    df.to_excel(writer, sheet_name='Sheet1',index=False)
    writer.save()
    processed_data = output.getvalue()
    return processed_data

def get_table_download_link(df):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    val = to_excel(df)
    b64 = base64.b64encode(val)
    return f'<a href="data:application/octet-stream;base64,{b64.decode()}" download='+option+'_'+Year+str(Month)+Day+".xlsx"'>Download Excel file</a>' # decode b'abc' => abc

st.title('Retailer Sales Reports (TEST)')

Date_End = st.date_input("Week ending: ")
Date_Start = Date_End - dt.timedelta(days=6)

if Date_End.day < 10:
    Day = '0'+str(Date_End.day)
else:
    Day = str(Date_End.day)
Month = Date_End.month
Year = str(Date_End.year)
Short_Date_Dict = {1:'Jan', 2:'Feb', 3:'Mar',4:'Apr',5:'May',6:'Jun',7:'Jul',8:'Aug',9:'Sep',10:'Oct',11:'Nov',12:'Dec'}
Long_Date_Dict = {1:'January', 2:'February', 3:'March',4:'April',5:'May',6:'June',7:'July',8:'August',9:'September',10:'October',11:'November',12:'December'}
Country_Dict = {'AO':'Angola', 'MW':'Malawi', 'MZ':'Mozambique', 'NG':'Nigeria', 'UG':'Uganda', 'ZA':'South Africa', 'ZM':'Zambia', 'ZW':'Zimbabwe'}

option = st.selectbox(
    'Please select a retailer:',
    ('Please select','Ackermans','Bradlows/Russels','Builders','Checkers','Clicks','Dealz','Dis-Chem','Dis-Chem-Pharmacies', 'H&H','HiFi','Incredible-Connection','Makro', 'Musica','Ok-Furniture', 'Outdoor-Warehouse','Pep-Africa','Pep-SA','PnP','Sportsmans-Warehouse','Takealot','TFG'))
st.write('You selected:', option)

st.write("")
st.markdown("Please ensure data is in the **_first sheet_** of your Excel Workbook")

map_file = st.file_uploader('Retailer Map', type='xlsx')
if map_file:
    df_map = pd.read_excel(map_file)

data_file = st.file_uploader('Weekly Sales Data',type='xlsx')
if data_file:
    df_data = pd.read_excel(data_file)

# Ackermans
if option == 'Ackermans':

    Units_Sold = 'Sales: ' + Day + '/' + str(Month) + '/' + Year
    CSOH = 'CSOH: ' + Day + '/' + str(Month) + '/' + Year

    try:
        # Get retailers map
        df_ackermans_retailers_map = df_map
        df_ackermans_retailers_map.columns = df_ackermans_retailers_map.iloc[1]
        df_ackermans_retailers_map = df_ackermans_retailers_map.iloc[2:]
        df_ackermans_retailers_map = df_ackermans_retailers_map.rename(columns={'Style Code': 'SKU No.'})
        df_ackermans_retailers_map_final = df_ackermans_retailers_map[['SKU No.','Product Description','SMD Product Code','SMD RSP']]

        # Get retailer data
        df_ackermans_data = df_data
        df_ackermans_data.columns = df_ackermans_data.iloc[6]
        df_ackermans_data = df_ackermans_data.iloc[7:]

        # Merge with retailer map
        df_ackermans_data['SKU No.'] = df_ackermans_data['Style Code'].astype(int)
        df_ackermans_merged = df_ackermans_data.merge(df_ackermans_retailers_map_final, how='left', on='SKU No.')

        # Find missing data
        missing_model_ackermans = df_ackermans_merged['SMD Product Code'].isnull()
        df_ackermans_missing_model = df_ackermans_merged[missing_model_ackermans]
        df_missing = df_ackermans_missing_model[['SKU No.','Style Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)
        st.write(" ")

        missing_rsp_ackermans = df_ackermans_merged['SMD RSP'].isnull()
        df_ackermans_missing_rsp = df_ackermans_merged[missing_rsp_ackermans]
        df_missing_2 = df_ackermans_missing_rsp[['SKU No.','Style Description']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** Style Code, Product Description, SMD Product Code & SMD RSP")
        st.markdown("**Retailer data column headings:** Style Code, Style Description, " + CSOH +", "+ Units_Sold)
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct") 

        
    # try:
    # Set date columns
    df_ackermans_merged['Start Date'] = Date_Start

    # Total amount column
    df_ackermans_merged['Total Amt'] = df_ackermans_merged[Units_Sold].astype(int) * df_ackermans_merged['SMD RSP']

    # Add retailer column and store column
    df_ackermans_merged['Forecast Group'] = 'Ackermans'
    df_ackermans_merged['Store Name'] = ''
    df_ackermans_merged['Style Description'] = df_ackermans_merged['Style Description'].str.title() 

    # Rename columns
    df_ackermans_merged = df_ackermans_merged.rename(columns={CSOH: 'SOH Qty'})
    df_ackermans_merged = df_ackermans_merged.rename(columns={Units_Sold: 'Sales Qty'})
    df_ackermans_merged = df_ackermans_merged.rename(columns={'SMD Product Code': 'Product Code'})

    # Don't change these headings. Rather change the ones above
    final_df_ackermans = df_ackermans_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
    final_df_ackermans_p = df_ackermans_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
    final_df_ackermans_s = df_ackermans_merged[['Store Name','Total Amt']]

    # Show final df
    total = final_df_ackermans['Total Amt'].sum()
    total_units = final_df_ackermans['Sales Qty'].sum()
    st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
    st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
    st.write('')
    st.write('**Top 10 products for the week:**')
    grouped_df_pt = final_df_ackermans_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
    grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
    st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
    st.write('')
    st.write('**Top 10 stores for the week:**')
    grouped_df_st = final_df_ackermans_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
    grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
    st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
    st.write('')
    st.write('**Bottom 10 products for the week:**')
    grouped_df_pb = final_df_ackermans_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
    grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
    st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
    st.write('')
    st.write('**Bottom 10 stores for the week:**')
    grouped_df_sb = final_df_ackermans_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
    grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
    st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))

    st.write('**Final Dataframe:**')
    final_df_ackermans

    # Output to .xlsx
    st.write('Please ensure that no products are missing before downloading!')
    st.markdown(get_table_download_link(final_df_ackermans), unsafe_allow_html=True)

    # except:
    #     st.write('Check data')


# Bradlows/Russels
elif option == 'Bradlows/Russels':
    try:
        # Get retailers map
        df_br_retailers_map = df_map
        df_br_retailers_map = df_br_retailers_map.rename(columns={'Article Number':'SKU No. B&R'})
        df_br_retailers_map = df_br_retailers_map[['SKU No. B&R','Product Code','Product Description','RSP']]

        # Get retailer data
        df_br_data = df_data
        df_br_data.columns = df_br_data.iloc[1]
        df_br_data = df_br_data.iloc[2:]

        # Fill sales qty
        df_br_data['Sales Qty*'].fillna(0,inplace=True)

        # Drop result rows
        df_br_data.drop(df_br_data[df_br_data['Article'] == 'Result'].index, inplace = True) 
        df_br_data.drop(df_br_data[df_br_data['Site'] == 'Result'].index, inplace = True) 
        df_br_data.drop(df_br_data[df_br_data['Cluster'] == 'Overall Result'].index, inplace = True) 

        # Get SKU No. column
        df_br_data['SKU No. B&R'] = df_br_data['Article'].astype(float)

        # Site columns
        df_br_data['Store Name'] = df_br_data['Site'] + ' - ' + df_br_data['Site Name'] 

        # Consolidate
        df_br_data_new = df_br_data[['Cluster','SKU No. B&R','Description','Store Name','Sales Qty*','Valuated Stock Qty(Total)']]

        # Merge with retailer map
        df_br_data_merged = df_br_data_new.merge(df_br_retailers_map, how='left', on='SKU No. B&R',indicator=True)

        # Find missing data
        missing_model_br = df_br_data_merged['Product Code'].isnull()
        df_br_missing_model = df_br_data_merged[missing_model_br]
        df_missing = df_br_missing_model[['SKU No. B&R','Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)
        st.write(" ")

        missing_rsp_br = df_br_data_merged['RSP'].isnull()
        df_br_missing_rsp = df_br_data_merged[missing_rsp_br]
        df_missing_2 = df_br_missing_rsp[['SKU No. B&R','Description']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)
        
    except:
        st.markdown("**Retailer map column headings:** Article Number, Product Code, Product Description & RSP")
        st.markdown("**Retailer data column headings:** Cluster, Article, Description, Site, Site Name, Valuated Stock Qty(Total), Sales Qty*")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct") 

    try:
        # Set date columns
        df_br_data_merged['Start Date'] = Date_Start

        # Total amount column
        df_br_data_merged['Total Amt'] = df_br_data_merged['Sales Qty*'] * df_br_data_merged['RSP']

        # Tidy columns
        df_br_data_merged['Forecast Group'] = 'Bradlows/Russels'
        df_br_data_merged['Store Name']= df_br_data_merged['Store Name'].str.title() 

        # Rename columns
        df_br_data_merged = df_br_data_merged.rename(columns={'Sales Qty*': 'Sales Qty'})
        df_br_data_merged = df_br_data_merged.rename(columns={'SKU No. B&R': 'SKU No.'})
        df_br_data_merged = df_br_data_merged.rename(columns={'Valuated Stock Qty(Total)': 'SOH Qty'})

        # Don't change these headings. Rather change the ones above
        final_df_br = df_br_data_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_br_p = df_br_data_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_br_s = df_br_data_merged[['Store Name','Total Amt']]

        # Show final df
        total = final_df_br['Total Amt'].sum()
        total_units = final_df_br['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_br_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_br_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_br_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_br_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))

        st.write('**Final Dataframe:**')
        final_df_br

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_br), unsafe_allow_html=True)

    except:
        st.write('Check data')

# Builders Warehouse

elif option == 'Builders':
    Week = st.text_input("Enter week number: ")
    weekly_sales = Week+'-'+Year[-1:]
    bw_stores = st.file_uploader('Stores', type='xlsx')
    if bw_stores:
        df_bw_stores = pd.read_excel(bw_stores)
   
    try:
        # Get retailers map
        df_bw_retailers_map = df_map
        df_bw_retailers_map = df_bw_retailers_map.rename(columns={'SMD Description':'Product Description'})
        df_retailers_map_bw_final = df_bw_retailers_map[['Article','SMD Product Code','Product Description']]

        # Get retailer data
        df_bw_data = df_data
        df_bw_data.columns = df_bw_data.iloc[6]
        df_bw_data = df_bw_data.iloc[8:]
        df_bw_data = df_bw_data.rename(columns={'  Incl SP': 'RSP'})
        df_bw_data = df_bw_data[df_bw_data['Article Description'].notna()]
        df_bw_data['RSP'] = df_bw_data['RSP'].replace(',','', regex=True)
        df_bw_data['RSP'] = df_bw_data['RSP'].astype(float)
        
        # Merge with retailer map 
        df_bw_merged = df_bw_data.merge(df_retailers_map_bw_final, how='left', on='Article')

        # Merge with stores
        df_bw_merged = df_bw_merged.merge(df_bw_stores, how='left', on='Site')
        
        # Find missing data
        missing_model_bw = df_bw_merged['SMD Product Code'].isnull()
        df_bw_missing_model = df_bw_merged[missing_model_bw]
        df_missing = df_bw_missing_model[['Article','Article Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp_bw = df_bw_merged['RSP'].isnull()
        df_bw_missing_rsp = df_bw_merged[missing_rsp_bw]  
        df_missing_2 = df_bw_missing_rsp[['Article','Article Description']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** Article, SMD Product Code")
        st.markdown("**Retailer data column headings:** Article, Article Desc, Site, Store Name (in Stores.xlsx), SOH, "+weekly_sales)
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_bw_merged['Start Date'] = Date_Start

        # Total amount column
        df_bw_merged[weekly_sales] = df_bw_merged[weekly_sales].astype(float)
        df_bw_merged['Total Amt'] = df_bw_merged[weekly_sales].astype(float) * df_bw_merged['RSP'].astype(float)
        
        # Add retailer column
        df_bw_merged['Forecast Group'] = 'Builders Warehouse'

        # Rename columns
        df_bw_merged = df_bw_merged.rename(columns={'Article': 'SKU No.'})
        df_bw_merged = df_bw_merged.rename(columns={'SMD Product Code': 'Product Code'})
        df_bw_merged = df_bw_merged.rename(columns={' SOH': 'SOH Qty'})
        df_bw_merged = df_bw_merged.rename(columns={weekly_sales: 'Sales Qty'})

        # Don't change these headings. Rather change the ones above
        final_df_bw = df_bw_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_bw_p = df_bw_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_bw_s = df_bw_merged[['Store Name','Total Amt']]

        # Show final df
        total = final_df_bw['Total Amt'].sum()
        total_units = final_df_bw['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_bw_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_bw_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_bw_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_bw_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))

        st.write('**Final Dataframe:**')
        final_df_bw

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_bw), unsafe_allow_html=True)

    except:
        st.write('Check data')

# Checkers

elif option == 'Checkers':

    checkers_soh = st.file_uploader('SOH', type='xlsx')
    if checkers_soh:
        df_checkers_soh = pd.read_excel(checkers_soh)

    Units_Sold = 'Units :'+ Day +' '+ Short_Date_Dict[Month] + ' ' + Year
    Value_Sold = 'Value :'+ Day +' '+ Short_Date_Dict[Month] + ' ' + Year

    try:
        # Get retailers data
        df_checkers_retailers_map = df_map

        # Get retailer data
        df_checkers_data = df_data
        df_checkers_data.columns = df_checkers_data.iloc[2]
        df_checkers_data = df_checkers_data.iloc[3:]
        df_checkers_data = df_checkers_data.rename(columns={'Item Code': 'Article'})
        df_checkers_data['Lookup'] = df_checkers_data['Article'].astype(str) + df_checkers_data['Branch']

        # Get stock on hand
        df_checkers_soh.columns = df_checkers_soh.iloc[2]
        df_checkers_soh = df_checkers_soh.iloc[3:]
        df_checkers_soh = df_checkers_soh.rename(columns={'Item Code': 'Article'})
        df_checkers_soh = df_checkers_soh.rename(columns={'Stock Qty':'SOH Qty'})
        df_checkers_soh['Lookup'] = df_checkers_soh['Article'].astype(str) + df_checkers_soh['Branch']
        df_checkers_soh_final = df_checkers_soh[['Lookup','SOH Qty']]
        
        # Merge SOH and Retailer Map
        df_checkers_data = df_checkers_data.merge(df_checkers_soh_final, how='left', on='Lookup')
        df_checkers_merged = df_checkers_data.merge(df_checkers_retailers_map, how='left', on='Article')
        
        # Find missing data
        missing_model_checkers = df_checkers_merged['SMD Product Code'].isnull()
        df_checkers_missing_model = df_checkers_merged[missing_model_checkers]
        df_missing = df_checkers_missing_model[['Article','Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)
        st.write(" ")

        missing_rsp_checkers = df_checkers_merged['RSP'].isnull()
        df_checkers_missing_rsp = df_checkers_merged[missing_rsp_checkers]
        df_missing_2 = df_checkers_missing_rsp[['Article','Description']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** Article, SMD Product Code, SMD Description & RSP")
        st.markdown("**Retailer data column headings:** Item Code, Description, "+Units_Sold)
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct") 

    try:
        # Add columns for dates
        df_checkers_merged['Start Date'] = Date_Start

        # Add Total Amount column
        df_checkers_merged = df_checkers_merged.rename(columns={Value_Sold: 'Total Amt'})
       
        # Add column for retailer and SOH
        df_checkers_merged['Forecast Group'] = 'Checkers'

        # Rename columns
        df_checkers_merged = df_checkers_merged.rename(columns={'Article': 'SKU No.'})
        df_checkers_merged = df_checkers_merged.rename(columns={Units_Sold: 'Sales Qty'})
        df_checkers_merged = df_checkers_merged.rename(columns={'SMD Product Code': 'Product Code'})
        df_checkers_merged = df_checkers_merged.rename(columns={'Branch': 'Store Name'})
        df_checkers_merged = df_checkers_merged.rename(columns={'SMD Description': 'Product Description'})

        # Final df. Don't change these headings. Rather change the ones above
        final_df_checkers_sales = df_checkers_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_checkers_p = df_checkers_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_checkers_s = df_checkers_merged[['Store Name','Total Amt']]

        # Show final df
        total = final_df_checkers_sales['Total Amt'].sum()
        total_units = final_df_checkers_sales['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_checkers_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_checkers_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_checkers_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_checkers_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))

        st.write('**Final Dataframe:**')
        final_df_checkers_sales

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_checkers_sales), unsafe_allow_html=True)

    except:
        st.write('Check data')

# Clicks

elif option == 'Clicks':
    try:
        # Get retailers map
        df_clicks_retailers_map = df_map
        df_retailers_map_clicks_final = df_clicks_retailers_map[['Clicks Product Number','SMD CODE','SMD DESC','RSP']]

        # Get retailer data
        df_clicks_data = df_data
        df_clicks_data.columns = df_clicks_data.iloc[3]
        df_clicks_data = df_clicks_data.iloc[5:]

        # Drop result rows
        df_clicks_data.drop(df_clicks_data[df_clicks_data['Product Status'] == 'Sum:'].index, inplace = True) 

        # Merge with retailer map 
        df_clicks_merged = df_clicks_data.merge(df_retailers_map_clicks_final, how='left', on='Clicks Product Number')

        # Find missing data
        missing_model_clicks = df_clicks_merged['SMD CODE'].isnull()
        df_clicks_missing_model = df_clicks_merged[missing_model_clicks]
        df_missing = df_clicks_missing_model[['Clicks Product Number','Product Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp_clicks = df_clicks_merged['RSP'].isnull()
        df_clicks_missing_rsp = df_clicks_merged[missing_rsp_clicks] 
        df_missing_2 = df_clicks_missing_rsp[['Clicks Product Number','Product Description']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)
    except:
        st.markdown("**Retailer map column headings:** Clicks Product Number,SMD CODE,SMD DESC,RSP")
        st.markdown("**Retailer data column headings:** Store Description, Clicks Product Number, Product Description, Store Stock Qty, Sales Qty LW TY")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")   

    try:
        # Set date columns
        df_clicks_merged['Start Date'] = Date_Start

        # Total amount column
        df_clicks_merged['Total Amt'] = df_clicks_merged['Sales Qty LW TY'] * df_clicks_merged['RSP']

        # Add retailer column
        df_clicks_merged['Forecast Group'] = 'Clicks'

        # Rename columns
        df_clicks_merged = df_clicks_merged.rename(columns={'Clicks Product Number': 'SKU No.'})
        df_clicks_merged = df_clicks_merged.rename(columns={'SMD CODE': 'Product Code'})
        df_clicks_merged = df_clicks_merged.rename(columns={'SMD DESC': 'Product Desc'})
        df_clicks_merged = df_clicks_merged.rename(columns={'Store Description': 'Store Name'})
        df_clicks_merged = df_clicks_merged.rename(columns={'Store Stock Qty': 'SOH Qty'})
        df_clicks_merged = df_clicks_merged.rename(columns={'Sales Qty LW TY': 'Sales Qty'})

        # Don't change these headings. Rather change the ones above
        final_df_clicks = df_clicks_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_clicks_p = df_clicks_merged[['Product Code','Product Desc','Total Amt']]
        final_df_clicks_s = df_clicks_merged[['Store Name','Total Amt']]

        # Show final df
        total = final_df_clicks['Total Amt'].sum()
        total_units = final_df_clicks['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_clicks_p.groupby("Product Desc").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_clicks_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_clicks_p.groupby("Product Desc").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_clicks_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))
        st.write('**Final Dataframe:**')
        final_df_clicks

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_clicks), unsafe_allow_html=True)
    
    except:
        st.write('Check data')

# Dealz
elif option == 'Dealz':

    units_sold = Long_Date_Dict[Month]

    try:
        # Get retailers map
        df_dealz_retailers_map = df_map
        df_retailers_map_dealz_final = df_dealz_retailers_map[['Style Code','Product Code','Product Description']]

        # Get retailer data
        df_dealz_data = df_data
        df_dealz_data.columns = df_dealz_data.iloc[5]
        df_dealz_data = df_dealz_data.iloc[6:]
        s = pd.Series(df_dealz_data.columns)
        s = s.fillna('Unnamed: ' + (s.groupby(s.isnull()).cumcount() + 1).astype(str))
        df_dealz_data.columns = s

        # Create SOH
        df_dealz_data['SOH Qty'] = df_dealz_data['Unnamed: 3'].astype(float) + df_dealz_data['Unnamed: 4'].astype(float)

        # Merge with Retailers Map
        df_dealz_merged = df_dealz_data.merge(df_retailers_map_dealz_final, how='left', on='Style Code')
        df_dealz_merged = df_dealz_merged[df_dealz_merged['Style Code'].notna()]

        # Find missing data
        missing_model = df_dealz_merged['Product Code'].isnull()
        df_dealz_missing_model = df_dealz_merged[missing_model]
        df_missing = df_dealz_missing_model[['Style Code','Style Desc']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)



    except:
        st.markdown("**Retailer map column headings:** Style Code, Product Code, Product Description")
        st.markdown("**Retailer data column headings:** Style Code, Style Desc, "+units_sold)
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct") 

    try:
        # Set date columns
        df_dealz_merged['Start Date'] = Date_Start

        # Add Total Amount column
        df_dealz_merged['Total Amt'] = df_dealz_merged[units_sold] * df_dealz_merged['Price']

        # Add column for retailer and store name
        df_dealz_merged['Forecast Group'] = 'Dealz'
        df_dealz_merged['Store Name'] = ''

        # Rename columns
        df_dealz_merged = df_dealz_merged.rename(columns={'Style Code': 'SKU No.'})
        df_dealz_merged = df_dealz_merged.rename(columns={units_sold: 'Sales Qty'})
        df_dealz_merged = df_dealz_merged.rename(columns={'Price': 'RSP'})

        # Final df. Don't change these headings. Rather change the ones above
        final_df_dealz_sales = df_dealz_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_dealz_p = df_dealz_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_dealz_s = df_dealz_merged[['Store Name','Total Amt']]

        # Show final df
        total = final_df_dealz_sales['Total Amt'].sum()
        total_units = final_df_dealz_sales['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_dealz_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_dealz_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_dealz_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_sb = final_df_dealz_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))
        st.write('**Final Dataframe:**')
        final_df_dealz_sales

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_dealz_sales), unsafe_allow_html=True)

    except:
        st.write('Check data')



# Dis-Chem

elif option == 'Dis-Chem':
    try:
        Units_Sold = (Short_Date_Dict[Month] + ' ' + Year)
        
        # Get retailers map
        df_dischem_retailers_map = df_map
        df_dischem_retailers_map = df_dischem_retailers_map.rename(columns={'Description': 'Product Description'})
        df_retailers_map_dischem_final = df_dischem_retailers_map[['Article Code','SMD Code','Product Description','RSP']]

        # Get retailer data
        df_dischem_data = df_data

        # Merge with retailer map
        df_dischem_merged = df_dischem_data.merge(df_retailers_map_dischem_final, how='left', on='Article Code')

        # Rename columns
        df_dischem_merged = df_dischem_merged.rename(columns={'Article Code': 'SKU No.'})
        df_dischem_merged = df_dischem_merged.rename(columns={'Oct 2020': 'Sales Qty'})
        df_dischem_merged = df_dischem_merged.rename(columns={'SMD Code': 'Product Code'})

        # Find missing data
        missing_model_dischem = df_dischem_merged['Product Code'].isnull()
        df_dischem_missing_model = df_dischem_merged[missing_model_dischem]
        df_missing = df_dischem_missing_model[['SKU No.','Article']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp_dischem = df_dischem_merged['RSP'].isnull()
        df_dischem_missing_rsp = df_dischem_merged[missing_rsp_dischem]
        df_missing_2 = df_dischem_missing_rsp[['SKU No.','Article']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** Article Code, SMD Code, Description & RSP")
        st.markdown("**Retailer data column headings:** Article Code, Article, Store Name, SOH Qty & "+Units_Sold)
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_dischem_merged['Start Date'] = Date_Start

        # Add Total Amount column
        df_dischem_merged['Total Amt'] = df_dischem_merged['Sales Qty'] * df_dischem_merged['RSP']
        df_dischem_merged['Total Amt'] = df_dischem_merged['Total Amt'].astype(float).round(2)

        # Add column for retailer and SOH
        df_dischem_merged['Forecast Group'] = 'Dis-Chem'

        # Final df. Don't change these headings. Rather change the ones above
        final_df_dischem_sales = df_dischem_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_dischem_p = df_dischem_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_dischem_s = df_dischem_merged[['Store Name','Total Amt']]

        # Show final df
        total = final_df_dischem_sales['Total Amt'].sum()
        total_units = final_df_dischem_sales['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_dischem_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_dischem_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_dischem_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_dischem_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))
        st.write('**Final Dataframe:**')
        final_df_dischem_sales

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_dischem_sales), unsafe_allow_html=True)

    except:
        st.write('Check data') 

# Dis-Chem-Pharmacies

elif option == 'Dis-Chem-Pharmacies':
    try:
        Units_Sold = (Short_Date_Dict[Month] + ' ' + Year)
        
        # Get retailers map
        df_dischemp_retailers_map = df_map
        df_dischemp_retailers_map = df_dischemp_retailers_map.rename(columns={'Description': 'Product Description'})
        df_retailers_map_dischemp_final = df_dischemp_retailers_map[['Article Code','SMD Code','Product Description','RSP']]

        # Get retailer data
        df_dischemp_data = df_data

        # Merge with retailer map
        df_dischemp_merged = df_dischemp_data.merge(df_retailers_map_dischemp_final, how='left', on='Article Code')

        # Rename columns
        df_dischemp_merged = df_dischemp_merged.rename(columns={'Article Code': 'SKU No.'})
        df_dischemp_merged = df_dischemp_merged.rename(columns={'Oct 2020': 'Sales Qty'})
        df_dischemp_merged = df_dischemp_merged.rename(columns={'SMD Code': 'Product Code'})

        # Find missing data
        missing_model_dischemp = df_dischemp_merged['Product Code'].isnull()
        df_dischemp_missing_model = df_dischemp_merged[missing_model_dischemp]
        df_missing = df_dischemp_missing_model[['SKU No.','Article']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp_dischemp = df_dischemp_merged['RSP'].isnull()
        df_dischemp_missing_rsp = df_dischemp_merged[missing_rsp_dischemp]
        df_missing_2 = df_dischemp_missing_rsp[['SKU No.','Article']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** Article Code, SMD Code, Description & RSP")
        st.markdown("**Retailer data column headings:** Article Code, Article, Store Name, SOH Qty & "+Units_Sold)
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_dischemp_merged['Start Date'] = Date_Start

        # Add Total Amount column
        df_dischemp_merged['Total Amt'] = df_dischemp_merged['Sales Qty'] * df_dischemp_merged['RSP']

        # Add column for retailer and SOH
        df_dischemp_merged['Forecast Group'] = 'Dis-Chem Pharmacies'

        # Final df. Don't change these headings. Rather change the ones above
        final_df_dischemp_sales = df_dischemp_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_dischemp_p = df_dischemp_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_dischemp_s = df_dischemp_merged[['Store Name','Total Amt']]

        # Show final df
        total = final_df_dischemp_sales['Total Amt'].sum()
        total_units = final_df_dischemp_sales['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_dischemp_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_dischemp_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_dischemp_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_dischemp_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))
        st.write('**Final Dataframe:**')        
        final_df_dischemp_sales

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_dischemp_sales), unsafe_allow_html=True)

    except:
        st.write('Check data') 

# HiFi Corp

elif option == 'HiFi':
    try:
        Units_Sold = ('Qty Sold '+ str(Month) + '.' + Year)

        # Get retailers map
        df_hifi_retailer_map = df_map
               

        # Get previous week
        hifi_data_prev = st.file_uploader('Previous week', type='xlsx')
        if hifi_data_prev:
            df_hifi_data_prev = pd.read_excel(hifi_data_prev)
        df_hifi_data_prev['Lookup'] = df_hifi_data_prev['Material'].astype(str) + df_hifi_data_prev['Plant']
        df_hifi_data_prev = df_hifi_data_prev.rename(columns={Units_Sold: 'Prev Sales'})
        df_hifi_data_prev = df_hifi_data_prev[['Lookup','Prev Sales']]

        # Get current week
        df_hifi_data = df_data
        df_hifi_data['Lookup'] = df_hifi_data['Material'].astype(str) + df_hifi_data['Plant']

        # Merge with retailer map and previous week
        df_hifi_data_merge_curr = df_hifi_data.merge(df_hifi_data_prev, how='left', on='Lookup')
        df_hifi_merged = df_hifi_data_merge_curr.merge(df_hifi_retailer_map, how='left', on='Material')

        missing_model_hifi = df_hifi_merged['SMD Code'].isnull()
        df_hifi_missing_model = df_hifi_merged[missing_model_hifi]
        df_missing = df_hifi_missing_model[['Material','Material Desc']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp_hifi = df_hifi_merged['RSP'].isnull()
        df_hifi_missing_rsp = df_hifi_merged[missing_rsp_hifi]
        df_missing_2 = df_hifi_missing_rsp[['Material','Material Desc']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** Material, SMD Code, Product Description & RSP")
        st.markdown("**Retailer data column headings:** Material, Material Desc, Plant, Plant Description, Total SOH Qty & "+Units_Sold)
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_hifi_merged['Start Date'] = Date_Start

        # Add Total Amount column
        df_hifi_merged['Sales Qty'] = df_hifi_merged[Units_Sold] - df_hifi_merged['Prev Sales']
        df_hifi_merged['Total Amt'] = df_hifi_merged['Sales Qty'] * df_hifi_merged['RSP']

        # Add column for retailer and SOH
        df_hifi_merged['Forecast Group'] = 'HIFI Corp'

        # Rename columns
        df_hifi_merged = df_hifi_merged.rename(columns={'Material': 'SKU No.'})
        df_hifi_merged = df_hifi_merged.rename(columns={'Total SOH Qty': 'SOH Qty'})
        df_hifi_merged = df_hifi_merged.rename(columns={'SMD Code': 'Product Code'})
        df_hifi_merged = df_hifi_merged.rename(columns={'Plant Description': 'Store Name'})

        # Final df. Don't change these headings. Rather change the ones above
        final_df_hifi_sales = df_hifi_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_hifi_p = df_hifi_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_hifi_s = df_hifi_merged[['Store Name','Total Amt']]

        # Show final df
        total = final_df_hifi_sales['Total Amt'].sum()
        total_units = final_df_hifi_sales['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_hifi_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_hifi_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_hifi_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_hifi_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))
        st.write('**Final Dataframe:**')          
        final_df_hifi_sales

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_hifi_sales), unsafe_allow_html=True)

    except:
        st.write('Check data')

# House and Home
elif option == 'H&H':

    try:
        # Get retailers map
        df_hh_retailers_map = df_map
        df_hh_retailers_map_final = df_hh_retailers_map[['SKU Number','SMD Product Code','SMD Description']]

        # Get previous week
        hh_data_prev = st.file_uploader('Previous week', type='xlsx')
        if hh_data_prev:
            df_hh_data_prev = pd.read_excel(hh_data_prev)
        df_hh_data_prev['Lookup'] = df_hh_data_prev['SKU Number'].astype(str) + df_hh_data_prev['Brn No'].astype(str)
        df_hh_data_prev = df_hh_data_prev.rename(columns={'Qty Sold': 'Prev Qty'})
        df_hh_data_prev = df_hh_data_prev.rename(columns={'Sold RSP': 'Prev Amt'})
        df_hh_data_prev_final = df_hh_data_prev[['Lookup','Prev Qty','Prev Amt']]

        # Get current week
        df_hh_data = df_data
        df_hh_data['Lookup'] = df_hh_data['SKU Number'].astype(str) + df_hh_data['Brn No'].astype(str)

        # Merge with retailer map and previous week
        df_hh_data_merge_curr = df_hh_data.merge(df_hh_data_prev_final, how='left', on='Lookup')
        df_hh_merged = df_hh_data_merge_curr.merge(df_hh_retailers_map_final, how='left', on='SKU Number')

        # Find missing data
        missing_model_hh = df_hh_merged['SMD Product Code'].isnull()
        df_hh_missing_model = df_hh_merged[missing_model_hh]
        df_missing = df_hh_missing_model[['SKU Number','SKU Description']]
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing)

    except:
        st.markdown("**Retailer map column headings:** SKU Number, SMD Product Code & SMD Description")
        st.markdown("**Retailer data column headings:** Brn No, Brn Description, SKU Number, SKU Description, Qty Sold, Sold RSP, Qty On Hand")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_hh_merged['Start Date'] = Date_Start

        # Add Total Amount column
        df_hh_merged['Sales Qty'] = df_hh_merged['Qty Sold'] - df_hh_merged['Prev Qty']
        df_hh_merged['Total Amt'] = (df_hh_merged['Sold RSP'] - df_hh_merged['Prev Amt'])*1.15

        # Add column for retailer and SOH
        df_hh_merged['Forecast Group'] = 'House and Home'
        df_hh_merged['Store Name'] = df_hh_merged['Brn Description'].str.title()

        # Rename columns
        df_hh_merged = df_hh_merged.rename(columns={'SKU Number': 'SKU No.'})
        df_hh_merged = df_hh_merged.rename(columns={'Qty On Hand': 'SOH Qty'})
        df_hh_merged = df_hh_merged.rename(columns={'SMD Product Code': 'Product Code'})
        df_hh_merged = df_hh_merged.rename(columns={'SMD Description': 'Product Description'})


        # Final df. Don't change these headings. Rather change the ones above
        final_df_hh_sales = df_hh_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_hh_p = df_hh_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_hh_s = df_hh_merged[['Store Name','Total Amt']]   

        # Show final df
        total = final_df_hh_sales['Total Amt'].sum()
        total_units = final_df_hh_sales['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_hh_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_hh_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_hh_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_hh_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))
        st.write('**Final Dataframe:**')          
        final_df_hh_sales

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_hh_sales), unsafe_allow_html=True)
    except:
        st.write('Check data')


# Incredible Connection
elif option == 'Incredible-Connection':
    try:
        Units_Sold = ('Qty Sold '+ str(Month) + '.' + Year)

        # Get retailers map
        df_ic_retailers_map = df_map
        

        # Get previous week
        ic_data_prev = st.file_uploader('Previous week', type='xlsx')
        if ic_data_prev:
            df_ic_data_prev = pd.read_excel(ic_data_prev)
        df_ic_data_prev['Lookup'] = df_ic_data_prev['Article'].astype(str) + df_ic_data_prev['Site']
        df_ic_data_prev = df_ic_data_prev.rename(columns={Units_Sold: 'Prev Sales'})
        df_ic_data_prev_final = df_ic_data_prev[['Lookup','Prev Sales']]

        # Get current week
        df_ic_data = df_data
        df_ic_data['Lookup'] = df_ic_data['Article'].astype(str) + df_ic_data['Site']

        # Rename columns
        df_ic_retailers_map = df_ic_retailers_map.rename(columns={'RRP': 'RSP'})

        # Merge with retailer map and previous week
        df_ic_data_merge_curr = df_ic_data.merge(df_ic_data_prev_final, how='left', on='Lookup')
        df_ic_merged = df_ic_data_merge_curr.merge(df_ic_retailers_map, how='left', on='Article')

        missing_model_ic = df_ic_merged['SMD Code'].isnull()
        df_ic_missing_model = df_ic_merged[missing_model_ic]
        df_missing = df_ic_missing_model[['Article','Article Name']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp_ic = df_ic_merged['RSP'].isnull()
        df_ic_missing_rsp = df_ic_merged[missing_rsp_ic]
        df_missing_2 = df_ic_missing_rsp[['Article','Article Name']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** Article, SMD Code, Product Description & RRP")
        st.markdown("**Retailer data column headings:** Article, Article Name, Site, Site Name, Total SOH Qty & "+Units_Sold)
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_ic_merged['Start Date'] = Date_Start

        # Add Total Amount column
        df_ic_merged['Sales Qty'] = df_ic_merged[Units_Sold] - df_ic_merged['Prev Sales']
        df_ic_merged['Total Amt'] = df_ic_merged['Sales Qty'] * df_ic_merged['RSP']

        # Add column for retailer and SOH
        df_ic_merged['Forecast Group'] = 'Incredible Connection'

        # Rename columns
        df_ic_merged = df_ic_merged.rename(columns={'Article': 'SKU No.'})
        df_ic_merged = df_ic_merged.rename(columns={'Total SOH Qty': 'SOH Qty'})
        df_ic_merged = df_ic_merged.rename(columns={'SMD Code': 'Product Code'})
        df_ic_merged = df_ic_merged.rename(columns={'Site Name': 'Store Name'})

        # Final df. Don't change these headings. Rather change the ones above
        final_df_ic_sales = df_ic_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_ic_p = df_ic_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_ic_s = df_ic_merged[['Store Name','Total Amt']]

        # Show final df
        total = final_df_ic_sales['Total Amt'].sum()
        total_units = final_df_ic_sales['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_ic_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_ic_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_ic_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_ic_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))
        st.write('**Final Dataframe:**')    
        final_df_ic_sales

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_ic_sales), unsafe_allow_html=True)

    except:
        st.write('Check data')


# Makro

elif option == 'Makro':
    Week = st.number_input("Enter week number: ")
    weekly_sales = Week+'-'+Year
    makro_stores = st.file_uploader('Stores', type='xlsx')
    if makro_stores:
        df_makro_stores = pd.read_excel(makro_stores)
   
    try:
        # Get retailers map
        df_makro_retailers_map = df_map
        df_makro_retailers_map = df_makro_retailers_map.rename(columns={'SMD Description': 'Product Description'})
        df_retailers_map_makro_final = df_makro_retailers_map[['Article','SMD Product Code','Product Description']]

        # Get retailer data
        df_makro_data = df_data
        df_makro_data = df_makro_data.rename(columns={'Incl SP': 'RSP'})

        # Merge with retailer map 
        df_makro_merged = df_makro_data.merge(df_retailers_map_makro_final, how='left', on='Article')

        # Merge with stores
        df_makro_merged = df_makro_merged.merge(df_makro_stores, how='left', on='Site')
        
        # Find missing data
        missing_model_makro = df_makro_merged['SMD Product Code'].isnull()
        df_makro_missing_model = df_makro_merged[missing_model_makro]
        df_missing = df_makro_missing_model[['Article','Article Desc']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp_makro = df_makro_merged['RSP'].isnull()
        df_makro_missing_rsp = df_makro_merged[missing_rsp_makro]  
        df_missing_2 = df_makro_missing_rsp[['Article','Article Desc']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** Article, SMD Product Code, SMD Description")
        st.markdown("**Retailer data column headings:** Article, Article Desc, Site, Store Name (in Stores.xlsx), SOH, "+weekly_sales)
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_makro_merged['Start Date'] = Date_Start

        # Total amount column
        df_makro_merged['Total Amt'] = df_makro_merged[weekly_sales] * df_makro_merged['RSP']
        
        # Add retailer column
        df_makro_merged['Forecast Group'] = 'Makro'

        # Rename columns
        df_makro_merged = df_makro_merged.rename(columns={'Article': 'SKU No.'})
        df_makro_merged = df_makro_merged.rename(columns={'SMD Product Code': 'Product Code'})
        df_makro_merged = df_makro_merged.rename(columns={'SOH': 'SOH Qty'})
        df_makro_merged = df_makro_merged.rename(columns={weekly_sales: 'Sales Qty'})

        # Don't change these headings. Rather change the ones above
        final_df_makro = df_makro_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_makro_p = df_makro_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_makro_s = df_makro_merged[['Store Name','Total Amt']]

        # Show final df
        total = final_df_makro['Total Amt'].sum()
        total_units = final_df_makro['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_makro_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_makro_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_makro_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_makro_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))
        st.write('**Final Dataframe:**')          
        final_df_makro

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_makro), unsafe_allow_html=True)

    except:
        st.write('Check data')

# Musica
elif option == 'Musica':
    try:
        # Get retailers map
        df_musica_retailers_map = df_map
        df_musica_retailers_map = df_musica_retailers_map.rename(columns={'SMD Desc': 'Product Description'})
        df_retailers_map_musica_final = df_musica_retailers_map[['Musica Code','SMD code','Product Description','RSP']]

        # Get retailer data
        df_musica_data = df_data
        df_musica_data = df_musica_data.rename(columns={'SKU No.': 'Musica Code'})
        df_musica_data = df_musica_data.rename(columns={'Sales.Qty': 'Sales Qty'})  
        #Merge with retailer map
        df_musica_merged = df_musica_data.merge(df_retailers_map_musica_final, how='left', on='Musica Code')    
        # Find missing data
        missing_model = df_musica_merged['SMD code'].isnull()
        df_musica_missing_model = df_musica_merged[missing_model]
        df_missing = df_musica_missing_model[['Musica Code','Title Desc']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp = df_musica_merged['RSP'].isnull()
        df_musica_missing_rsp = df_musica_merged[missing_rsp]
        df_missing_2 = df_musica_missing_rsp[['Musica Code','Title Desc']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** Musica Code, SMD code, SMD Desc, RSP")
        st.markdown("**Retailer data column headings:** Store Name, SKU No., Title Desc, Sales.Qty, SOH Qty")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_musica_merged['Start Date'] = Date_Start
   
        # Total amount column
        df_musica_merged['Total Amt'] = df_musica_merged['Sales Qty'] * df_musica_merged['RSP']

        # Add retailer column
        df_musica_merged['Forecast Group'] = 'Musica'

        # Rename columns
        df_musica_merged = df_musica_merged.rename(columns={'Musica Code': 'SKU No.'})
        df_musica_merged = df_musica_merged.rename(columns={'SMD code': 'Product Code'})

        # Don't change these headings. Rather change the ones above
        final_df_musica = df_musica_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_musica_p = df_musica_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_musica_s = df_musica_merged[['Store Name','Total Amt']]        

        # Show final df
        total = final_df_musica['Total Amt'].sum()
        total_units = final_df_musica['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_musica_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_musica_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_musica_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_musica_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))
        st.write('**Final Dataframe:**')          
        final_df_musica

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_musica), unsafe_allow_html=True)
    except:
        st.write('Check data')

# Ok Furniture
elif option == 'Ok-Furniture':

    try:
        # Get retailers map
        df_okf_retailers_map = df_map
        df_okf_retailers_map_final = df_okf_retailers_map[['SKU Number','SMD Product Code','SMD Description']]

        # Get previous week
        okf_data_prev = st.file_uploader('Previous week', type='xlsx')
        if okf_data_prev:
            df_okf_data_prev = pd.read_excel(okf_data_prev)
        df_okf_data_prev['Lookup'] = df_okf_data_prev['SKU Number'].astype(str) + df_okf_data_prev['Brn No'].astype(str)
        df_okf_data_prev = df_okf_data_prev.rename(columns={'Qty Sold': 'Prev Qty'})
        df_okf_data_prev = df_okf_data_prev.rename(columns={'Sold RSP': 'Prev Amt'})
        df_okf_data_prev_final = df_okf_data_prev[['Lookup','Prev Qty','Prev Amt']]

        # Get current week
        df_okf_data = df_data
        df_okf_data['Lookup'] = df_okf_data['SKU Number'].astype(str) + df_okf_data['Brn No'].astype(str)

        # Merge with retailer map and previous week
        df_okf_data_merge_curr = df_okf_data.merge(df_okf_data_prev_final, how='left', on='Lookup')
        df_okf_merged = df_okf_data_merge_curr.merge(df_okf_retailers_map_final, how='left', on='SKU Number')

        # Find missing data
        missing_model_okf = df_okf_merged['SMD Product Code'].isnull()
        df_okf_missing_model = df_okf_merged[missing_model_okf]
        df_missing = df_okf_missing_model[['SKU Number','SKU Description']]
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing)

    except:
        st.markdown("**Retailer map column headings:** SKU Number, SMD Product Code & SMD Description")
        st.markdown("**Retailer data column headings:** Brn No, Brn Description, SKU Number, SKU Description, Qty Sold, Sold RSP, Qty On Hand")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_okf_merged['Start Date'] = Date_Start

        # Add Total Amount column
        df_okf_merged['Sales Qty'] = df_okf_merged['Qty Sold'] - df_okf_merged['Prev Qty']
        df_okf_merged['Total Amt'] = (df_okf_merged['Sold RSP'] - df_okf_merged['Prev Amt'])*1.15

        # Add column for retailer and SOH
        df_okf_merged['Forecast Group'] = 'OK Furniture'
        df_okf_merged['Store Name'] = df_okf_merged['Brn Description'].str.title()

        # Rename columns
        df_okf_merged = df_okf_merged.rename(columns={'SKU Number': 'SKU No.'})
        df_okf_merged = df_okf_merged.rename(columns={'Qty On Hand': 'SOH Qty'})
        df_okf_merged = df_okf_merged.rename(columns={'SMD Product Code': 'Product Code'})
        df_okf_merged = df_okf_merged.rename(columns={'SMD Description': 'Product Description'})


        # Final df. Don't change these headings. Rather change the ones above
        final_df_ok_sales = df_okf_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_okf_p = df_okf_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_okf_s = df_okf_merged[['Store Name','Total Amt']]   

        # Show final df
        total = final_df_ok_sales['Total Amt'].sum()
        total_units = final_df_ok_sales['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_okf_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_okf_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_okf_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_okf_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))
        st.write('**Final Dataframe:**')          
        final_df_ok_sales

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_ok_sales), unsafe_allow_html=True)
    except:
        st.write('Check data')


# Outdoor Warehouse
elif option == 'Outdoor-Warehouse':

    st.markdown("**Stock on hand needs to be in a separate sheet**")

    ow_soh = st.file_uploader('SOH', type='xlsx')
    if ow_soh:
        df_ow_soh = pd.read_excel(ow_soh)

    try:
        # Get retailers map
        df_ow_retailers_map = df_map
        df_ow_retailers_map = df_ow_retailers_map.rename(columns={'SKUCode': 'Article Code'})
        df_ow_retailers_map = df_ow_retailers_map.rename(columns={'SMD Desc': 'Product Description'})
        df_retailers_map_ow_final = df_ow_retailers_map[['Article Code','SMD Code','Product Description','RSP']]

        # Get retailer data
        df_ow_data = df_data
        df_ow_data = df_ow_data.iloc[1:]

        # Get rid of extra columns
        del df_ow_data['Code']
        del df_ow_data['Size']
        del df_ow_data['Colour']
        del df_ow_data['Total']

        # Melt data
        df_ow_data = pd.melt(df_ow_data, id_vars=['Product', 'SKUCode'])

        # Rename columns
        df_ow_data = df_ow_data.rename(columns={'variable': 'Store Name'})
        df_ow_data = df_ow_data.rename(columns={'value': 'Sales Qty'})
        df_ow_data = df_ow_data.rename(columns={'SKUCode': 'Article Code'})

        # Get rid of commas
        df_ow_data['Sales Qty'] = df_ow_data['Sales Qty'].replace(',','', regex=True)
        df_ow_data['Sales Qty'] = df_ow_data['Sales Qty'].astype(float)

        # Lookup column
        df_ow_data['Lookup'] = df_ow_data['Article Code'].astype(str) + df_ow_data['Store Name']

        # Get stock on hand
        df_ow_soh = df_ow_soh.iloc[1:]
        del df_ow_soh['Code']
        del df_ow_soh['Size']
        del df_ow_soh['Colour']
        del df_ow_soh['Total']
        df_ow_soh = pd.melt(df_ow_soh, id_vars=['Product', 'SKUCode'])
        df_ow_soh = df_ow_soh.rename(columns={'variable': 'Store Name'})
        df_ow_soh = df_ow_soh.rename(columns={'value': 'SOH Qty'})
        df_ow_soh['SOH Qty'] = df_ow_soh['SOH Qty'].replace(',','', regex=True)
        df_ow_soh['SOH Qty'] = df_ow_soh['SOH Qty'].astype(float)
        df_ow_soh['Lookup'] = df_ow_soh['SKUCode'].astype(str) + df_ow_soh['Store Name']
        df_ow_soh_final = df_ow_soh[['Lookup','SOH Qty']]

        # Merge with SOH
        df_ow_data = df_ow_data.merge(df_ow_soh_final, how='left', on='Lookup')

        # Merge with retailer map
        df_ow_merged = df_ow_data.merge(df_retailers_map_ow_final, how='left', on='Article Code')

        # Rename columns
        df_ow_merged = df_ow_merged.rename(columns={'Article Code': 'SKU No.'})
        df_ow_merged = df_ow_merged.rename(columns={'SMD Code': 'Product Code'})

        # Find missing data
        missing_model = df_ow_merged['Product Code'].isnull()
        df_ow_missing_model = df_ow_merged[missing_model]
        df_missing = df_ow_missing_model[['SKU No.','Product']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp = df_ow_merged['RSP'].isnull()
        df_ow_missing_rsp = df_ow_merged[missing_rsp]
        df_missing_2 = df_ow_missing_rsp[['SKU No.','Product']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** Article Code, SMD Code, SMD Desc ,RSP")
        st.markdown("**Retailer data column headings:** Code, Product, SKUCode")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")


    try:
        # Set date columns
        df_ow_merged['Start Date'] = Date_Start

        # Total amount column
        df_ow_merged['Total Amt'] = df_ow_merged['Sales Qty'] * df_ow_merged['RSP']

        # Add retailer and store column
        df_ow_merged['Forecast Group'] = 'Outdoor Warehouse'

        # Don't change these headings. Rather change the ones above
        final_df_ow = df_ow_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_ow_p = df_ow_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_ow_s = df_ow_merged[['Store Name','Total Amt']]    

        # Show final df
        total = final_df_ow['Total Amt'].sum()
        total_units = final_df_ow['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_ow_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_ow_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_ow_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_ow_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))
        st.write('**Final Dataframe:**')           
        final_df_ow

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_ow), unsafe_allow_html=True)
    except:
        st.write('Check data') 

#Pep Africa
elif option == 'Pep-Africa':
      
    try:
        Wk = int(st.text_input("Enter week number: "))
        Wk_sales = 'Wk ' + str(Wk)

        # Get retailers map
        df_pepaf_retailers_map = df_map

        # Get retailer data
        df_pepaf_data = df_data
        df_pepaf_data.columns = df_pepaf_data.iloc[1]
        df_pepaf_data = df_pepaf_data.iloc[2:]
        df_pepaf_data = df_pepaf_data.rename(columns={'Style Code': 'SKU No.'})
        df_pepaf_data['Store Name'] = df_pepaf_data['Country Code'].map(Country_Dict)
        df_pepaf_data = df_pepaf_data.rename(columns={Wk_sales: 'Sales Qty'})
        df_pepaf_data = df_pepaf_data.rename(columns={'WSOH': 'SOH Qty'})
        
        # Merge with retailer map
        df_pepaf_merged = df_pepaf_data.merge(df_pepaf_retailers_map, how='left', on='SKU No.')
        
        # Find missing data
        missing_model = df_pepaf_merged['Product Code'].isnull()
        df_pepaf_missing_model = df_pepaf_merged[missing_model]
        df_missing = df_pepaf_missing_model[['SKU No.','Style Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ") 
        missing_rsp = df_pepaf_merged['RSP'].isnull()
        df_pepaf_missing_rsp = df_pepaf_merged[missing_rsp]
        df_missing_2 = df_pepaf_missing_rsp[['SKU No.','Style Description']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** SKU No., Product Code, Product Description, RSP")
        st.markdown("**Retailer data column headings:** Country Code, Style Code, Style Description, WSOH")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_pepaf_merged['Start Date'] = Date_Start

        # Total amount column
        df_pepaf_merged['Total Amt'] = df_pepaf_merged['Sales Qty'] * df_pepaf_merged['RSP']
        df_pepaf_merged['Total Amt'] = df_pepaf_merged['Total Amt'].apply(lambda x: round(x,2))

        # Add retailer column
        df_pepaf_merged['Forecast Group'] = 'Pep Africa'

        # Don't change these headings. Rather change the ones above
        final_df_pepaf = df_pepaf_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_pepaf_p = df_pepaf_merged[['Product Code','Style Description', 'Sales Qty','Total Amt']]
        final_df_pepaf_s = df_pepaf_merged[['Store Name','Total Amt']]   

        # Show final df
        total = final_df_pepaf['Total Amt'].sum()
        total_units = final_df_pepaf['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_pepaf_p.groupby("Style Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_pepaf_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_pepaf_p.groupby("Style Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_pepaf_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))
        st.write('**Final Dataframe:**')  
        final_df_pepaf

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_pepaf), unsafe_allow_html=True)
    except:
        st.write('Check data') 


#Pep South Africa
elif option == 'Pep-SA':
      
    try:
        Wk = int(st.text_input("Enter week number: "))

        # Get retailers map
        df_pep_retailers_map = df_map

        # Get retailer data
        df_pep_data = df_data
        df_pep_data['SKU Number'] = df_pep_data.apply(lambda x: 'Wk '+ str(x['Unnamed: 1']) if x['Unnamed: 1'] == Wk else x['SKU Number'], axis = 1)
        
        # Get rid of extra columns
        del df_pep_data['Accessories']
        del df_pep_data['Accessories.1']
        del df_pep_data['Accessories.2']
        del df_pep_data['Accessories.3']
        del df_pep_data['Total']
        del df_pep_data['Total.1']

        # Rename trash then delete trash
        df_pep_data = df_pep_data.rename(columns={df_pep_data.filter(regex='Unnamed:*').columns[0]:'Unnamed'})
        del df_pep_data['Unnamed']

        df_pep_data = df_pep_data.rename(columns={df_pep_data.filter(regex='Unnamed:*').columns[0]:'Unnamed'})
        del df_pep_data['Unnamed']

        df_pep_data = df_pep_data.rename(columns={df_pep_data.filter(regex='Unnamed:*').columns[0]:'Unnamed'})
        del df_pep_data['Unnamed']

        # Transpose data
        df_pep_data = df_pep_data.T

        # Get column headings
        df_pep_data.columns = df_pep_data.iloc[0]
        df_pep_data = df_pep_data.iloc[1:]


        # Rename columns
        df_pep_data = df_pep_data.rename(columns={'Month': 'Description'})



        # Merge with retailer map
        df_pep_merged = df_pep_data.merge(df_pep_retailers_map, how='left', on='Style Code')

        # Rename columns
        df_pep_merged = df_pep_merged.rename(columns={'Style Code': 'SKU No.'})
        df_pep_merged = df_pep_merged.rename(columns={'Total Company Stock': 'SOH Qty'})
        df_pep_merged = df_pep_merged.rename(columns={'Wk '+str(Wk): 'Sales Qty'})
        
        # Find missing data
        missing_model = df_pep_merged['Product Code'].isnull()
        df_pep_missing_model = df_pep_merged[missing_model]
        df_missing = df_pep_missing_model[['SKU No.','Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ") 
        missing_rsp = df_pep_merged['RSP'].isnull()
        df_pep_missing_rsp = df_pep_merged[missing_rsp]
        df_missing_2 = df_pep_missing_rsp[['SKU No.','Description']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** Style Code, Product Code, Product Description, RSP")
        st.markdown("**Retailer data column headings:** Style Code, Month, Total Company Stock")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_pep_merged['Start Date'] = Date_Start

        # Total amount column
        df_pep_merged['Total Amt'] = df_pep_merged['Sales Qty'] * df_pep_merged['RSP']
        df_pep_merged['Total Amt'] = df_pep_merged['Total Amt'].apply(lambda x: round(x,2))

        # Add retailer and store column
        df_pep_merged['Forecast Group'] = 'Pep South Africa'
        df_pep_merged['Store Name'] = ''

        # Don't change these headings. Rather change the ones above
        final_df_pep = df_pep_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_pep_p = df_pep_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_pep_s = df_pep_merged[['Store Name','Total Amt']]   

        # Show final df
        total = final_df_pep['Total Amt'].sum()
        total_units = final_df_pep['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_pep_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_pep_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_pep_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_pep_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))
        st.write('**Final Dataframe:**')  
        final_df_pep

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_pep), unsafe_allow_html=True)
    except:
        st.write('Check data') 

# Pick n Pay
elif option == 'PnP':

    pnp_soh = st.file_uploader('SOH', type='xlsx')
    if pnp_soh:
        df_pnp_soh = pd.read_excel(pnp_soh)

    try:
        # Get retailers map
        df_pnp_retailers_map = df_map
        df_retailers_map_pnp_final = df_pnp_retailers_map[['Article Number','SMD code','Product Description', 'RSP']]
        
        # Get retailer data
        df_pnp_data = df_data
        df_pnp_data = df_pnp_data.rename(columns={'PnP ArticleNumber': 'Article Number'})
        df_pnp_data = df_pnp_data.rename(columns={'Product Description': 'Article Desc'})
        df_pnp_data = df_pnp_data.rename(columns={'Store': 'Store Name'})

        # Lookup column
        df_pnp_data['Lookup'] = df_pnp_data['Article Number'].astype(str) + df_pnp_data['Store ID']

        # Get stock on hand
        df_pnp_soh['Lookup'] = df_pnp_soh['Article Number'].astype(str) + df_pnp_soh['Site Code']
        df_pnp_soh_final = df_pnp_soh[['Lookup','SOH Qty']]

        # Merge with SOH
        df_pnp_data = df_pnp_data.merge(df_pnp_soh_final, how='left', on='Lookup')

        # Merge with retailer map
        df_pnp_merged = df_pnp_data.merge(df_retailers_map_pnp_final, how='left', on='Article Number')

        # Rename columns
        df_pnp_merged = df_pnp_merged.rename(columns={'Article Number': 'SKU No.'})
        df_pnp_merged = df_pnp_merged.rename(columns={'SMD code': 'Product Code'})
        df_pnp_merged = df_pnp_merged.rename(columns={'Units': 'Sales Qty'})

        # Find missing data
        missing_model = df_pnp_merged['Product Code'].isnull()
        df_pnp_missing_model = df_pnp_merged[missing_model]
        df_missing = df_pnp_missing_model[['SKU No.','Article Desc']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ") 
        missing_rsp = df_pnp_merged['RSP'].isnull()
        df_pnp_missing_rsp = df_pnp_merged[missing_rsp] 
        df_missing_2 = df_pnp_missing_rsp[['SKU No.','Article Desc']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** Article Number, SMD code, Product Description, RSP")
        st.markdown("**Retailer data column headings:** Product Description, Store ID, Store, Units, PnP ArticleNumber")
        st.markdown("**Retailer SOH column headings:** Site Code, Article Number, SOH Qty")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_pnp_merged['Start Date'] = Date_Start

        # Total amount column
        df_pnp_merged['Total Amt'] = df_pnp_merged['Sales Qty'] * df_pnp_merged['RSP']

        # Add retailer and store column
        df_pnp_merged['Forecast Group'] = 'Pick n Pay'

        # Don't change these headings. Rather change the ones above
        final_df_pnp = df_pnp_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_pnp_p = df_pnp_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_pnp_s = df_pnp_merged[['Store Name','Total Amt']]  

        # Show final df
        total = final_df_pnp['Total Amt'].sum()
        total_units = final_df_pnp['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_pnp_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_pnp_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_pnp_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_pnp_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))
        st.write('**Final Dataframe:**')          
        final_df_pnp

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_pnp), unsafe_allow_html=True)
    except:
        st.write('Check data') 


# Sportsmans Warehouse
elif option == 'Sportsmans-Warehouse':

    st.markdown("**Stock on hand needs to be in a separate sheet**")
    st.markdown("**Please delete the size column in both data sheets**")

    sw_soh = st.file_uploader('SOH', type='xlsx')
    if sw_soh:
        df_sw_soh = pd.read_excel(sw_soh)

    try:
        # Get retailers map
        df_sw_retailers_map = df_map
        df_sw_retailers_map = df_sw_retailers_map.rename(columns={'SKUCode': 'Article Code'})
        df_sw_retailers_map = df_sw_retailers_map.rename(columns={'Description': 'Product Description'})
        df_retailers_map_sw_final = df_sw_retailers_map[['Article Code','SMD Code', 'Product Description', 'RSP']]

        # Get retailer data
        df_sw_data = df_data
        df_sw_data = df_sw_data.iloc[1:]

        # Get rid of extra columns
        del df_sw_data['Code']
        del df_sw_data['Colour']
        del df_sw_data['Total']

        # Melt data
        df_sw_data = pd.melt(df_sw_data, id_vars=['Product', 'SKUCode'])

        # Rename columns
        df_sw_data = df_sw_data.rename(columns={'variable': 'Store Name'})
        df_sw_data = df_sw_data.rename(columns={'value': 'Sales Qty'})
        df_sw_data = df_sw_data.rename(columns={'SKUCode': 'Article Code'})

        # Get rid of commas
        df_sw_data['Sales Qty'] = df_sw_data['Sales Qty'].replace(',','', regex=True)
        df_sw_data['Sales Qty'] = df_sw_data['Sales Qty'].astype(float)

        # Lookup column
        df_sw_data['Lookup'] = df_sw_data['Article Code'].astype(str) + df_sw_data['Store Name']

        # Get stock on hand
        df_sw_soh = df_sw_soh.iloc[1:]
        del df_sw_soh['Code']
        del df_sw_soh['Colour']
        del df_sw_soh['Total']
        df_sw_soh = pd.melt(df_sw_soh, id_vars=['Product', 'SKUCode'])
        df_sw_soh = df_sw_soh.rename(columns={'variable': 'Store Name'})
        df_sw_soh = df_sw_soh.rename(columns={'value': 'SOH Qty'})
        df_sw_soh['SOH Qty'] = df_sw_soh['SOH Qty'].replace(',','', regex=True)
        df_sw_soh['SOH Qty'] = df_sw_soh['SOH Qty'].astype(float)
        df_sw_soh['Lookup'] = df_sw_soh['SKUCode'].astype(str) + df_sw_soh['Store Name']
        df_sw_soh_final = df_sw_soh[['Lookup','SOH Qty']]

        # Merge with SOH
        df_sw_data = df_sw_data.merge(df_sw_soh_final, how='left', on='Lookup')

        # Merge with retailer map
        df_sw_merged = df_sw_data.merge(df_retailers_map_sw_final, how='left', on='Article Code')

        # Rename columns
        df_sw_merged = df_sw_merged.rename(columns={'Article Code': 'SKU No.'})
        df_sw_merged = df_sw_merged.rename(columns={'SMD Code': 'Product Code'})

        # Find missing data
        missing_model = df_sw_merged['Product Code'].isnull()
        df_sw_missing_model = df_sw_merged[missing_model]
        df_missing = df_sw_missing_model[['SKU No.','Product']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")    
        missing_rsp = df_sw_merged['RSP'].isnull()
        df_sw_missing_rsp = df_sw_merged[missing_rsp]
        df_missing_2 = df_sw_missing_rsp[['SKU No.','Product']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** SKUCode, SMD Code, Description, RSP")
        st.markdown("**Retailer data column headings:** Code, Product, SKUCode")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_sw_merged['Start Date'] = Date_Start

        # Total amount column
        df_sw_merged['Total Amt'] = df_sw_merged['Sales Qty'] * df_sw_merged['RSP']

        # Add retailer and store column
        df_sw_merged['Forecast Group'] = 'Sportsmans Warehouse'

        # Don't change these headings. Rather change the ones above
        final_df_sw = df_sw_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_sw_p = df_sw_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_sw_s = df_sw_merged[['Store Name','Total Amt']]  

        # Show final df
        total = final_df_sw['Total Amt'].sum()
        total_units = final_df_sw['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_sw_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_sw_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_sw_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_sw_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))
        st.write('**Final Dataframe:**')  
        final_df_sw

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_sw), unsafe_allow_html=True)
    except:
        st.write('Check data') 

# Takealot
elif option == 'Takealot':
    try:
        # Get retailers map
        df_takealot_retailers_map = df_map
        df_takealot_retailers_map = df_takealot_retailers_map.rename(columns={'Description': 'Product Description'})
        df_retailers_map_takealot_final = df_takealot_retailers_map[['idProduct','Product Description','Manufacturer','SMD Code','RSP']]
        # Get retailer data
        df_takealot_data = df_data
        df_takealot_data = df_takealot_data.iloc[1:]
        #Merge with retailer map
        df_takealot_merged = df_takealot_data.merge(df_retailers_map_takealot_final, how='left', on='idProduct')    
        # Find missing data
        missing_model = df_takealot_merged['SMD Code'].isnull()
        df_takealot_missing_model = df_takealot_merged[missing_model]
        df_missing = df_takealot_missing_model[['idProduct','Supplier Code']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp = df_takealot_merged['RSP'].isnull()
        df_takealot_missing_rsp = df_takealot_merged[missing_rsp]
        df_missing_2 = df_takealot_missing_rsp[['idProduct','Supplier Code']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** idProduct, SMD Code, RSP")
        st.markdown("**Retailer data column headings:** idProduct, Supplier Code, Total SOH, Units Sold Qty")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_takealot_merged['Start Date'] = Date_Start

        # Total amount column
        df_takealot_merged['Total Amt'] = df_takealot_merged['Units Sold Value'] * 1.15

        # Add retailer and store column
        df_takealot_merged['Forecast Group'] = 'Takealot'
        df_takealot_merged['Store Name'] = ''

        # Rename columns
        df_takealot_merged = df_takealot_merged.rename(columns={'idProduct': 'SKU No.','Units Sold Qty' :'Sales Qty','Total SOH':'SOH Qty','SMD Code':'Product Code' })

        # Don't change these headings. Rather change the ones above
        final_df_takealot = df_takealot_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_takealot_p = df_takealot_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_takealot_s = df_takealot_merged[['Store Name','Total Amt']]  

        # Show final df
        total = final_df_takealot['Total Amt'].sum()
        total_units = final_df_takealot['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_takealot_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_takealot_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_takealot_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_takealot_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))
        st.write('**Final Dataframe:**')         
        final_df_takealot

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_takealot), unsafe_allow_html=True)
    except:
        st.write('Check data')


# TFG
elif option == 'TFG':
    try:
        # Get retailers map
        df_tfg_retailers_map = df_map
        df_tfg_retailers_map = df_tfg_retailers_map.rename(columns={'DES':'Product Description'})
        df_retailers_map_tfg_final = df_tfg_retailers_map[['Article Code','Code','Product Description','RSP']]
        
        # Get retailer data
        df_tfg_data = df_data
        # Apply the split string method on the Style code to get the SKU No. out
        df_tfg_data['Article Code'] = df_tfg_data['Style'].astype(str).str.split(' ').str[0]
        # Convert to float
        df_tfg_data['Article Code'] = df_tfg_data['Article Code'].astype(float)
        # Merge with retailer map 
        df_tfg_merged = df_tfg_data.merge(df_retailers_map_tfg_final, how='left', on='Article Code')
        
        # Find missing data
        missing_model_tfg = df_tfg_merged['Code'].isnull()
        df_tfg_missing_model = df_tfg_merged[missing_model_tfg]
        df_missing = df_tfg_missing_model[['Article Code','Style']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp_tfg = df_tfg_merged['RSP'].isnull()
        df_tfg_missing_rsp = df_tfg_merged[missing_rsp_tfg] 
        df_missing_2 = df_tfg_missing_rsp[['Article Code','Style']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** Article Code, Code, DES, RSP")
        st.markdown("**Retailer data column headings:** Style, Sls (U), CSOH Incl IT (U)")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_tfg_merged['Start Date'] = Date_Start

        # Rename columns
        df_tfg_merged = df_tfg_merged.rename(columns={'Article Code': 'SKU No.','Sls (U)' :'Sales Qty', 'CSOH Incl IT (U)':'SOH Qty', 'Code' : 'Product Code' })

        # Total amount column
        df_tfg_merged['Total Amt'] = df_tfg_merged['Sales Qty'] * df_tfg_merged['RSP']

        # Add retailer and store column
        df_tfg_merged['Forecast Group'] = 'TFG'
        df_tfg_merged['Store Name'] = ''

        # Don't change these headings. Rather change the ones above
        final_df_tfg = df_tfg_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_tfg_p = df_tfg_merged[['Product Code','Product Description','Total Amt']]
        final_df_tfg_s = df_tfg_merged[['Store Name','Total Amt']]

        # Show final df
        total = df_tfg_merged['Total Amt'].sum()
        total_units = final_df_tfg['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_tfg_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt)
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_tfg_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st)
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_tfg_p.groupby("Product Description").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb)
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_tfg_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb)

        st.write('**Final Dataframe:**')
        final_df_tfg

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(df_tfg_merged), unsafe_allow_html=True)
    except:
        st.write('Check data')

else:
    st.write('Retailer not selected yet')

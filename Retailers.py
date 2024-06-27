
# Import libraries
import streamlit as st
import numpy as np
import pandas as pd
import base64
from io import BytesIO
import io
import datetime as dt
# import locale
# locale.setlocale( locale.LC_ALL, 'en_ZA.ANSI' )
# st.set_page_config(layout="centered")

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

def df_stats(df,df_p,df_s):
        total = df['Total Amt'].sum()
        total_units = df['Sales Qty'].sum()
        st.write('**The total sales for the week are:** R',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = df_p.groupby(["Product Description"]).agg({"Sales Qty":"sum", "Total Amt":"sum"}).sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = df_s.groupby("Store Name").agg({"Total Amt":"sum"}).sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('R{0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = df_p.groupby("Product Description").agg({"Sales Qty":"sum", "Total Amt":"sum"}).sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'R{:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = df_s.groupby("Store Name").agg({"Total Amt":"sum"}).sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('R{0:,.2f}'))
        st.write('**Final Dataframe:**') 
        df 

st.title('Retailer Sales Reports - SA')

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
    ('Please select','Ackermans','Bradlows/Russels','Builders','Checkers',
    'Clicks', 'CNA', 'Cross_Trainer','Dealz', 'Decofurn','Dis-Chem','Dis-Chem-Pharmacies', 'Game', 'H&H','HiFi',
    'Incredible-Connection','J.A.M.','Makro', 'Mr-Price-Sport', 'Musica','Ok-Furniture', 'Ok-Furniture-Africa', 
    'Outdoor-Warehouse','Pep-Africa','Pep-SA','PnP','Retailability', 'Sportsmans-Warehouse','Takealot','TFG','TFG_Cosmetics','TRU'))
st.write('You selected:', option)

st.write("")
st.markdown("Please ensure data is in the **_first sheet_** of your Excel Workbook")

map_file = st.file_uploader('Retailer Map', type='xlsx')
if map_file:
    df_map = pd.read_excel(map_file)

data_file = st.file_uploader('Weekly Sales Data',type=['csv','txt','xlsx','xls'])
if data_file:    
    if data_file.name[-3:] == 'csv':
        data_file.seek(0)
        df_data = pd.read_csv(io.StringIO(data_file.read().decode('utf-8')), delimiter='|')
        try:
            df_data = df_data.rename(columns=lambda x: x.strip())
        except:
            df_data = df_data

    elif data_file.name[-3:] == 'txt':
        data_file.seek(0)
        df_data = pd.read_csv(io.StringIO(data_file.read().decode('utf-8')), delimiter='|')
        try:
            df_data = df_data.rename(columns=lambda x: x.strip())
        except:
            df_data = df_data

    else:
        df_data = pd.read_excel(data_file)
        try:
            df_data = df_data.rename(columns=lambda x: x.strip())
        except:
            df_data = df_data

# Ackermans
if option == 'Ackermans':

    if Date_End.month < 10:
        Month = '0'+str(Date_End.month)
    else:
        Month = str(Date_End.month)

    Units_Sold = 'Sales: ' + Year + '/' + str(Month) + '/' + Day
    CSOH = 'CSOH: ' + Year + '/' + str(Month) + '/' + Day



    try:
        # Get retailers map
        df_ackermans_retailers_map = df_map
        df_ackermans_retailers_map = df_ackermans_retailers_map.rename(columns={'Style Code': 'SKU No.'})
        df_ackermans_retailers_map_final = df_ackermans_retailers_map[['SKU No.','Product Description','SMD Product Code']]
        
        # Get retailer data
        df_ackermans_data = df_data
        df_ackermans_data['SKU No.'] = df_ackermans_data['Style'].astype(str).str.split(' ').str[0]
        
        # Merge with retailer map
        df_ackermans_data['SKU No.'] = df_ackermans_data['SKU No.'].astype(int)
        df_ackermans_merged = df_ackermans_data.merge(df_ackermans_retailers_map_final, how='left', on='SKU No.')

        # Find missing data
        missing_model_ackermans = df_ackermans_merged['SMD Product Code'].isnull()
        df_ackermans_missing_model = df_ackermans_merged[missing_model_ackermans]
        df_missing = df_ackermans_missing_model[['SKU No.','Style']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)
        st.write(" ")

    except:
        st.markdown("**Retailer map column headings:** Style Code, Product Description, SMD Product Code")
        st.markdown("**Retailer data column headings:** Store, Style, Closing Stock Units, Nett Sale Units, Nett Sale Value")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct") 

        
    try:
        # Set date columns
        df_ackermans_merged['Start Date'] = Date_End

        # Add retailer column and store column
        df_ackermans_merged['Forecast Group'] = 'Ackermans'

        # Rename columns
        df_ackermans_merged = df_ackermans_merged.rename(columns={'Closing Stock Units': 'SOH Qty'})
        df_ackermans_merged = df_ackermans_merged.rename(columns={'Nett Sale Units': 'Sales Qty'})
        df_ackermans_merged = df_ackermans_merged.rename(columns={'Nett Sale Value': 'Total Amt'})
        df_ackermans_merged = df_ackermans_merged.rename(columns={'SMD Product Code': 'Product Code'})
        df_ackermans_merged = df_ackermans_merged.rename(columns={'Store': 'Store Name'})

        # Don't change these headings. Rather change the ones above
        final_df_ackermans = df_ackermans_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_ackermans_p = df_ackermans_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_ackermans_s = df_ackermans_merged[['Store Name','Total Amt']]

        # Show final df
        df_stats(final_df_ackermans, final_df_ackermans_p, final_df_ackermans_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_ackermans), unsafe_allow_html=True)

    except:
        st.write('Check data')


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
        df_stats(final_df_br,final_df_br_p,final_df_br_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_br), unsafe_allow_html=True)

    except:
        st.write('Check data')

# Builders Warehouse

elif option == 'Builders':

    Week = st.number_input("Enter week number: ",min_value = 0, value = 0)
    if int(Week) < 10:
        Week = str(0) + str(Week)
    else:
        Week = str(Week)
    
    weekly_sales = Week +'-'+Year[-1:]
    
    
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
        df_bw_data = df_bw_data.rename(columns={'Incl SP': 'RSP'})
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
        st.markdown("**Please remove all spacing in headings!**")
        st.markdown("**Retailer map column headings:** Article, SMD Product Code")
        st.markdown("**Retailer data column headings:** Article, Article Description, Site, Store Name (in Stores.xlsx), SOH, "+weekly_sales)
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_bw_merged['Start Date'] = Date_Start

        # Total amount column
        df_bw_merged['Total Amt'] = df_bw_merged[weekly_sales].astype(float) * df_bw_merged['RSP'].astype(float)

        # Add retailer column
        df_bw_merged['Forecast Group'] = 'Builders Warehouse'

        # Rename columns
        df_bw_merged = df_bw_merged.rename(columns={'Article': 'SKU No.'})
        df_bw_merged = df_bw_merged.rename(columns={'SMD Product Code': 'Product Code'})
        df_bw_merged = df_bw_merged.rename(columns={'SOH': 'SOH Qty'})
        df_bw_merged = df_bw_merged.rename(columns={weekly_sales: 'Sales Qty'})

        # Don't change these headings. Rather change the ones above
        final_df_bw = df_bw_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_bw_p = df_bw_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_bw_s = df_bw_merged[['Store Name','Total Amt']]

        # Show final df
        df_stats(final_df_bw,final_df_bw_p,final_df_bw_s)

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
        df_checkers_soh = df_checkers_soh.rename(columns=lambda x: x.strip())
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
        df_stats(final_df_checkers_sales,final_df_checkers_p,final_df_checkers_s)

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
        df_clicks_merged['Total Amt'] = df_clicks_merged['Sales Value LW TY'] * 1.15

        # Add retailer column
        df_clicks_merged['Forecast Group'] = 'Clicks'

        # Rename columns
        df_clicks_merged = df_clicks_merged.rename(columns={'Clicks Product Number': 'SKU No.'})
        df_clicks_merged = df_clicks_merged.rename(columns={'SMD CODE': 'Product Code'})
        df_clicks_merged = df_clicks_merged.rename(columns={'Product Description' : 'Clicks Desc'})
        df_clicks_merged = df_clicks_merged.rename(columns={'SMD DESC': 'Product Description'})
        df_clicks_merged = df_clicks_merged.rename(columns={'Store Description': 'Store Name'})
        df_clicks_merged = df_clicks_merged.rename(columns={'Store Stock Qty': 'SOH Qty'})
        df_clicks_merged = df_clicks_merged.rename(columns={'Sales Qty LW TY': 'Sales Qty'})
        
        # Don't change these headings. Rather change the ones above
        final_df_clicks = df_clicks_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_clicks_p = df_clicks_merged[['Product Code','Product Description','Sales Qty', 'Total Amt']]
        final_df_clicks_s = df_clicks_merged[['Store Name','Total Amt']]
        
        # Show final df
        df_stats(final_df_clicks,final_df_clicks_p,final_df_clicks_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_clicks), unsafe_allow_html=True)

    except:
        st.write('Check data')

# CNA
elif option == 'CNA':

    st.markdown("**Stock on hand needs to be in a separate sheet**")

    cna_soh = st.file_uploader('SOH', type='xlsx')
    if cna_soh:
        df_cna_soh = pd.read_excel(cna_soh)

    try:
        # Get retailers map
        df_cna_retailers_map = df_map
        df_cna_retailers_map = df_cna_retailers_map.rename(columns={'NEW- Retailers Article code':'SKU No.', 'SMD Code': 'Product Code'})
        df_cna_retailers_map = df_cna_retailers_map[['SKU No.','Product Code']]

        # Get retailer data
        df_cna_data = df_data
        
        # Rename columns
        df_cna_data = df_cna_data.rename(columns={'Part Number': 'SKU No.'})
        df_cna_data = df_cna_data.rename(columns={'Branch Name': 'Store Name'})
        df_cna_data = df_cna_data.rename(columns={'Sum of Unit Sales': 'Sales Qty'})
        df_cna_data = df_cna_data.rename(columns={'Date Decarded': 'Start Date'})

        # Lookup column
        df_cna_data['Lookup'] = df_cna_data['SKU No.'].astype(str) + df_cna_data['Store Name']

        # Get stock on hand
        df_cna_soh = df_cna_soh.rename(columns={'Branch Name': 'Store Name'})
        df_cna_soh = df_cna_soh.rename(columns={'Sum of Total Stock': 'SOH Qty'})
        df_cna_soh['Lookup'] = df_cna_soh['Product Code'].astype(str) + df_cna_soh['Store Name']
        df_cna_soh_final = df_cna_soh[['Lookup','SOH Qty']]

        # Merge with SOH
        df_cna_data = df_cna_data.merge(df_cna_soh_final, how='left', on='Lookup')

        # Merge with retailer map
        df_cna_merged = df_cna_data.merge(df_cna_retailers_map, how='left', on='SKU No.')
        df_cna_merged = df_cna_merged.rename(columns={'Full Description': 'Product Description'})

        # Find missing data
        missing_model_cna = df_cna_merged['Product Code'].isnull()
        df_cna_missing_model = df_cna_merged[missing_model_cna]
        df_missing = df_cna_missing_model[['SKU No.','Product Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)
        st.write(" ")

    except:
        st.markdown("**Retailer map column headings:** NEW- Retailers Article code, SMD Code, Description ,RSP")
        st.markdown("**Retailer data column headings:** Branch Name, Part Number, Sum of Unit Sales, Sum of Sales Excl VAT, Date Decarded, Full Description")
        st.markdown("**Retailer SOH column headings:** Branch Name, Product Code, Sum of Total Stock")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")


    try:
        # Total amount column
        df_cna_merged['Total Amt'] = df_cna_merged['Sum of Sales Excl VAT'] * 1.15

        # Add retailer and store column
        df_cna_merged['Forecast Group'] = 'Edcon CNA Audio and Digital'

        # Don't change these headings. Rather change the ones above
        final_df_cna = df_cna_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_cna_p = df_cna_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_cna_s = df_cna_merged[['Store Name','Total Amt']]    

        # Show final df
        df_stats(final_df_cna,final_df_cna_p,final_df_cna_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_cna), unsafe_allow_html=True)

    except:
        st.write('Check data') 


# Cross Trainer
elif option == 'Cross_Trainer':

    try:
        # Get retailers map
        df_ct_retailers_map = df_map
        df_ct_retailers_map = df_ct_retailers_map.rename(columns={'Cross Trainer Product Code':'Item Code'})
        df_retailers_map_ct_final = df_ct_retailers_map[['Item Code','SMD Product Code', 'SMD Description','RSP']] 
        
        # Get retailer data
        df_ct_data = df_data
        
        # Merge with retailer map
        df_ct_merged = df_ct_data.merge(df_retailers_map_ct_final, how='left', on='Item Code')

        # Rename columns
        df_ct_merged = df_ct_merged.rename(columns={'Item Code': 'SKU No.'})
        df_ct_merged = df_ct_merged.rename(columns={'Qty': 'Sales Qty'})
        df_ct_merged = df_ct_merged.rename(columns={'SOH': 'SOH Qty'})
        df_ct_merged = df_ct_merged.rename(columns={'SMD Product Code': 'Product Code'})
        df_ct_merged = df_ct_merged.rename(columns={'Stores': 'Store Name'})
        df_ct_merged = df_ct_merged.rename(columns={'SMD Description': 'Product Description'})

        # Find missing data
        missing_model_ct = df_ct_merged['Product Code'].isnull()
        df_ct_missing_model = df_ct_merged[missing_model_ct]
        df_missing = df_ct_missing_model[['SKU No.','Item Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp_ct = df_ct_merged['RSP'].isnull()
        df_ct_missing_rsp = df_ct_merged[missing_model_ct] 
        df_missing_2 = df_ct_missing_rsp[['SKU No.','Item Description']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** SMD Product Code, SMD Description, Cross Trainer Product Code, RSP")
        st.markdown("**Retailer data column headings:** Stores, Item Code, Item Description, SOH, Qty")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")   

    try:
        # Set date columns
        df_ct_merged['Start Date'] = Date_Start

        # Add Total Amount column
        df_ct_merged['Total Amt'] = df_ct_merged['Sales Qty'] * df_ct_merged['RSP']

        # Add column for retailer and store name
        df_ct_merged['Forecast Group'] = 'Cross Trainer'
        df_ct_merged['Store Name'] = df_ct_merged['Store Name'].str.title()
        
        # Final df. Don't change these headings. Rather change the ones above
        final_df_ct_sales = df_ct_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_ct_p = df_ct_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_ct_s = df_ct_merged[['Store Name','Total Amt']]

        # Show final df
        df_stats(final_df_ct_sales,final_df_ct_p,final_df_ct_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_ct_sales), unsafe_allow_html=True)

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
        df_dealz_merged[units_sold] = df_dealz_merged[units_sold].fillna(0)
        df_dealz_merged['Total Amt'] = df_dealz_merged[units_sold].astype(int) * df_dealz_merged['Price'].astype(float)

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
        final_df_dealz_sales['Total Amt'] = final_df_dealz_sales['Total Amt'].astype(float)
        final_df_dealz_sales['Sales Qty'] = final_df_dealz_sales['Sales Qty'].astype(float)
        df_stats(final_df_dealz_sales,final_df_dealz_p,final_df_dealz_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_dealz_sales), unsafe_allow_html=True)

    except:
        st.write('Check data. Make sure the month in long form eg. '+units_sold)


# Decofurn
elif option == 'Decofurn':
    try:
        # Get retailers map
        df_dcf_retailers_map = df_map
        df_retailers_map_dcf_final = df_dcf_retailers_map[['Article','Product Code', 'SMD Description', 'RSP']]
        df_retailers_map_dcf_final['Article'] = df_retailers_map_dcf_final['Article'].astype(str)    

        # Get retailer data
        df_dcf_data = df_data
        df_dcf_data.columns = df_dcf_data.iloc[0]
        df_dcf_data = df_dcf_data.iloc[1:]
        
        # Merge with retailer map
        df_dcf_merged = df_dcf_data.merge(df_retailers_map_dcf_final, how='left', on='Article')
        df_dcf_merged.columns = df_dcf_merged.columns.str.title()

        # Rename columns
        df_dcf_merged = df_dcf_merged.rename(columns={'Article': 'SKU No.'})
        df_dcf_merged = df_dcf_merged.rename(columns={'Soh': 'SOH Qty'})
        df_dcf_merged = df_dcf_merged.rename(columns={'Rsp': 'RSP'})
        df_dcf_merged = df_dcf_merged.rename(columns={'Sales': 'Sales Qty'})
        df_dcf_merged = df_dcf_merged.rename(columns={'Smd Description': 'Product Description'})
        
        # Find missing data
        missing_model_dcf = df_dcf_merged['Product Code'].isnull()
        df_dcf_missing_model = df_dcf_merged[missing_model_dcf]
        df_missing = df_dcf_missing_model[['SKU No.','Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

        st.write(" ")
        missing_rsp_dcf = df_dcf_merged['RSP'].isnull()
        df_dcf_missing_rsp = df_dcf_merged[missing_rsp_dcf] 
        df_missing_2 = df_dcf_missing_rsp[['SKU No.','Description']]
        df_missing_unique_2 = df_missing_2.drop_duplicates()
        st.write("The following products are missing the RSP on the map: ")
        st.table(df_missing_unique_2)

    except:
        st.markdown("**Retailer map column headings:** Product Code, Article, SMD Description, RSP")
        st.markdown("**Retailer data column headings:** Article, Description, Store Name, SOH, Sales "+str(Short_Date_Dict[Month]))
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_dcf_merged['Start Date'] = Date_End

        # Total amount column
        df_dcf_merged['Total Amt'] = df_dcf_merged['Sales Qty'] * df_dcf_merged['RSP']

        # Add retailer and store column
        df_dcf_merged['Forecast Group'] = 'Decofurn'

        # Don't change these headings. Rather change the ones above
        final_df_dcf = df_dcf_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_dcf_p = df_dcf_merged[['Product Code','Product Description', 'Sales Qty', 'Total Amt']]
        final_df_dcf_s = df_dcf_merged[['Store Name','Total Amt']]

        # Show final df
        df_stats(final_df_dcf,final_df_dcf_p,final_df_dcf_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_dcf), unsafe_allow_html=True)
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
        df_dischem_merged = df_dischem_merged.rename(columns={Units_Sold: 'Sales Qty'})
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
        df_stats(final_df_dischem_sales,final_df_dischem_p,final_df_dischem_s)

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
        df_dischemp_merged = df_dischemp_merged.rename(columns={Units_Sold: 'Sales Qty'})
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
        df_stats(final_df_dischemp_sales,final_df_dischemp_p,final_df_dischemp_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_dischemp_sales), unsafe_allow_html=True)

    except:
        st.write('Check data') 

# Game

elif option == 'Game':

    game_soh = st.file_uploader('SOH',type=['csv','txt','xlsx'])
    if game_soh:    
        if game_soh.name[-3:] == 'csv':
            game_soh.seek(0)
            df_game_soh = pd.read_csv(io.StringIO(game_soh.read().decode('utf-8')), delimiter='|')
            df_game_soh = df_game_soh.rename(columns=lambda x: x.strip())

        elif game_soh.name[-3:] == 'txt':
            game_soh.seek(0)
            df_game_soh = pd.read_csv(io.StringIO(game_soh.read().decode('utf-8')), delimiter='|')
            df_game_soh = df_game_soh.rename(columns=lambda x: x.strip())

        else:
            df_game_soh = pd.read_excel(game_soh)
            df_game_soh = df_game_soh.rename(columns=lambda x: x.strip())
   
    try:
        # Get retailers map
        df_game_retailers_map = df_map
        df_game_retailers_map = df_game_retailers_map.rename(columns={'SMD Description': 'Product Description'})
        df_game_retailers_map = df_game_retailers_map.rename(columns={'Article number': 'Article'})
        df_retailers_map_game_final = df_game_retailers_map[['Article','SMD Code','Product Description']]

        # Get retailer data
        df_game_data = df_data
        df_game_data = df_game_data[df_game_data['StartDate'].notna()]
        df_game_data['Lookup'] = df_game_data['MaterialCode'].astype(str) + df_game_data['PlantCode']
        df_game_data = df_game_data.rename(columns={'MaterialCode': 'Article'})

        # Merge with retailer map 
        df_game_merged = df_game_data.merge(df_retailers_map_game_final, how='left', on='Article')

        # Merge with SOH
        df_game_soh['Lookup'] = df_game_soh['MaterialCode'].astype(str) + df_game_soh['PlantCode']
        df_game_soh_final = df_game_soh[['Lookup', 'StockOnHand']]
        df_game_merged = df_game_merged.merge(df_game_soh_final, how='left', on='Lookup')

        # Find missing data
        missing_model_game = df_game_merged['SMD Code'].isnull()
        df_game_missing_model = df_game_merged[missing_model_game]
        df_missing = df_game_missing_model[['Article','MaterialDescription']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

    except:
        st.markdown("**Retailer map column headings:** Article, SMD Product Code, SMD Description")
        st.markdown("**Retailer data column headings:** EndDate, ProductCode, ProductDescription, SiteDescription, Quantity, ValueExcl")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")
    
    try:
        # Set date columns
        df_game_merged['EndDate'] = pd.to_datetime(df_game_merged['EndDate'])

        # Total amount column
        df_game_merged['Total Amt'] = df_game_merged['ValueExcl'] + df_game_merged['VAT']
        
        # Add retailer column
        df_game_merged['Forecast Group'] = 'Game'

        # Rename columns and Store Name proper
        df_game_merged = df_game_merged.rename(columns={'EndDate': 'Start Date'})
        df_game_merged = df_game_merged.rename(columns={'Article': 'SKU No.'})
        df_game_merged = df_game_merged.rename(columns={'SMD Code': 'Product Code'})
        df_game_merged = df_game_merged.rename(columns={'Quantity': 'Sales Qty'})
        df_game_merged = df_game_merged.rename(columns={'PlantName': 'Store Name'})
        df_game_merged = df_game_merged.rename(columns={'StockOnHand': 'SOH Qty'})
        df_game_merged['Store Name'] = df_game_merged['Store Name'].str.title()
        
        # Don't change these headings. Rather change the ones above
        final_df_game = df_game_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_game_p = df_game_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_game_s = df_game_merged[['Store Name','Total Amt']]

        # Show final df
        df_stats(final_df_game,final_df_game_p,final_df_game_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_game), unsafe_allow_html=True)

    except:
        st.write('Check data')

# HiFi Corp (Monthly)

elif option == 'HiFi':
    try:
        Units_Sold = ('Qty Sold '+Long_Date_Dict[Month])
        Value_Sold = ('Sales Value '+Long_Date_Dict[Month])
        
        # Get retailers map
        df_hifi_retailer_map = df_map
        df_hifi_retailer_map = df_hifi_retailer_map.drop_duplicates(subset=['Article'])

        # Get current week
        df_hifi_data = df_data
        df_hifi_data = df_hifi_data.rename(columns=lambda x: x.strip())
        df_hifi_data[Value_Sold] = df_hifi_data[Value_Sold].replace(' ','', regex=True)
        df_hifi_data[Value_Sold] = df_hifi_data[Value_Sold].fillna(0)
        df_hifi_data[Value_Sold] = df_hifi_data[Value_Sold].astype(int)
        df_hifi_data['Lookup'] = df_hifi_data['Article'].astype(str) + df_hifi_data['Site']

        # Merge with retailer map and previous week
        df_hifi_merged = df_hifi_data.merge(df_hifi_retailer_map, how='left', on='Article')

        # Find missing data
        missing_model_hifi = df_hifi_merged['SMD Code'].isnull()
        df_hifi_missing_model = df_hifi_merged[missing_model_hifi]
        df_missing = df_hifi_missing_model[['Article','Article Name']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

    except:
        st.markdown("**Retailer map column headings:** Article, SMD Code, Product Description & RSP")
        st.markdown("**Retailer data column headings:** Material, Material Desc, Plant, Plant Description, Total SOH Qty & "+Units_Sold+", "+Value_Sold)
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_hifi_merged['Start Date'] = Date_End

        # Add Total Amount column
        df_hifi_merged = df_hifi_merged.rename(columns={Units_Sold : 'Sales Qty'})
        df_hifi_merged['Total Amt'] = df_hifi_merged[Value_Sold] * 1.15

        # Add column for retailer and SOH
        df_hifi_merged['Forecast Group'] = 'HIFI Corp'

        # Rename columns
        df_hifi_merged = df_hifi_merged.rename(columns={'Article': 'SKU No.'})
        df_hifi_merged = df_hifi_merged.rename(columns={'Total SOH Qty': 'SOH Qty'})
        df_hifi_merged = df_hifi_merged.rename(columns={'SMD Code': 'Product Code'})
        df_hifi_merged = df_hifi_merged.rename(columns={'Site Name': 'Store Name'})

        # Final df. Don't change these headings. Rather change the ones above
        final_df_hifi_sales = df_hifi_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_hifi_p = df_hifi_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_hifi_s = df_hifi_merged[['Store Name','Total Amt']]

        # Show final df
        df_stats(final_df_hifi_sales,final_df_hifi_p,final_df_hifi_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_hifi_sales), unsafe_allow_html=True)

    except:
        st.write('Check data')


# HiFi Corp (Accumalated)

# elif option == 'HiFi':
#     try:
#         Units_Sold = ('Qty Sold '+ str(Month) + '.' + Year)

#         # Get retailers map
#         df_hifi_retailer_map = df_map
           
#         # Get previous week
#         hifi_data_prev = st.file_uploader('Previous week', type='xlsx')
#         if hifi_data_prev:
#             df_hifi_data_prev = pd.read_excel(hifi_data_prev)
#         df_hifi_data_prev = df_hifi_data_prev.rename(columns=lambda x: x.strip())
#         df_hifi_data_prev['Lookup'] = df_hifi_data_prev['Material'].astype(str) + df_hifi_data_prev['Plant']
#         df_hifi_data_prev = df_hifi_data_prev.rename(columns={Units_Sold: 'Prev Sales'})
#         df_hifi_data_prev = df_hifi_data_prev[['Lookup','Prev Sales']]

#         # Get current week
#         df_hifi_data = df_data
#         df_hifi_data['Lookup'] = df_hifi_data['Material'].astype(str) + df_hifi_data['Plant']

#         # Merge with retailer map and previous week
#         df_hifi_data_merge_curr = df_hifi_data.merge(df_hifi_data_prev, how='outer', on='Lookup')
#         df_hifi_merged = df_hifi_data_merge_curr.merge(df_hifi_retailer_map, how='left', on='Material')
#         # df_hifi_merged = df_hifi_merged.drop_duplicates(subset=['Lookup'])

#         missing_model_hifi = df_hifi_merged['SMD Code'].isnull()
#         df_hifi_missing_model = df_hifi_merged[missing_model_hifi]
#         df_missing = df_hifi_missing_model[['Material','Material Desc']]
#         df_missing_unique = df_missing.drop_duplicates()
#         st.write("The following products are missing the SMD code on the map: ")
#         st.table(df_missing_unique)

#         st.write(" ")
#         missing_rsp_hifi = df_hifi_merged['RSP'].isnull()
#         df_hifi_missing_rsp = df_hifi_merged[missing_rsp_hifi]
#         df_missing_2 = df_hifi_missing_rsp[['Material','Material Desc']]
#         df_missing_unique_2 = df_missing_2.drop_duplicates()
#         st.write("The following products are missing the RSP on the map: ")
#         st.table(df_missing_unique_2)

#     except:
#         st.markdown("**Retailer map column headings:** Material, SMD Code, Product Description & RSP")
#         st.markdown("**Retailer data column headings:** Material, Material Desc, Plant, Plant Description, Total SOH Qty & "+Units_Sold)
#         st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

#     try:
#         # Set date columns
#         df_hifi_merged['Start Date'] = Date_Start

#         # Add Total Amount column
#         df_hifi_merged['Sales Qty'] = df_hifi_merged[Units_Sold] - df_hifi_merged['Prev Sales']
#         df_hifi_merged['Total Amt'] = df_hifi_merged['Sales Qty'] * df_hifi_merged['RSP']

#         # Add column for retailer and SOH
#         df_hifi_merged['Forecast Group'] = 'HIFI Corp'

#         # Rename columns
#         df_hifi_merged = df_hifi_merged.rename(columns={'Material': 'SKU No.'})
#         df_hifi_merged = df_hifi_merged.rename(columns={'Total SOH Qty': 'SOH Qty'})
#         df_hifi_merged = df_hifi_merged.rename(columns={'SMD Code': 'Product Code'})
#         df_hifi_merged = df_hifi_merged.rename(columns={'Plant Description': 'Store Name'})

#         # Final df. Don't change these headings. Rather change the ones above
#         final_df_hifi_sales = df_hifi_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
#         final_df_hifi_p = df_hifi_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
#         final_df_hifi_s = df_hifi_merged[['Store Name','Total Amt']]

#         # Show final df
#         df_stats(final_df_hifi_sales,final_df_hifi_p,final_df_hifi_s)

#         # Output to .xlsx
#         st.write('Please ensure that no products are missing before downloading!')
#         st.markdown(get_table_download_link(final_df_hifi_sales), unsafe_allow_html=True)

#     except:
#         st.write('Check data')

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
        df_hh_data_prev = df_hh_data_prev.rename(columns=lambda x: x.strip())
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
        df_hh_merged = df_hh_merged.drop_duplicates(subset=['Lookup'])

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
        df_stats(final_df_hh_sales,final_df_hh_p,final_df_hh_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_hh_sales), unsafe_allow_html=True)
    except:
        st.write('Check data')

# Incredible Connection
elif option == 'Incredible-Connection':

    try:
        Units_Sold = ('Qty Sold '+ Long_Date_Dict[Month])
        Value_Sold = ('Sales Value '+Long_Date_Dict[Month])
        
        # Get retailers map
        df_ic_retailers_map = df_map
        df_ic_retailers_map = df_ic_retailers_map.drop_duplicates(subset=['Article'])

        # Get current week
        df_ic_data = df_data
        df_ic_data = df_ic_data.rename(columns=lambda x: x.strip())
        df_ic_data[Value_Sold] = df_ic_data[Value_Sold].replace(' ','', regex=True)
        df_ic_data[Value_Sold] = df_ic_data[Value_Sold].fillna(0)
        df_ic_data[Value_Sold] = df_ic_data[Value_Sold].astype(int)

        # Rename columns
        df_ic_retailers_map = df_ic_retailers_map.rename(columns={'RRP': 'RSP'})

        # Merge with retailer map and previous week
        df_ic_merged = df_ic_data.merge(df_ic_retailers_map, how='left', on='Article')

        # Find missing data
        missing_model_ic = df_ic_merged['SMD Code'].isnull()
        df_ic_missing_model = df_ic_merged[missing_model_ic]
        df_missing = df_ic_missing_model[['Article','Article Name']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)


    except:
        st.markdown("**Retailer map column headings:** Article, SMD Code, Product Description & RRP")
        st.markdown("**Retailer data column headings:** Article, Article Name, Site, Site Name, Total SOH Qty & "+Units_Sold+", "+Value_Sold)
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_ic_merged['Start Date'] = Date_End

        # Add Total Amount column
        df_ic_merged = df_ic_merged.rename(columns={Units_Sold: 'Sales Qty'})
        df_ic_merged['Total Amt'] = df_ic_merged[Value_Sold]*1.15

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
        df_stats(final_df_ic_sales,final_df_ic_p,final_df_ic_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_ic_sales), unsafe_allow_html=True)

    except:
        st.write('Check data')

# # Incredible Connection (Accumulated)
# elif option == 'Incredible-Connection':
#     try:
#         Units_Sold = ('Qty Sold '+ str(Month) + '.' + Year)

#         # Get retailers map
#         df_ic_retailers_map = df_map
  
#         # Get previous week
#         ic_data_prev = st.file_uploader('Previous week', type='xlsx')
#         if ic_data_prev:
#             df_ic_data_prev = pd.read_excel(ic_data_prev)
#         df_ic_data_prev = df_ic_data_prev.rename(columns=lambda x: x.strip())
#         df_ic_data_prev['Lookup'] = df_ic_data_prev['Article'].astype(str) + df_ic_data_prev['Site']
#         df_ic_data_prev = df_ic_data_prev.rename(columns={Units_Sold: 'Prev Sales'})
#         df_ic_data_prev_final = df_ic_data_prev[['Lookup','Prev Sales']]

#         # Get current week
#         df_ic_data = df_data
#         df_ic_data = df_ic_data.rename(columns=lambda x: x.strip())
#         df_ic_data['Lookup'] = df_ic_data['Article'].astype(str) + df_ic_data['Site']

#         # Rename columns
#         df_ic_retailers_map = df_ic_retailers_map.rename(columns={'RRP': 'RSP'})

#         # Merge with retailer map and previous week
#         df_ic_data_merge_curr = df_ic_data.merge(df_ic_data_prev_final, how='left', on='Lookup')
#         df_ic_merged = df_ic_data_merge_curr.merge(df_ic_retailers_map, how='left', on='Article')
#         df_ic_merged = df_ic_merged.drop_duplicates(subset=['Lookup'])

#         missing_model_ic = df_ic_merged['SMD Code'].isnull()
#         df_ic_missing_model = df_ic_merged[missing_model_ic]
#         df_missing = df_ic_missing_model[['Article','Article Name']]
#         df_missing_unique = df_missing.drop_duplicates()
#         st.write("The following products are missing the SMD code on the map: ")
#         st.table(df_missing_unique)

#         # Find missing data
#         st.write(" ")
#         missing_rsp_ic = df_ic_merged['RSP'].isnull()
#         df_ic_missing_rsp = df_ic_merged[missing_rsp_ic]
#         df_missing_2 = df_ic_missing_rsp[['Article','Article Name']]
#         df_missing_unique_2 = df_missing_2.drop_duplicates()
#         st.write("The following products are missing the RSP on the map: ")
#         st.table(df_missing_unique_2)

#     except:
#         st.markdown("**Retailer map column headings:** Article, SMD Code, Product Description & RRP")
#         st.markdown("**Retailer data column headings:** Article, Article Name, Site, Site Name, Total SOH Qty & "+Units_Sold)
#         st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

#     try:
#         # Set date columns
#         df_ic_merged['Start Date'] = Date_Start

#         # Add Total Amount column
#         df_ic_merged['Sales Qty'] = df_ic_merged[Units_Sold] - df_ic_merged['Prev Sales']
#         df_ic_merged['Total Amt'] = df_ic_merged['Sales Qty'] * df_ic_merged['RSP']

#         # Add column for retailer and SOH
#         df_ic_merged['Forecast Group'] = 'Incredible Connection'

#         # Rename columns
#         df_ic_merged = df_ic_merged.rename(columns={'Article': 'SKU No.'})
#         df_ic_merged = df_ic_merged.rename(columns={'Total SOH Qty': 'SOH Qty'})
#         df_ic_merged = df_ic_merged.rename(columns={'SMD Code': 'Product Code'})
#         df_ic_merged = df_ic_merged.rename(columns={'Site Name': 'Store Name'})

#         # Final df. Don't change these headings. Rather change the ones above
#         final_df_ic_sales = df_ic_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
#         final_df_ic_p = df_ic_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
#         final_df_ic_s = df_ic_merged[['Store Name','Total Amt']]

#         # Show final df
#         df_stats(final_df_ic_sales,final_df_ic_p,final_df_ic_s)

#         # Output to .xlsx
#         st.write('Please ensure that no products are missing before downloading!')
#         st.markdown(get_table_download_link(final_df_ic_sales), unsafe_allow_html=True)

#     except:
#         st.write('Check data')

# J.A.M.

elif option == 'J.A.M.':

    try:
        # Get retailers map
        df_jam_retailers_map = df_map
        df_jam_retailers_map = df_jam_retailers_map.rename(columns={'Description': 'Product Description'})

        # Get retailer data
        df_jam_data = df_data
        df_jam_data.columns = df_jam_data.iloc[6]
        df_jam_data = df_jam_data.iloc[7:]
        df_jam_data = df_jam_data.dropna(subset=['Description'])
        df_jam_data = df_jam_data.rename(columns={'Product': 'Item Number'})
        
        # Merge with retailer map
        df_jam_merged = df_jam_data.merge(df_jam_retailers_map, how='left', on='Item Number')
                
        # Find missing data
        missing_model_jam = df_jam_merged['Product Code'].isnull()
        df_jam_missing_model = df_jam_merged[missing_model_jam]
        df_missing = df_jam_missing_model[['Item Number','Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

    except:
        st.markdown("**Retailer map column headings:** Item Number, Product Code & Description")
        st.markdown("**Retailer data column headings:** Product, Description, SOO, SOH, SIT, Price (Incl) & Qty Sold")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_jam_merged['Start Date'] = Date_Start

        # Stores
        df_jam_merged['Store Name'] = ''

        # Add Total Amount column
        df_jam_merged['Total Amt'] = df_jam_merged['Price (Incl)'] * df_jam_merged['Qty Sold']

        # Add column for retailer and SOH
        df_jam_merged['Forecast Group'] = 'J.A.M Clothing'
        df_jam_merged['SOH Qty'] = df_jam_merged['SOO'] + df_jam_merged['SOH'] + df_jam_merged['SIT']

        # Rename columns
        df_jam_merged = df_jam_merged.rename(columns={'Item Number': 'SKU No.'})
        df_jam_merged = df_jam_merged.rename(columns={'Qty Sold': 'Sales Qty'})

        # Final df. Don't change these headings. Rather change the ones above
        final_df_jam_sales = df_jam_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_jam_p = df_jam_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_jam_s = df_jam_merged[['Store Name','Total Amt']]

        # Show final df
        df_stats(final_df_jam_sales,final_df_jam_p,final_df_jam_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_jam_sales), unsafe_allow_html=True)

    except:
        st.write('Check data')

# Makro

elif option == 'Makro':

    week = (dt.date(int(Year),int(Month),int(Day)).isocalendar()[1]) + 1
    if week < 10:
        week_sales = ('0'+str(week)+'-'+str(Year))
    else:
        week_sales = (str(week)+'-'+str(Year))

    makro_soh = st.file_uploader('SOH',type=['csv','txt','xlsx'])
    if makro_soh:    
        if makro_soh.name[-3:] == 'csv':
            makro_soh.seek(0)
            df_makro_soh = pd.read_csv(io.StringIO(makro_soh.read().decode('utf-8')), delimiter='|')
            df_makro_soh = df_makro_soh.rename(columns=lambda x: x.strip())

        elif makro_soh.name[-3:] == 'txt':
            makro_soh.seek(0)
            df_makro_soh = pd.read_csv(io.StringIO(makro_soh.read().decode('utf-8')), delimiter='|')
            df_makro_soh = df_makro_soh.rename(columns=lambda x: x.strip())

        else:
            df_makro_soh = pd.read_excel(makro_soh)
            df_makro_soh = df_makro_soh.rename(columns=lambda x: x.strip())
    
    try:
        # Get retailers map
        df_makro_retailers_map = df_map
        df_makro_retailers_map = df_makro_retailers_map.rename(columns={'SMD Description': 'Product Description'})
        df_retailers_map_makro_final = df_makro_retailers_map[['Article','SMD Product Code','Product Description']]

        # Get retailer data
        df_makro_data = df_data
        df_makro_data = df_makro_data[df_makro_data['StartDate'].notna()]
        df_makro_data['Lookup'] = df_makro_data['ProductCode'].astype(str) + df_makro_data['SiteCode']
        df_makro_data = df_makro_data.rename(columns={'ProductCode': 'Article'})

        # Merge with retailer map 
        df_makro_merged = df_makro_data.merge(df_retailers_map_makro_final, how='left', on='Article')

        # Merge with SOH
        df_makro_soh['Lookup'] = df_makro_soh['ProductCode'].astype(str) + df_makro_soh['SiteCode']
        df_makro_soh_final = df_makro_soh[['Lookup', 'StockOnHand']]
        df_makro_merged = df_makro_merged.merge(df_makro_soh_final, how='left', on='Lookup')

        # Find missing data
        missing_model_makro = df_makro_merged['SMD Product Code'].isnull()
        df_makro_missing_model = df_makro_merged[missing_model_makro]
        df_missing = df_makro_missing_model[['Article','ProductDescription']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

    except:
        st.markdown("**Retailer map column headings:** Article, SMD Product Code, SMD Description")
        st.markdown("**Retailer data column headings:** EndDate, ProductCode, ProductDescription, SiteDescription, Quantity, ValueExcl")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")
    
    try:

        # Get Date
        df_makro_merged['EndDate'] = pd.to_datetime(df_makro_merged['EndDate'])

        # Total amount column
        df_makro_merged['Total Amt'] = df_makro_merged['ValueExcl'] + df_makro_merged['VAT']
        
        # Add retailer column
        df_makro_merged['Forecast Group'] = 'Makro'

        # Rename columns
        df_makro_merged = df_makro_merged.rename(columns={'EndDate': 'Start Date'})
        df_makro_merged = df_makro_merged.rename(columns={'Article': 'SKU No.'})
        df_makro_merged = df_makro_merged.rename(columns={'SMD Product Code': 'Product Code'})
        df_makro_merged = df_makro_merged.rename(columns={'StockOnHand': 'SOH Qty'})
        df_makro_merged = df_makro_merged.rename(columns={'Quantity': 'Sales Qty'})
        df_makro_merged = df_makro_merged.rename(columns={'SiteDescription': 'Store Name'})
        
        # Don't change these headings. Rather change the ones above
        final_df_makro = df_makro_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_makro_p = df_makro_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_makro_s = df_makro_merged[['Store Name','Total Amt']]

        # Show final df
        df_stats(final_df_makro,final_df_makro_p,final_df_makro_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_makro), unsafe_allow_html=True)

    except:
        st.write('Check data')

# Mr Price Sport
elif option == 'Mr-Price-Sport':
    try:
        # Get retailers map
        df_mrp_retailers_map = df_map
        df_mrp_retailers_map = df_mrp_retailers_map.rename(columns={'RRP':'RSP'})
        df_mrp_retailers_map = df_mrp_retailers_map.rename(columns={'Retailer Item No.':'Item Number'})
        df_retailers_map_mrp_final = df_mrp_retailers_map[['Item Number','SMD Code', 'Product Description', 'RSP']]

        # Get retailer data
        df_mrp_data = df_data
        df_mrp_data = df_mrp_data.iloc[2:] 

        # Merge with retailer map
        df_mrp_merged = df_mrp_data.merge(df_retailers_map_mrp_final, how='left', on='Item Number') 
        
        # Rename columns
        df_mrp_merged = df_mrp_merged.rename(columns={'Item Number': 'SKU No.'})
        df_mrp_merged = df_mrp_merged.rename(columns={'T/Y Sales.1': 'Total Amt'})
        df_mrp_merged = df_mrp_merged.rename(columns={'T/Y Sales': 'Sales Qty'})
        df_mrp_merged = df_mrp_merged.rename(columns={'T/Y Close SOH': 'SOH Qty'})
        df_mrp_merged = df_mrp_merged.rename(columns={'SMD Code': 'Product Code'})

        # Find missing data
        missing_model = df_mrp_merged['Product Code'].isnull()
        df_mrp_missing_model = df_mrp_merged[missing_model]
        df_missing = df_mrp_missing_model[['SKU No.','Item Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

    except:
        st.markdown("**Retailer map column headings:** Retailer Item No., SMD Code, Product Description")
        st.markdown("**Retailer data column headings:** Branch Description, Item Number, Item Description, T/Y Sales, T/Y Close SOH")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_mrp_merged['Start Date'] = Date_Start

        # Add retailer column and Store Name
        df_mrp_merged['Forecast Group'] = 'Mr Price Sport'
        df_mrp_merged['Store Name'] = df_mrp_merged['Branch Description'].str.title()

        # Don't change these headings. Rather change the ones above
        final_df_mrp = df_mrp_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_mrp_p = df_mrp_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_mrp_s = df_mrp_merged[['Store Name','Total Amt']]        

        # Show final df
        df_stats(final_df_mrp,final_df_mrp_p,final_df_mrp_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_mrp), unsafe_allow_html=True)

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
        df_musica_data.columns = df_musica_data.iloc[0]
        df_musica_data = df_musica_data.iloc[1:]
        df_musica_data = df_musica_data.rename(columns={'SKU No.': 'Musica Code'})
        df_musica_data = df_musica_data.rename(columns={'4 Wks sales Qty': 'Sales Qty'})  

        # Merge with retailer map
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
        st.markdown("**Retailer data column headings:** Store Name, SKU No., Title Desc, Selling_Price, 4 Wks sales Qty, SOH Qty")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_musica_merged['Start Date'] = Date_Start

        # Total amount column
        df_musica_merged['Total Amt'] = df_musica_merged['Sales Qty'] * df_musica_merged['Selling_Price']

        # Add retailer column
        df_musica_merged['Forecast Group'] = 'Musica'
        df_musica_merged['Store Name'] = ''

        # Rename columns
        df_musica_merged = df_musica_merged.rename(columns={'Musica Code': 'SKU No.'})
        df_musica_merged = df_musica_merged.rename(columns={'SMD code': 'Product Code'})

        # Don't change these headings. Rather change the ones above
        final_df_musica = df_musica_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_musica_p = df_musica_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_musica_s = df_musica_merged[['Store Name','Total Amt']]        

        # Show final df
        df_stats(final_df_musica,final_df_musica_p,final_df_musica_s)

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
        df_okf_data_prev = df_okf_data_prev.rename(columns=lambda x: x.strip())
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
        df_okf_merged['Qty Sold'].fillna(0,inplace=True)
        df_okf_merged['Prev Qty'].fillna(0,inplace=True)

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
        final_df_okf_p = df_okf_merged[['Product Code', 'Product Description','Sales Qty','Total Amt']]
        final_df_okf_s = df_okf_merged[['Store Name','Total Amt']]   

        # Show final df
        df_stats(final_df_ok_sales,final_df_okf_p,final_df_okf_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_ok_sales), unsafe_allow_html=True)
    except:
        st.write('Check data')

# Ok Furniture Africa

elif option == 'Ok-Furniture-Africa':

    try:
        # Get retailers map
        df_okfa_retailers_map = df_map
        df_okfa_retailers_map_final = df_okfa_retailers_map[['SKU Number','SMD Product Code','SMD Description']]

        # Get previous week
        okfa_data_prev = st.file_uploader('Previous week', type='xlsx')
        if okfa_data_prev:
            df_okfa_data_prev = pd.read_excel(okfa_data_prev)
        df_okfa_data_prev = df_okfa_data_prev.rename(columns=lambda x: x.strip())
        df_okfa_data_prev['Lookup'] = df_okfa_data_prev['SKU Number'].astype(str) + df_okfa_data_prev['Brn No'].astype(str)
        df_okfa_data_prev = df_okfa_data_prev.rename(columns={'Qty Sold': 'Prev Qty'})
        df_okfa_data_prev = df_okfa_data_prev.rename(columns={'Sold RSP': 'Prev Amt'})
        df_okfa_data_prev_final = df_okfa_data_prev[['Lookup','Prev Qty','Prev Amt']]

        # Get current week
        df_okfa_data = df_data
        df_okfa_data['Lookup'] = df_okfa_data['SKU Number'].astype(str) + df_okfa_data['Brn No'].astype(str)

        # Merge with retailer map and previous week
        df_okfa_data_merge_curr = df_okfa_data.merge(df_okfa_data_prev_final, how='left', on='Lookup')
        df_okfa_merged = df_okfa_data_merge_curr.merge(df_okfa_retailers_map_final, how='left', on='SKU Number')
        df_okfa_merged['Qty Sold'].fillna(0,inplace=True)
        df_okfa_merged['Prev Qty'].fillna(0,inplace=True)

        # Find missing data
        missing_model_okfa = df_okfa_merged['SMD Product Code'].isnull()
        df_okfa_missing_model = df_okfa_merged[missing_model_okfa]
        df_missing = df_okfa_missing_model[['SKU Number','SKU Description']]
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing)

    except:
        st.markdown("**Retailer map column headings:** SKU Number, SMD Product Code & SMD Description")
        st.markdown("**Retailer data column headings:** Brn No, Brn Description, SKU Number, SKU Description, Qty Sold, Sold RSP, Qty On Hand")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_okfa_merged['Start Date'] = Date_Start

        # Add Total Amount column
        df_okfa_merged['Sales Qty'] = df_okfa_merged['Qty Sold'] - df_okfa_merged['Prev Qty']
        df_okfa_merged['Total Amt'] = (df_okfa_merged['Sold RSP'] - df_okfa_merged['Prev Amt'])*1.15

        # Add column for retailer and SOH
        df_okfa_merged['Forecast Group'] = 'OK Furniture Africa'
        df_okfa_merged['Store Name'] = df_okfa_merged['Brn Description'].str.title()

        # Rename columns
        df_okfa_merged = df_okfa_merged.rename(columns={'SKU Number': 'SKU No.'})
        df_okfa_merged = df_okfa_merged.rename(columns={'Qty On Hand': 'SOH Qty'})
        df_okfa_merged = df_okfa_merged.rename(columns={'SMD Product Code': 'Product Code'})
        df_okfa_merged = df_okfa_merged.rename(columns={'SMD Description': 'Product Description'})

        # Final df. Don't change these headings. Rather change the ones above
        final_df_okfa_sales = df_okfa_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_okfa_p = df_okfa_merged[['Product Code', 'Product Description','Sales Qty','Total Amt']]
        final_df_okfa_s = df_okfa_merged[['Store Name','Total Amt']]   

        # Show final df
        df_stats(final_df_okfa_sales,final_df_okfa_p,final_df_okfa_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_okfa_sales), unsafe_allow_html=True)
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
        df_stats(final_df_ow,final_df_ow_p,final_df_ow_s)

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
        df_pepaf_data = df_pepaf_data.rename(columns={'Total': 'SOH Qty'})
        
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
        df_pepaf_merged = df_pepaf_merged.rename(columns={Wk_sales: 'Sales Qty'})
        df_pepaf_merged['Total Amt'] = df_pepaf_merged['Sales Qty'] * df_pepaf_merged['RSP']

        # Add retailer column
        df_pepaf_merged['Forecast Group'] = 'Pep Africa'

        # Don't change these headings. Rather change the ones above
        final_df_pepaf = df_pepaf_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_pepaf_p = df_pepaf_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_pepaf_s = df_pepaf_merged[['Store Name','Total Amt']]   

        # Show final df
        df_stats(final_df_pepaf,final_df_pepaf_p,final_df_pepaf_s)

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
        df_pep_merged['Total Amt'] = df_pep_merged['Sales Qty'].astype(float) * df_pep_merged['RSP']
        df_pep_merged['Total Amt'] = df_pep_merged['Total Amt'].apply(lambda x: round(x,2))

        # Add retailer and store column
        df_pep_merged['Forecast Group'] = 'Pep South Africa'
        df_pep_merged['Store Name'] = ''

        # Don't change these headings. Rather change the ones above
        final_df_pep = df_pep_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_pep_p = df_pep_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_pep_s = df_pep_merged[['Store Name','Total Amt']]   
        
        # Show final df
        df_stats(final_df_pep,final_df_pep_p,final_df_pep_s)

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
        df_pnp_retailers_map = df_pnp_retailers_map.rename(columns={'Article Number': 'SKU No.'})
        df_pnp_retailers_map = df_pnp_retailers_map.drop_duplicates(subset='SKU No.')
        df_retailers_map_pnp_final = df_pnp_retailers_map[['SKU No.','SMD code','Product Description']]
        
        # Get retailer data
        df_pnp_data = df_data
        df_pnp_data = df_pnp_data.rename(columns={'Day': 'Start Date'})
        df_pnp_data = df_pnp_data.rename(columns={'PnP ArticleNumber': 'SKU No.'})
        df_pnp_data = df_pnp_data.rename(columns={'Store': 'Store Name'})
        df_pnp_data = df_pnp_data.rename(columns={'Units': 'Sales Qty'})
        df_pnp_data = df_pnp_data.rename(columns={'Amount': 'Total Amt'})
        df_pnp_data = df_pnp_data.rename(columns={'Product Description': 'Article description'})
        df_pnp_data['SOH Qty'] = 0
        df_pnp_data_final = df_pnp_data[['Start Date','SKU No.','Article description','Store Name','SOH Qty','Sales Qty','Total Amt']]
        
        # Get stock on hand
        df_pnp_soh = df_pnp_soh.rename(columns={'Week Ending Date': 'Start Date'})
        df_pnp_soh = df_pnp_soh.rename(columns={'Article Number': 'SKU No.'})
        df_pnp_soh = df_pnp_soh.rename(columns={'Site Description': 'Store Name'})
        df_pnp_soh['Start Date'] =  pd.to_datetime(df_pnp_soh['Start Date'])
        df_pnp_soh['Sales Qty'] = 0
        df_pnp_soh['Total Amt'] = 0
        df_pnp_soh_final = df_pnp_soh[['Start Date','SKU No.','Article description','Store Name','SOH Qty','Sales Qty','Total Amt']]
        
        # Concatenate SOH and Sales
        df_pnp_data_concat = pd.concat([df_pnp_data_final, df_pnp_soh_final])
        df_pnp_data_concat['Store Name'] = df_pnp_data_concat['Store Name'].str.title()

        # Merge with retailer map
        df_pnp_merged = df_pnp_data_concat.merge(df_retailers_map_pnp_final, how='left', on='SKU No.')

        # Rename columns
        df_pnp_merged = df_pnp_merged.rename(columns={'SMD code': 'Product Code'})

        # Find missing data
        missing_model = df_pnp_merged['Product Code'].isnull()
        df_pnp_missing_model = df_pnp_merged[missing_model]
        df_missing = df_pnp_missing_model[['SKU No.','Article description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

    except:
        st.markdown("**Retailer map column headings:** Article Number, SMD code, Product Description, RSP")
        st.markdown("**Retailer data column headings:** Product Description, Store ID, Store, Units, PnP ArticleNumber")
        st.markdown("**Retailer SOH column headings:** Site Code, Article Number, SOH Qty")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Add retailer and store column
        df_pnp_merged['Forecast Group'] = 'Pick n Pay'

        # Don't change these headings. Rather change the ones above
        final_df_pnp = df_pnp_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_pnp_p = df_pnp_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_pnp_s = df_pnp_merged[['Store Name','Total Amt']]  

        # Show final df
        df_stats(final_df_pnp,final_df_pnp_p,final_df_pnp_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_pnp), unsafe_allow_html=True)
    except:
        st.write('Check data') 

# Retailability

elif option == 'Retailability':

    week = dt.date(int(Year),int(Month),int(Day)).isocalendar()[1]
    if week < 10:
        week_sales = ('Week 0'+str(week))
    else:
        week_sales = ('Week '+str(week))
    
    
    try:
        # Get retailers map
        df_ret_retailers_map = df_map       
        df_ret_retailers_map = df_ret_retailers_map.rename(columns={'Article Code': 'Item Colour'})
        df_ret_retailers_map_final = df_ret_retailers_map[['Item Colour','Code', 'Product Description', 'RSP']]
        
        # Get retailer data
        df_ret_data = df_data
            
        # Merge with retailer map
        df_ret_merged = df_ret_data.merge(df_ret_retailers_map_final, how='left', on='Item Colour')
        df_ret_merged = df_ret_merged.rename(columns={'Item Colour':'SKU No.'})
        df_ret_merged = df_ret_merged.rename(columns={'Code':'Product Code'})
        df_ret_merged = df_ret_merged.rename(columns={week_sales:'Sales Qty'})
        

        # Find missing data
        missing_model = df_ret_merged['Product Code'].isnull()
        df_ret_missing_model = df_ret_merged[missing_model]
        df_missing = df_ret_missing_model[['SKU No.','Item Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

    except:
        st.markdown("**Retailer map column headings:** Article Code, Code, Product Description, RSP")
        st.markdown("**Retailer data column headings:** Item Colour, Item Description, SOH Qty, Current Price (Stock)"+week_sales)
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_ret_merged['Start Date'] = Date_Start

        # Total amount column
        df_ret_merged['Total Amt'] = df_ret_merged['Sales Qty'] * df_ret_merged['Current Price (Stock)']

        # Add retailer and store column
        df_ret_merged['Forecast Group'] = 'Retailability'
        df_ret_merged['Store Name'] = ''

        # Don't change these headings. Rather change the ones above
        final_df_ret = df_ret_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_ret_p = df_ret_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_ret_s = df_ret_merged[['Store Name','Total Amt']]  

        # Show final df
        df_stats(final_df_ret,final_df_ret_p,final_df_ret_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_ret), unsafe_allow_html=True)
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
        df_sw_soh = df_sw_soh.rename(columns=lambda x: x.strip())
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
        df_stats(final_df_sw,final_df_sw_p,final_df_sw_s)

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
        df_retailers_map_takealot_final = df_takealot_retailers_map[['idProduct','Product Description','SMD Code']]

        # Get retailer data
        df_takealot_data = df_data
        df_takealot_data = df_takealot_data.iloc[1:]

        #Merge with retailer map
        df_takealot_merged = df_takealot_data.merge(df_retailers_map_takealot_final, how='left', on='idProduct')   

        # Find missing data
        missing_model = df_takealot_merged['SMD Code'].isnull()
        df_takealot_missing_model = df_takealot_merged[missing_model]
        df_missing = df_takealot_missing_model[['idProduct','ProdTitle']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)


    except:
        st.markdown("**Retailer map column headings:** idProduct, Description, SMD Code")
        st.markdown("**Retailer data column headings:** idProduct, Total_Stock, Qty, SaleValueEx")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_takealot_merged['Start Date'] = Date_Start

        # Total amount column
        df_takealot_merged['Total Amt'] = df_takealot_merged['SaleValueEx'] * 1.15

        # Add retailer and store column
        df_takealot_merged['Forecast Group'] = 'Takealot'
        df_takealot_merged['Store Name'] = ''

        # Rename columns
        df_takealot_merged = df_takealot_merged.rename(columns={'idProduct': 'SKU No.','Qty' :'Sales Qty','Total_Stock':'SOH Qty','SMD Code':'Product Code' })

        # Don't change these headings. Rather change the ones above
        final_df_takealot = df_takealot_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_takealot_p = df_takealot_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_takealot_s = df_takealot_merged[['Store Name','Total Amt']]  

        # Show final df
        df_stats(final_df_takealot,final_df_takealot_p,final_df_takealot_s)

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
        
        # Total Amount column
        df_tfg_merged['Total Amt'] = df_tfg_merged['Sales Qty'] * df_tfg_merged['RSP']

        # Add retailer and store column
        df_tfg_merged['Forecast Group'] = 'TFG'
        df_tfg_merged['Store Name'] = ''

        # Don't change these headings. Rather change the ones above
        final_df_tfg = df_tfg_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_tfg_p = df_tfg_merged[['Product Code','Product Description', 'Sales Qty', 'Total Amt']]
        final_df_tfg_s = df_tfg_merged[['Store Name','Total Amt']]

        # Show final df
        df_stats(final_df_tfg,final_df_tfg_p,final_df_tfg_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_tfg), unsafe_allow_html=True)
    except:
        st.write('Check data')

# TFG Cosmetics
elif option == 'TFG_Cosmetics':
    try:
        # Get retailers map
        df_tfgc_retailers_map = df_map
        df_retailers_map_tfgc_final = df_tfgc_retailers_map[['Supplier Style No','SMD Product Code','Product Description']]
        
        # Get retailer data
        df_tfgc_data = df_data
        df_tfgc_data.drop(df_tfgc_data[df_tfgc_data['Branch'] == 'Total'].index, inplace = True)

        # Merge with retailer map 
        df_tfgc_merged = df_tfgc_data.merge(df_retailers_map_tfgc_final, how='left', on='Supplier Style No')
        
        # Find missing data
        missing_model_tfgc = df_tfgc_merged['SMD Product Code'].isnull()
        df_tfgc_missing_model = df_tfgc_merged[missing_model_tfgc]
        df_missing = df_tfgc_missing_model[['Supplier Style No','Supplier Style Desc']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

    except:
        st.markdown("**Retailer map column headings:** Supplier Style No, SMD Product Code, Product Description, RSP")
        st.markdown("**Retailer data column headings:** Supplier Style No, Supplier Style Desc, Sls (R), Sls (U), CSOH Incl IT (U)")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_tfgc_merged['Start Date'] = Date_Start

        # Rename columns
        df_tfgc_merged = df_tfgc_merged.rename(columns={'Sls (R)': 'Total Amt','Supplier Style No': 'SKU No.','Sls (U)' :'Sales Qty', 'CSOH Incl IT (U)':'SOH Qty', 'SMD Product Code' : 'Product Code' })

        # Add retailer and store column
        df_tfgc_merged['Forecast Group'] = 'TFG - Cosmetics'
        df_tfgc_merged['Store Name'] = df_tfgc_merged['Branch'].str.title() 

        # Don't change these headings. Rather change the ones above
        final_df_tfgc = df_tfgc_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_tfgc_p = df_tfgc_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_tfgc_s = df_tfgc_merged[['Store Name','Total Amt']]

        # Show final df
        df_stats(final_df_tfgc,final_df_tfgc_p,final_df_tfgc_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_tfgc), unsafe_allow_html=True)
    except:
        st.write('Check data')

# 'Toys R Us Audio & Gaming'

elif option == 'TRU':

    units_sold = str(Long_Date_Dict[Date_End.month]) + " " + str(Date_End.year)
    
    try:
        # Get retailers map
        df_tru_retailers_map = df_map
        df_tru_retailers_map = df_tru_retailers_map[['Product Code', 'Product Description', 'SMD Code']]
        
        # Get retailer data
        df_tru_data = df_data
        df_tru_data = df_tru_data.dropna(subset=['Description'])
        df_tru_data['Product Code'] = df_tru_data['Product Code'].astype(int)
        df_tru_data = df_tru_data[~df_tru_data['Description'].str.contains('REDUCED')]
            
        # Merge with retailer map 
        df_tru_merged = df_tru_data.merge(df_tru_retailers_map, how='left', on='Product Code')
        
        # Find missing data
        missing_model_tru = df_tru_merged['SMD Code'].isnull()
        df_tru_missing_model = df_tru_merged[missing_model_tru]
        df_missing = df_tru_missing_model[['Product Code','Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)

    except:
        st.markdown("**Retailer map column headings:** Product Code, SMD Code")
        st.markdown("**Retailer data column headings:** Product Code, Description, Store Name, SOH, "+units_sold)
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct")

    try:
        # Set date columns
        df_tru_merged['Start Date'] = Date_End

        # Rename columns
        df_tru_merged = df_tru_merged.rename(columns={'Product Code': 'SKU No.',units_sold :'Sales Qty', 'SOH':'SOH Qty', 'SMD Code' : 'Product Code' })

        # Total Amount
        df_tru_merged['Total Amt'] = df_tru_merged['Sales Qty'] * df_tru_merged['RSP (incl)']

        # Add retailer and store column
        df_tru_merged['Forecast Group'] = 'Toys R Us Audio & Gaming'
        df_tru_merged['Store Name'] = df_tru_merged['Store Name'].str.title() 

        # Don't change these headings. Rather change the ones above
        final_df_tru = df_tru_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_tru_p = df_tru_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_tru_s = df_tru_merged[['Store Name','Total Amt']]

        # Show final df
        df_stats(final_df_tru,final_df_tru_p,final_df_tru_s)

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_tru), unsafe_allow_html=True)
    except:
        st.write('Check data')

else:
    st.write('Retailer not selected yet')
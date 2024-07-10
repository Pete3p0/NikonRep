# Import libraries
import streamlit as st
import pandas as pd
from io import BytesIO
import base64

def to_excel(df_bino, df_else):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_bino.to_excel(writer, sheet_name='Bino', index=False)
        df_else.to_excel(writer, sheet_name='Everything Else', index=False)
    processed_data = output.getvalue()
    return processed_data

def get_table_download_link(df_bino, df_else, date_end, week_num_use_str, filename="transformed_data.xlsx"):
    """Generates a link allowing the data in a given panda dataframe to be downloaded
    in:  dataframe
    out: href string
    """
    val = to_excel(df_bino, df_else)
    b64 = base64.b64encode(val).decode()  # Some strings <-> bytes conversions necessary here
    formatted_date = date_end.strftime('%Y-%m-%d')
    formatted_filename = f"{formatted_date}_{week_num_use_str.replace(' ', '-')}_{filename}"
    return f'<a href="data:application/octet-stream;base64,{b64}" download="{formatted_filename}">Download Excel file</a>'

def df_stats(df, df_p, df_s):
    total_units = df['Sell Out'].sum()
    st.write('**Number of units sold:** ' "{:0,.0f}".format(total_units).replace(',', ' '))
    st.write('')
    st.write('**Top 10 products sold:**')
    grouped_df_pt = df_p.groupby(["Product Description"]).agg({"Sell Out": "sum"}).sort_values("Sell Out", ascending=False)
    grouped_df_final_pt = grouped_df_pt[['Sell Out']].head(10)
    st.table(grouped_df_final_pt.style.format({'Sell Out': '{:,.0f}'}))
    st.write('')
    st.write('**Top 10 stores:**')
    grouped_df_st = df_s.groupby("Retailer").agg({"Sell Out": "sum"}).sort_values("Sell Out", ascending=False)
    grouped_df_final_st = grouped_df_st[['Sell Out']].head(10)
    st.table(grouped_df_final_st.style.format({'Sell Out': '{:,.0f}'}))
    st.write('')
    st.write('**Final Dataframe:**')
    st.dataframe(df)

st.title('Rep Sell Out & Stock on Hand')

option = st.selectbox("Select the type of report:", ["Weekly Report", "Monthly Report"])

if option == "Weekly Report":
    Date_End = st.date_input("Week ending: ")
    WeekNumUse = st.number_input("Week to look at: ", min_value=0, max_value=9, step=1, format="%d")
    WeekNumUseStr = 'Week ' + str(int(WeekNumUse))
    st.write(f"The week we are looking at is: {WeekNumUseStr}")

    WeekNumCall = st.number_input("Week to call it: ", min_value=0, max_value=9, step=1, format="%d")
    WeekNumCallStr = 'Week ' + str(int(WeekNumCall))
    st.write(f"The week we are calling it is: {WeekNumCallStr}")

    st.write("")
    st.markdown("This will handle the following reps: **Bernie, Lee, Ryan**")
    st.markdown("Please make sure the sheets in your file are named correctly")

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file is not None:
        def transform_data(df):
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
            filterdf.loc[:, 'Sell Out'] = filterdf_Sales['Value']

            filterdf = filterdf[~filterdf['Variable'].str.contains('Notes', na=False)]

            # Rename columns
            df = filterdf.rename(columns={
                'Unnamed: 0 | 365 Code': '365 Code',
                'Unnamed: 1 | Product Description': 'Product Description',
                'Unnamed: 2 | Category': 'Category',
                'Unnamed: 3 | Date SOH was Collected: | Sub-Cat': 'Sub-Cat',
                'Value': 'Stock on Hand'
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

            # Strip spaces from 'Week' column
            df['Week No.'] = df['Week No.'].str.strip()

            # Remove dots and subsequent numbers, and then strip spaces from 'Retailer' column
            df["Retailer"] = df["Retailer"].str.replace(r"\.*\d+", "", regex=True)

            # Convert 'Date SOH was Collected' column to date type
            df['Date SOH was Collected'] = pd.to_datetime(df['Date SOH was Collected']).dt.date

            return df

        # Read all sheets from the uploaded Excel file
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None)

        transformed_dfs = []
        for sheet_name, df in all_sheets.items():
            transformed_df = transform_data(df)
            transformed_df['Rep'] = sheet_name  # Add the sheet name as the 'Rep' column
            transformed_dfs.append(transformed_df)

        # Concatenate all transformed DataFrames
        final_df = pd.concat(transformed_dfs, ignore_index=True)
        
        # Filter data to include only the selected week number and call it the new week number
        final_df = final_df[final_df['Week No.'] == WeekNumUseStr]
        final_df['Week No.'] = WeekNumCallStr

        # Change the date to week ending
        final_df['Week Ending'] = Date_End

        # Don't change these headings. Rather change the ones above
        final_df = final_df[['365 Code', 'Product Description', 'Category', 'Sub-Cat', 'Rep', 'Week Ending', 'Retailer', 'Week No.', 'Stock on Hand', 'Sell Out']]
        final_df_p = final_df[['365 Code', 'Product Description', 'Sell Out']]
        final_df_s = final_df[['Retailer', 'Sell Out']]

        # Show final df
        df_stats(final_df, final_df_p, final_df_s)

        # Split the final DataFrame
        df_bino = final_df[final_df['Category'] == 'Bino']
        df_else = final_df[final_df['Category'] != 'Bino']

        st.markdown(get_table_download_link(df_bino, df_else, Date_End, WeekNumUseStr), unsafe_allow_html=True)

else:
    st.write("No report type selected")

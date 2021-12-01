
# Import libraries
import streamlit as st
import numpy as np
import pandas as pd
import base64
from io import BytesIO
import io
import datetime as dt
# import locale
# locale.setlocale( locale.LC_ALL, 'en_ZA.UTF-8' )
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

st.title('Retailer Sales Reports - USA')

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
    ('Please select','FYE','Giant_Tiger'))
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


# FYE
if option == 'FYE':

    try:
    # Get retailers map
        df_fye_retailers_map = df_map
        df_fye_retailers_map = df_fye_retailers_map.rename(columns={'FYE UPC':'UPC'})
        
        # Get retailer data
        df_fye_data = df_data
                
        # Merge with retailer map
        df_fye_data_merged = df_fye_data.merge(df_fye_retailers_map, how='left', on='UPC')
        
        # Find missing data
        missing_model_fye = df_fye_data_merged['SMD code'].isnull()
        df_fye_missing_model = df_fye_data_merged[missing_model_fye]
        df_missing = df_fye_missing_model[['UPC','Item Description']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)
        st.write(" ")

    except:
        st.markdown("**Retailer map column headings:** FYE UPC, SMD SKU, SMD Desc, RSP")
        st.markdown("**Retailer data column headings:** Store Name, UPC, Item Description, Unit Sales, EOD Sat On Hand Qty, EOD Sat In Transit Qty")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct") 

        
    try:
        # Set date columns
        df_fye_data_merged['Start Date'] = Date_End
        
        # Total amount column
        df_fye_data_merged['Total Amt'] = df_fye_data_merged['Unit Sales'] * df_fye_data_merged['MSRP']
        
        # Add retailer column and store column
        df_fye_data_merged['Forecast Group'] = 'FYE'
        df_fye_data_merged['Item Description'] = df_fye_data_merged['Item Description'].str.title() 
        df_fye_data_merged['SOH Qty'] = df_fye_data_merged['EOD Sat On Hand Qty'] + df_fye_data_merged['EOD Sat In Transit Qty']

        # Rename columns
        df_fye_data_merged = df_fye_data_merged.rename(columns={'UPC': 'SKU No.'})
        df_fye_data_merged = df_fye_data_merged.rename(columns={'Unit Sales': 'Sales Qty'})
        df_fye_data_merged = df_fye_data_merged.rename(columns={'SMD code': 'Product Code'})
        df_fye_data_merged = df_fye_data_merged.rename(columns={'SMD Desc': 'Product Description'})

        # Don't change these headings. Rather change the ones above
        final_df_fye = df_fye_data_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_fye_p = df_fye_data_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_fye_s = df_fye_data_merged[['Store Name','Total Amt']]

        # Show final df
        total = final_df_fye['Total Amt'].sum()
        total_units = final_df_fye['Sales Qty'].sum()
        st.write('**The total sales for the week are:** $',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_fye_p.groupby(["Product Description"]).agg({"Sales Qty":"sum", "Total Amt":"sum"}).sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'${:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_fye_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('${0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_fye_p.groupby(["Product Description"]).agg({"Sales Qty":"sum", "Total Amt":"sum"}).sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'${:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_fye_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('${0:,.2f}'))

        st.write('**Final Dataframe:**')
        final_df_fye

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_fye), unsafe_allow_html=True)

    except:
        st.write('Check data')

# Giant Tiger
elif option == 'Giant_Tiger':

    try:
        # Get retailers map
        df_gt_retailers_map = df_map

        # Get retailer data
        df_gt_data = df_data
        df_gt_data.columns = df_gt_data.iloc[2]
        df_gt_data = df_gt_data.iloc[3:]
        df_gt_data = df_gt_data[df_gt_data['SKU'].notna()]
                
        # Merge with retailer map
        df_gt_data_merged = df_gt_data.merge(df_gt_retailers_map, how='left', on='SKU')
        
        # Find missing data
        missing_model_gt = df_gt_data_merged['SMD Code'].isnull()
        df_gt_missing_model = df_gt_data_merged[missing_model_gt]
        df_missing = df_gt_missing_model[['SKU','Style']]
        df_missing_unique = df_missing.drop_duplicates()
        st.write("The following products are missing the SMD code on the map: ")
        st.table(df_missing_unique)
        st.write(" ")

    except:
        st.markdown("**Retailer map column headings:** SKU, SMD Code, SMD Description")
        st.markdown("**Retailer data column headings:** SKU, Style, LW Sales Units, LW Sales $, STORE OH, OO, GTW Net Units")
        st.markdown("Column headings are **case sensitive.** Please make sure they are correct") 

        
    try:
        # Set date columns
        df_gt_data_merged['Start Date'] = Date_End
        
        # Total amount column
        df_gt_data_merged = df_gt_data_merged.rename(columns={'LW Sales $':'Total Amt'})
        
        # Add retailer column and store column
        df_gt_data_merged['Forecast Group'] = 'Giant Tiger'
        df_gt_data_merged['Store Name'] = ''
        df_gt_data_merged['Style'] = df_gt_data_merged['Style'].str.title() 
        df_gt_data_merged['SOH Qty'] = df_gt_data_merged['STORE OH'] + df_gt_data_merged['OO'] + df_gt_data_merged['GTW Net Units']
        
        # Rename columns
        df_gt_data_merged = df_gt_data_merged.rename(columns={'SKU': 'SKU No.'})
        df_gt_data_merged = df_gt_data_merged.rename(columns={'LW Sales Units': 'Sales Qty'})
        df_gt_data_merged = df_gt_data_merged.rename(columns={'SMD Code': 'Product Code'})
        df_gt_data_merged = df_gt_data_merged.rename(columns={'SMD Description': 'Product Description'})

        # Don't change these headings. Rather change the ones above
        final_df_gt = df_gt_data_merged[['Start Date','SKU No.', 'Product Code', 'Forecast Group','Store Name','SOH Qty','Sales Qty','Total Amt']]
        final_df_gt_p = df_gt_data_merged[['Product Code','Product Description','Sales Qty','Total Amt']]
        final_df_gt_s = df_gt_data_merged[['Store Name','Total Amt']]

        # Show final df
        total = final_df_gt['Total Amt'].sum()
        total_units = final_df_gt['Sales Qty'].sum()
        st.write('**The total sales for the week are:** $',"{:0,.2f}".format(total).replace(',', ' '))
        st.write('**Number of units sold:** '"{:0,.0f}".format(total_units).replace(',', ' '))
        st.write('')
        st.write('**Top 10 products for the week:**')
        grouped_df_pt = final_df_gt_p.groupby(["Product Description"]).agg({"Sales Qty":"sum", "Total Amt":"sum"}).sort_values("Total Amt", ascending=False)
        grouped_df_final_pt = grouped_df_pt[['Sales Qty', 'Total Amt']].head(10)
        st.table(grouped_df_final_pt.style.format({'Sales Qty':'{:,.0f}','Total Amt':'${:,.2f}'}))
        st.write('')
        st.write('**Top 10 stores for the week:**')
        grouped_df_st = final_df_gt_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_st = grouped_df_st[['Total Amt']].head(10)
        st.table(grouped_df_final_st.style.format('${0:,.2f}'))
        st.write('')
        st.write('**Bottom 10 products for the week:**')
        grouped_df_pb = final_df_gt_p.groupby(["Product Description"]).agg({"Sales Qty":"sum", "Total Amt":"sum"}).sort_values("Total Amt", ascending=False)
        grouped_df_final_pb = grouped_df_pb[['Sales Qty', 'Total Amt']].tail(10)
        st.table(grouped_df_final_pb.style.format({'Sales Qty':'{:,.0f}','Total Amt':'${:,.2f}'}))
        st.write('')
        st.write('**Bottom 10 stores for the week:**')
        grouped_df_sb = final_df_gt_s.groupby("Store Name").sum().sort_values("Total Amt", ascending=False)
        grouped_df_final_sb = grouped_df_sb[['Total Amt']].tail(10)
        st.table(grouped_df_final_sb.style.format('${0:,.2f}'))

        st.write('**Final Dataframe:**')
        final_df_gt

        # Output to .xlsx
        st.write('Please ensure that no products are missing before downloading!')
        st.markdown(get_table_download_link(final_df_gt), unsafe_allow_html=True)

    except:
        st.write('Check data')
else:
    st.write('Retailer not selected yet')

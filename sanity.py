import streamlit as st
import pandas as pd
import pandas as pd
import numpy as np
import io
import base64
from io import BytesIO
#read excel sheet     
df=pd.read_excel(r"C:/Users/shital.ingole/Downloads/FW06 Ecoms Processed Volume Till 10 March.xlsx")
df.rename(columns={"Parts List - Descriptions": "PartsDescriptions"}, inplace=True)




def process_data(df):
    df.rename(columns={"Parts List - Descriptions": "PartsDescriptions"}, inplace=True)
    df.PartsDescriptions = df.PartsDescriptions.fillna(0)
    
    # 1. Top/Bottom Cover Peeling or Chipping
    keywords = ['Latitude 5310', 'Latitude 5320', 'Latitude 5330', 'Latitude 5410', 'Latitude 5411',
                'Latitude 5420', 'Latitude 5421', 'Latitude 5430', 'Latitude 5431', 'Latitude 5511',
                'Latitude 5520', 'Latitude 5521', 'Latitude 5530', 'Latitude 5531']
    df['lob1'] = df['Brand Name'].str.contains('|'.join(keywords), regex=True)
    troubleshoot_keywords = ['Peel', 'Paint', 'peeling off', 'paint off', 'paint damage', 'paint fade']
    df['F21'] = df['Troubleshooting Performed'].str.contains('|'.join(troubleshoot_keywords), regex=True)
    parts_keywords = ['Cover LCD', 'Bottom Door']
    df['f21'] = df['PartsDescriptions'].str.contains('|'.join(parts_keywords), regex=True)

    df['f31'] = df['Request Complete Care'].str.contains('Yes')
    df['f51'] = df['Status Code'].str.contains('Issued Pending|Issued')

    # Assigning Peelingorchipping based on conditions
    df['Peelingorchipping'] = 'compliance'
    df.loc[(df['F21'] & df['f51'] & df['lob1'] & df['f31'] & df['f21']), 'Peelingorchipping'] = 'Peelingorchipping'
    
    # 2. Gen 11 and Gen 12 of Latitude and Precision Notebooks Keyboard Paint Peeling off
    keywords = ['Latitude 3310', 'Latitude 3320', 'Latitude 3410', 'Latitude 3420', 'Latitude 3510',
                'Latitude 3520', 'Latitude 5310', 'Latitude 5320', 'Latitude 5410', 'Latitude 5411',
                'Latitude 5420', 'Latitude 5421', 'Latitude 5510', 'Latitude 5511', 'Latitude 5520',
                'Latitude 5521', 'Latitude 7210', 'Latitude 7310', 'Latitude 7320', 'Latitude 7410',
                'Latitude 7420', 'Latitude 9410', 'Latitude 9420', 'Latitude 9510', 'Latitude 9520',
                'PRECISION 3550', 'PRECISION 3551', 'PRECISION 3561', 'PRECISION 5550', 'PRECISION 5750',
                'PRECISION 7540', 'PRECISION 7550', 'PRECISION 7740', 'PRECISION 7750']
    df['lob2'] = df['Brand Name'].str.contains('|'.join(keywords), regex=True)
    df['F22'] = df['Troubleshooting Performed'].str.contains('Peel|Paint|paint peeling off error|paint off|Paint fade|Key fade|key missing|key letter missing', regex=True)
    df['f22'] = df['PartsDescriptions'].str.contains('Keyboard', regex=True)
    df['f32'] = df['Request Complete Care'].str.contains('Yes', regex=False)
    df['f52'] = df['Status Code'].str.contains('Issued Pending|Issued', regex=True)

    # Assigning keyboardpaintpeeling based on conditions
    df['keyboardpaintpeeling'] = 'compliance'
    df.loc[(df['F22'] & df['f52'] & df['lob2'] & df['f32'] & df['f22']), 'keyboardpaintpeeling'] = 'keyboardpaintpeeling'
    
    # 3. Rubber Strip Peeling Off The Bottom Of Latitude 7310 and 7410 Notebooks
    troubleshoot_keywords = ['Peel', 'rubber', 'peeling off', 'rubber feat', 'rubber come off']
    parts_keywords = ['Bottom door']
    brand_keywords = ['Latitude 7310', 'Latitude 7410']
    df['lob3'] = df['Brand Name'].str.contains('|'.join(brand_keywords), regex=True)
    df['F23'] = df['Troubleshooting Performed'].str.contains('|'.join(troubleshoot_keywords), regex=True)
    df['f23'] = df['PartsDescriptions'].str.contains('|'.join(parts_keywords), regex=True)
    df['f33'] = df['Request Complete Care'].str.contains('Yes', case=False)
    df['f53'] = df['Status Code'].str.contains('Issued Pending|Issued', regex=True)
    df['RubberStrippeelingoff'] = 'compliance'
    # Update 'RubberStrippeelingoff' column based on conditions
    df.loc[(df['F23'] & df['f53'] & df['lob3'] & df['f33'] & df['f23']), 'RubberStrippeelingoff'] = 'RubberStrippeelingoff'
    
    # 4. Bezel Loose
    parts_keywords = ['Bezel', 'The Bezel does not connect properly']
    troubleshoot_keywords = ['Loose', 'not attaching', 'peel off', 'peel', 'not sticking', 'coming off']
    brand_keywords = ['Latitude 7280', 'Latitude E6440', 'Latitude E7240', 'Latitude E7450', 'Latitude 5450',
                      'Latitude 7440', 'Latitude 7250', 'Latitude 5550', 'Latitude E6540', 'PRECISION M2800',
                      'Latitude Exx70', 'Latitude E6440', 'Latitude E6540', 'PRECISION M2800', 'Latitude E7240',
                      'Latitude E7440', 'Latitude E7450', 'Latitude E5450', 'Latitude E7250', 'Latitude E5550',
                      'Latitude Exx70', 'Latitude 7280']
    # Define conditions for each column
    df['lob4'] = df['Brand Name'].str.contains('|'.join(brand_keywords), regex=True)
    df['F24'] = df['Troubleshooting Performed'].str.contains('|'.join(troubleshoot_keywords), regex=True)
    df['f24'] = df['PartsDescriptions'].str.contains('|'.join(parts_keywords), regex=True)
    df['f34'] = df['Request Complete Care'].str.contains('Yes', case=False)
    df['f54'] = df['Status Code'].str.contains('Issued Pending|Issued', regex=True)
    df['Bezelloose'] = 'compliance'
    # Update 'Bezelloose' column based on conditions
    df.loc[(df['lob4'] & df['F24'] & df['f24'] & df['f34'] & df['f54']), 'Bezelloose'] = 'Bezelloose'
    
    # 5. Wear and Tear
    parts_keywords = ['system board', 'systemboard', 'motherboard', 'DC - IN', 'DC-IN', 'DC IN']
    troubleshoot_keywords = ['loose DC', 'loose USB', 'loose HDMI', 'loose video', 'loose port', 'port loose',
                             'DC in loose', 'DC-in loose', 'USB loose', 'HDMI loose', 'wiggle', 'certain position',
                             'certain angle', 'loosing connection', 'loose port']
    # Define conditions for each column
    df['f25'] = df['PartsDescriptions'].str.contains('|'.join(parts_keywords), case=False)
    df['F25'] = df['Troubleshooting Performed'].str.contains('|'.join(troubleshoot_keywords), regex=True)
    df['f35'] = df['Request Complete Care'].str.contains('No', case=False)
    df['f55'] = df['Status Code'].str.contains('Issued Pending|Issued', regex=True)

    # Assign 'compliance' to all rows initially
    df['WearandTear'] = 'compliance'

    # Update 'WearandTear' column based on conditions
    df.loc[(df['F25'] & df['f55'] & df['f35'] & df['f25']), 'WearandTear'] = 'WearandTear'
    
    df = df.assign(Scenario='Nan')
    df.loc[((df['Peelingorchipping'] != 'compliance') | (df['keyboardpaintpeeling'] != 'compliance') | (df['RubberStrippeelingoff'] != 'compliance') | (df['Bezelloose'] != 'compliance') | (df['WearandTear'] != 'compliance')), 'Scenario'] = 'Error'
    
    return df[['Work Order Code', 'Status Code', 'Status Date', 'Brand Name', 'Service Tag', 'DPS Number', 'SP Name',
               'Branch Name', 'Entitlement', 'Request on-site technician', 'Allow Complete Care', 'Allow Return To Depot',
               'Request Complete Care', 'Pro Support', 'Keep Your Hard Drive', 'Override Address Name', 'Override Address City',
               'Override Address State', 'Override Address Postal Code', 'Override Address Time Zone', 'Description', 'CRU/FRU',
               'Order Denied Reason', 'E PSA diagnostic ID', 'E PSA validation code', 'ACCT PO Reference', 'Override Address Country',
               'Track', 'PartsDescriptions', 'Parts List - Components', 'Parts List - Numbers', 'Dispatcher Name', 'Username',
               'Customer', 'Troubleshooting Performed', 'Entercoms ID', 'Dispatcher', 'Primary Contact Name', 'Primary Contact Email',
               'Override Address 1', 'Override Address 2', 'Override Address 3', 'Country (Group)', 'Region Code (Group)',
               'Product Description','Dell Internal Notes', 'Denial Line of Business', 'Denial Reason2',
               'Denial Category', 'Auto Dispatch', 'Swivel Seat Reason', 'Swivel Seat Reason Category', 'Submitter', 'Assigned',
               'Battery Damaged', 'BIL Response', 'Denial Reason', 'OSP', 'Fweek', 'Date','Fweek', 'Peelingorchipping',
               'keyboardpaintpeeling', 'RubberStrippeelingoff', 'Bezelloose', 'WearandTear', 'Scenario']]


###### Main function to run the app
def main():
    st.title("Sanity Automation")
    st.image(r"D:/sanityAutomation/new_image.png", width=250)
    
 # Allow user to upload Excel file
    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            df['DPS Number'] = df['DPS Number'].astype(str)

            processed_df = process_data(df)
            
            # Add a button to download the processed data as an Excel file
            download_link = get_table_download_link(processed_df)
            st.markdown(download_link, unsafe_allow_html=True)
            
        except Exception as e:
            st.write("An error occurred:", e)



# Function to create a download link for a DataFrame as an Excel file


def get_table_download_link(df):
    # Convert DataFrame to Excel file
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sanity_Automation')
    writer.close()
    excel_data = output.getvalue()
    b64 = base64.b64encode(excel_data).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="Sanity_Automation.xlsx">Download Output file as Excel File</a>'
    return href


if __name__ == "__main__":
    main()

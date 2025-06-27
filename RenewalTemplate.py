#import all required libraries
import streamlit as st
import pyodbc
import pandas as pd
import numpy as np
import datetime as dt
from PIL import Image
from dateutil.relativedelta import relativedelta
import textwrap
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os

#set the page configuration
# st.set_page_config(page_title= 'Premium Calculator',layout='wide', initial_sidebar_state='expanded')

#add a image header to the page
# image = Image.open('RenewalPortal.png')
# st.image(image, use_column_width=True)

#write the queries to pull data from the DB
query = 'select * from vw_client_renewal_portal_active_client_data'
query1 = 'select * from vw_tbl_final_client_mlr'
query2 = 'select * from premium_calculator_pa_data'
query3 = 'select * from tbl_renewal_portal_template_module_plan_data a\
            where convert(date, date_submitted) = (select max(convert(date,date_submitted))\
										            from tbl_renewal_portal_template_module_plan_data b\
										            where a.client = b.client)'
query4 = 'select * from tbl_renewal_portal_template_module_client_data a\
            where date_submitted = (select max(date_submitted) from tbl_renewal_portal_template_module_client_data b where a.client = b.client)'


# assign the DB credentials to variables
server = os.environ.get('server_name')
database = os.environ.get('db_name')
username = os.environ.get('db_username')
password = os.environ.get('db_password')

# define the DB connection
conn = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};SERVER='
        + server
        +';DATABASE='
        + database
        +';UID='
        + username
        +';PWD='
        + password
        )

# #define the connection for the DBs
# conn = pyodbc.connect(
#         'DRIVER={ODBC Driver 17 for SQL Server};SERVER='
#         +st.secrets['server']
#         +';DATABASE='
#         +st.secrets['database']
#         +';UID='
#         +st.secrets['username']
#         +';PWD='
#         +st.secrets['password']
#         ) 

#write a function to read the data from the DBs
@st.cache_data(ttl = dt.timedelta(hours=24))
def get_data_from_sql():  
    active_clients = pd.read_sql(query, conn)
    client_mlr = pd.read_sql(query1, conn)
    pa_data = pd.read_sql(query2, conn)
    return active_clients, client_mlr, pa_data

#assign the data to variables as below
active_clients, client_mlr, pa_data = get_data_from_sql()

#read data from already filled client renewals and assigned to variables as shown
plan_renewal_df = pd.read_sql(query3, conn)
client_renewal_df = pd.read_sql(query4, conn)


#convert some certain columns from the 3 tables to the required data type
cols_to_convert = ['MemberNo', 'PolicyNo']
pa_data[cols_to_convert] = pa_data[cols_to_convert].astype(str)
pa_data['PAIssueDate'] = pd.to_datetime(pa_data['PAIssueDate']).dt.date

# client_mlr['TxDate'] = pd.to_datetime(client_mlr['TxDate']).dt.date
client_mlr['PolicyNo'] = client_mlr['PolicyNo'].astype(str)
columns_to_convert = ['PREMIUM', 'CLAIMS', 'CAPITATION', 'ADDITIONAL_PA', 'TOTAL_MEDICAL', 'MLR']
for col in columns_to_convert:
    client_mlr[col] = pd.to_numeric(client_mlr[col], errors='coerce').fillna(0)

active_clients[cols_to_convert] = active_clients[cols_to_convert].astype(int).astype(str)

def validate_non_empty(value, field_name):
    if not value:
        st.error(f"{field_name} cannot be empty.")
        return False
    return True

#write a function to groupby required columns and format the result
def calc_n_format_percent(data, col, value, total, agg_op):
    col_aggregate = data.groupby(col).agg({
    value: agg_op
    }).reset_index()
    #calculate the percentage utilization by col.
    col_aggregate['Percentage'] = round((col_aggregate[value] / total) * 100, 2)
    # Sort the data by Percentage in decreasing order
    col_aggregate = col_aggregate[col_aggregate['Percentage'] > 1]
    col_aggregate = col_aggregate.sort_values(by='Percentage', ascending=False)
    col_percentage = col_aggregate[[col, 'Percentage']]
    # Format the percentage values
    col_percentage['Percentage'] = col_percentage['Percentage'].map('{:.2f}%'.format)
    # Concatenate each record with dash and commas
    formatted_col_percentage = ',\n'.join(
    col_percentage.apply(lambda x: f'{x[col]} - {x["Percentage"]}', axis=1)
    )
    return formatted_col_percentage

def extract_percentage_utilization(plan_name, plan_utilization):
    for entry in plan_utilization.split(','):
        parts = entry.strip().split(' - ')
        if len(parts) == 2 and parts[0] == plan_name:
                return float(parts[1][:-1]) # Convert '30.00%' to 30.00 as a float
    return None  # Return None if the plan_name is not found

#function to retrieve existing data from invoice and template module

def get_existing_plan_data(client_name, conn):
    query = """
    SELECT DISTINCT 
        a.PolicyNo, 
        a.client, 
        bb.PlanName, 
        bb.PremiumType, 
        bb.Count, 
        bb.PremiumAmt AS Premium, 
        cc.category, 
        cc.upsell_last_3yrs, 
        cc.upsell_year, 
        cc.upsell_note,
        cc.repriced_last_3yrs, 
        cc.year_repriced, 
        cc.repriced_percent, 
        a.total_lives, 
        total_premium, 
        contact_person_title, 
        address, 
        state
    FROM tbl_renewal_portal_invoice_module_client_data a
    JOIN tbl_renewal_portal_invoice_module_plan_data bb 
        ON a.PolicyNo = LEFT(bb.PlanID, CHARINDEX('-', bb.PlanID) - 1) 
        AND a.invoiceno = bb.invoiceno
    LEFT JOIN tbl_renewal_portal_template_module_plan_data cc 
        ON bb.PlanID = cc.PlanID 
        AND bb.PlanName = cc.plan_name 
        AND cc.year_repriced != 0
    WHERE submitted_date = (
        SELECT MAX(submitted_date)
        FROM tbl_renewal_portal_invoice_module_client_data b
        WHERE a.client = b.client
    )
    AND a.client = ?
    """

    query1 = """
    select distinct a.PolicyNo, a.client, total_lives_client_mgr, b.total_premium,client_onboarding_yr, shared_portfolio, Competitor_HMO, AdditionalNotes
	from tbl_renewal_portal_template_module_client_data a
	join tbl_renewal_portal_invoice_module_client_data b on a.PolicyNo = b.PolicyNo
	where b.submitted_date = (SELECT MAX(submitted_date)
        FROM tbl_renewal_portal_invoice_module_client_data b
        WHERE a.client = b.client)
	and a.date_submitted = (SELECT MAX(date_submitted)
        FROM tbl_renewal_portal_template_module_client_data b
        WHERE a.client = b.client)
	and client_manager != 'BI Admin'
    and a.client = ?
     """

    
    df = pd.read_sql(query, conn, params=[client_name])
    df1 = pd.read_sql(query1, conn, params=[client_name])

    plan_data_dict = {}

    for _, row in df.iterrows():
        plan_name = row['PlanName']
        premium_type = row['PremiumType'].capitalize()  # 'Individual' or 'Family'
        key = f"{plan_name}_{premium_type}"

        plan_data_dict[key] = {
            'category': row.get('category'),
            'i_num_lives': row['Count'] if premium_type == 'Individual' else 0,
            'f_num_lives': row['Count'] if premium_type == 'Family' else 0,
            'i_premium_paid': row['Premium'] if premium_type == 'Individual' else 0,
            'f_premium_paid': row['Premium'] if premium_type == 'Family' else 0,
            'total_lives': row.get('total_lives', 0),
            'upsell': row.get('upsell_last_3yrs', 'No'),
            'upsell_yr': row.get('upsell_year'),
            'upsell_notes': row.get('upsell_note'),
            'repriced': row.get('repriced_last_3yrs', 'No'),
            'repriced_yr': row.get('year_repriced'),
            'repriced_percent': row.get('repriced_percent')
        }

    # Add client-level data
    if not df1.empty:
        client_data = df1.iloc[0]
        plan_data_dict['client_info'] = {
            'total_premium': client_data['total_premium'],
            'client_onboarding_yr': client_data['client_onboarding_yr'],
            'shared_portfolio': client_data['shared_portfolio'],
            'Competitor_HMO': client_data['Competitor_HMO'],
            'AdditionalNotes': client_data['AdditionalNotes']
        }

    return plan_data_dict


# Function to dynamically generate input fields based on the number of plans
def generate_input_fields(client, plan_names, existing_plan_data=None, enable_editing=True):
    plan_data = []
    for i, plan_name in enumerate(plan_names):
        # Display the plan name at the top of each input section
        st.markdown(f"### Plan: **{plan_name}**")

        # Use the plan_data_dict key format: "{plan_name}_Individual" and "{plan_name}_Family"
        plan_ind_key = f"{plan_name}_Individual"
        plan_fam_key = f"{plan_name}_Family"

        # Merge individual and family data for the plan
        existing_ind = existing_plan_data.get(plan_ind_key, {}) if existing_plan_data else {}
        existing_fam = existing_plan_data.get(plan_fam_key, {}) if existing_plan_data else {}

        # Compose a merged dict for easier access
        merged = {
            'category': existing_ind.get('category') or existing_fam.get('category'),
            'i_num_lives': existing_ind.get('i_num_lives', 0),
            'f_num_lives': existing_fam.get('f_num_lives', 0),
            'i_premium_paid': existing_ind.get('i_premium_paid', 0),
            'f_premium_paid': existing_fam.get('f_premium_paid', 0),
            'total_lives': existing_ind.get('total_lives', 0) or existing_fam.get('total_lives', 0),
            'upsell': existing_ind.get('upsell', 'No') or existing_fam.get('upsell', 'No'),
            'upsell_yr': existing_ind.get('upsell_yr') or existing_fam.get('upsell_yr'),
            'upsell_notes': existing_ind.get('upsell_notes') or existing_fam.get('upsell_notes'),
            'repriced': existing_ind.get('repriced', 'No') or existing_fam.get('repriced', 'No'),
            'repriced_yr': existing_ind.get('repriced_yr') or existing_fam.get('repriced_yr'),
            'repriced_percent': existing_ind.get('repriced_percent') or existing_fam.get('repriced_percent'),
        }

        # All fields are always editable; removed the editable checkbox

        plan_category = st.radio(
            label='Is the Plan Customised or Standard?',
            key=f'{plan_name}_category',
            index=0 if merged.get('category') == 'Standard' else 1 if merged.get('category') == 'Customised' else None,
            options=['Standard', 'Customised'],
        )

        i_num_lives = st.number_input(
            'Number of Individual Plans',
            step=1,
            key=f'{plan_name}_ilives',
            value=merged.get('i_num_lives', 0),
        )
        i_premium_paid = st.number_input(
            'Premium for Individual Plan',
            key=f'{plan_name}_ipremium',
            value=merged.get('i_premium_paid', 0),
        )
        f_num_lives = st.number_input(
            'Number of Family Plans',
            step=1,
            key=f'{plan_name}_flives',
            value=merged.get('f_num_lives', 0),
        )
        f_premium_paid = st.number_input(
            'Premium for Family Plan',
            key=f'{plan_name}_fpremium',
            value=merged.get('f_premium_paid', 0),
        )
        total_lives = st.number_input(
            'Input the total number of Lives (Individual + Family)',
            step=1,
            key=f'{plan_name}_tlives',
            value=merged.get('total_lives') or 0,
        )

        upsell = st.radio(
            label='Was this plan upsold in the last 3 years?',
            key=f'{plan_name}_upsell',
            options=['No', 'Yes'],
            index=0 if merged.get('upsell') == 'No' else 1 if merged.get('upsell') == 'Yes' else None,
        )
        upsell_yr = st.number_input(
            'If Upsell is Yes, Input Year Plan was Upsold',
            min_value=None,
            max_value=dt.date.today().year,
            key=f'{plan_name}_upsell_yr',
            value=merged.get('upsell_yr') or 2020,
        )
        upsell_note = st.text_input(
            'Provide Relevant Details of the Upgrade',
            key=f'{plan_name}_upsell_note',
            value=merged.get('upsell_notes') or '',
        )

        repriced = st.radio(
            label='Has this plan been repriced?',
            key=f'{plan_name}_repriced',
            options=['No', 'Yes'],
            index=0 if merged.get('repriced') == 'No' else 1 if merged.get('repriced') == 'Yes' else None,
        )
        repriced_yr = st.number_input(
            f'If {plan_name} was Repriced, Input Repriced Year',
            min_value=None,
            max_value=dt.date.today().year,
            key=f'{plan_name}_repriced_yr',
            value=merged.get('repriced_yr') or 2020,
        )
        repriced_percent = st.number_input(
            f'By what % was {plan_name} repriced?',
            key=f'{plan_name}_repriced_percent',
            value=merged.get('repriced_percent') or 0,
        )

        policyid = active_clients.loc[active_clients['PolicyName'] == client, 'PolicyNo'].values[0]

        i_baserate = active_clients.loc[
            (active_clients['PolicyName'] == client) &
            (active_clients['PlanType'] == plan_name), 'IndividualBaseRate'
        ].values[0]
        f_baserate = active_clients.loc[
            (active_clients['PolicyName'] == client) &
            (active_clients['PlanType'] == plan_name), 'FamilyBaseRate'
        ].values[0]
        i_circulationrate = active_clients.loc[
            (active_clients['PolicyName'] == client) &
            (active_clients['PlanType'] == plan_name), 'IndividualCirculationRate'
        ].values[0]
        f_circulationrate = active_clients.loc[
            (active_clients['PolicyName'] == client) &
            (active_clients['PlanType'] == plan_name), 'FamilyCirculationRate'
        ].values[0]

        if (i_num_lives is not None and i_num_lives > 0) or (f_num_lives is not None and f_num_lives > 0):
            plan_data.append({
                'client': client,
                'plan_id': str(policyid) + '-' + str(i+1),
                'plan_name': plan_name,
                'category': plan_category,
                'i_num_lives': i_num_lives,
                'f_num_lives': f_num_lives,
                'i_premium_paid': i_premium_paid,
                'f_premium_paid': f_premium_paid,
                'total_lives': total_lives,
                'repriced': repriced,
                'repriced_yr': repriced_yr,
                'upsell': upsell,
                'upsell_yr': upsell_yr,
                'upsell_notes': upsell_note,
                'iBaseRate': i_baserate,
                'fBaseRate': f_baserate,
                'iCirculationRate': i_circulationrate,
                'fCirculationRate': f_circulationrate,
                'repriced_percent': repriced_percent
            })
    return plan_data
#function to clear the plan_data when it has been submitted and written to the DB
def reset_data():
    plan_data.clear()  
    repricing_metrics.clear()

#function to assign scores and recommendations based on the inputted data
def assign_scores_n_recommendation(plan_data, mlr, utili,premium):
    recommendations = []
    for plan in plan_data:
        plan_utilization = extract_percentage_utilization(plan['plan_name'], utili)
        if plan_utilization is not None:
            score = 5 *(plan_utilization/100)
        else:
            score = 0
        if int(premium) < 5000000:
            score += 5
        elif 5000000 <= int(premium) < 10000000:
            score += 4
        elif 10000000 <= int(premium) < 50000000:
            score += 3
        elif 50000000 <= int(premium) < 100000000:
            score += 2
        elif int(premium) > 100000000:
            score += 1


        if (mlr is not None) and (mlr > 0 and mlr < 50):
            score += 2
        elif (mlr is not None) and (mlr >= 50 and mlr < 70):
            score += 4
        elif (mlr is not None) and (mlr >= 70 and mlr < 100):
            score += 8
        elif (mlr is not None) and mlr >= 100:
            score += 16
        else:
            score += 0    

        if plan['upsell'] == 'No':
            if plan['repriced'] == 'No':
                score += 12
            elif (plan['repriced'] == 'Yes') and (plan['repriced_percent'] < 10):
                score += 8
            elif (plan['repriced'] == 'Yes') and (10 <= plan['repriced_percent'] < 25):
                score += 6
            elif (plan['repriced'] == 'Yes') and (plan['repriced_percent'] >= 25):
                score += 4
        else:
            score += 2     

        if plan['category'] == 'Standard':
            i_score = 0
            if plan['i_premium_paid'] > 0:
                if plan['i_premium_paid'] < plan['iBaseRate']:
                    i_score += 10                    
                elif plan['i_premium_paid'] == plan['iBaseRate']:
                    i_score += 6
                elif plan['i_premium_paid'] > plan['iBaseRate']:
                    i_score += 2
            else:
                i_score += 0

            f_score = 0
            if plan['f_premium_paid'] > 0:
                if plan['f_premium_paid'] < plan['fBaseRate']:
                    f_score += 10
                elif plan['f_premium_paid']== plan['fBaseRate']:
                    f_score += 6
                elif plan['f_premium_paid'] > plan['fBaseRate']:
                    f_score += 2
            else:
                f_score += 0

            # Check if both i_premium_paid and f_premium_paid are not None
            if (plan['i_premium_paid'] > 0) and (plan['f_premium_paid'] > 0):
                avg_score = (i_score + f_score) / 2
                score += avg_score
            elif plan['i_premium_paid'] > 0:
                score += i_score
            elif plan['f_premium_paid'] > 0:
                score += f_score
        elif plan['category'] == 'Customised':
            score += 0

        #calculate the recommendation
        if (plan['category'] == 'Standard' and score >= 35) or (plan['category'] == 'Customised' and score >= 25):
            rec = 40
        elif (plan['category'] == 'Standard' and 30 <= score < 35) or (plan['category'] == 'Customised' and 20 <= score < 25):
            rec = 30
        elif (plan['category'] == 'Standard' and 25 <= score < 30) or (plan['category'] == 'Customised' and 15 <= score < 20):
            rec = 25
        elif (plan['category'] == 'Standard' and 20 <= score < 25) or (plan['category'] == 'Customised' and 10 <= score < 15):
            rec = 20
        elif (plan['category'] == 'Standard' and 15 <= score < 20) or (plan['category'] == 'Customised' and 5 <= score < 10):
            rec = 15
        elif (plan['category'] == 'Standard' and score < 15) or (plan['category'] == 'Customised' and  score < 5):
            rec = 10
        else:
            rec = 0
        plan['score'] = score
        plan['recommendation'] = rec

        if plan['i_premium_paid'] > 0:
            premium_increase = (rec/100) * plan['i_premium_paid']
            plan['new_ipremium'] = plan['i_premium_paid'] + premium_increase
        else:
            plan['new_ipremium'] = None

        if plan['f_premium_paid'] > 0:
            premium_increase = (rec/100) * plan['f_premium_paid']
            plan['new_fpremium'] = plan['f_premium_paid'] + premium_increase
        else:
            plan['new_fpremium'] = None

        recommendations.append(plan)
    return recommendations

# --- Add this function below your imports or in a suitable place in your script ---
def notify_finance_team(client, policyno, name):

    finance_email = 'olagoke.dawodu@avonhealthcare.com'
    cc_email_list = ['financedepartment@avonhealthcare.com', email, 'internalauditriskandcontroldept@avonhealthcare.com']
    cc_email = ','.join(cc_email_list)
    bcc_email = 'bi_dataanalytics@avonhealthcare.com'
    myemail = 'noreply@avonhealthcare.com'
    password = os.environ.get('emailpassword')
    excel_path = r"https://avonhealthcareltd.sharepoint.com/:x:/r/sites/BIAnalytics/Shared%20Documents/Finance/R%26G%20Local%20Client%20Renewal%20MLR.xlsx?d=wed5aa667c7b84962a19572e3822534df&csf=1&web=1&e=bxuv60"
    subject = f"MLR Update Required for Client: {client} (Policy No: {policyno})"
    body = f"""
    Dear Olagoke,<br><br>
    This is to notify you that the MLR for the following client is not updated on the Renewal Portal:<br><br>
    <b>Client Name:</b> {client}<br>
    <b>Policy Number:</b> {policyno}<br><br>
    Kindly update the client's MLR in the <a href="{excel_path}">live Excel sheet</a> as soon as possible.<br><br>
    Regards,<br>
    Client Manager - {name}
    """

    # Path to the live Excel sheet (update this path as needed)
    
    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(myemail, password)

        msg = MIMEMultipart()
        msg['From'] = 'AVON HMO Client Services'
        msg['To'] = finance_email
        msg['Cc'] = cc_email
        msg['Bcc'] = bcc_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))

        # # Attach the Excel file
        # with open(excel_path, "rb") as f:
        #     part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        #     part.set_payload(f.read())
        #     encoders.encode_base64(part)
        #     part.add_header('Content-Disposition', f'attachment; filename="Live MLR Update Sheet.xlsx"')
        #     msg.attach(part)

        recipients = [finance_email] + cc_email_list + [bcc_email]
        server.sendmail(myemail, recipients, msg.as_string())
        server.quit()
    except Exception as e:
        st.error(f"Failed to send notification to Finance Team: {e}")
#retrieve the user details from session state
name = st.session_state.get('name', None)
email = st.session_state.get('email', None)

#create a select box on the sidebar to allow users select a client from the active client list ans assigne to a variable
client = st.sidebar.selectbox(label='Select Client', placeholder='Pick a Client', index=None, options=active_clients['PolicyName'].unique())

#a conditional statement to be executed when a client is selected
if client is not None:
    #assign the policyno of the selected client to a variable
    policyno = str(active_clients.loc[active_clients['PolicyName'] == client, 'PolicyNo'].values[0])

    start_date = active_clients.loc[
                        active_clients['PolicyName'] == client,
                        'PolicyStartDate'
                        ].dt.date.unique()
    start_date = np.min(start_date)


    end_date = active_clients.loc[
        active_clients['PolicyName'] == client,
        'PolicyEndDate'
        ].dt.date.unique()
    end_date = np.max(end_date)
    end_date = end_date + relativedelta(months=1)


    # Extract the full month name
    policy_end_month = end_date.strftime('%B')

    #check is policyno is available in client_mlr data
    if policyno in client_mlr['PolicyNo'].values:
        client_revenue = client_mlr.loc[
            (client_mlr['PolicyNo'] == policyno),
            'PREMIUM'
        ].values[0]
        #extract the selected client medical cost data within the policy start and end date
        # from finance data and assign to a variable
        client_medicalcost = client_mlr.loc[
            (client_mlr['PolicyNo'] == policyno), 
            'TOTAL_MEDICAL'
        ].values[0]
        client_mlr = round((client_medicalcost / client_revenue) * 100, 2)
        f_client_revenue = '#' + '{:,}'.format(client_revenue)
        f_client_medicalcost = '#' + '{:,}'.format(client_medicalcost)
    

    # if client_revenue is not None and client_medicalcost is not None:
    #             client_mlr = round((client_medicalcost / client_revenue) * 100, 2)
    #             f_client_revenue = '#' + '{:,}'.format(client_revenue)
    #             f_client_medicalcost = '#' + '{:,}'.format(client_medicalcost)
    # else:
    # # Set client_mlr to a specific value when client_revenue is zero
    #     f_client_revenue = None
    #     f_client_medicalcost = None
    #     client_mlr = None

    #extract the selected client PA utilization within its policy period from the PA data and assign to a variable
        client_pa_data = pa_data.loc[
            (pa_data['PolicyNo'] == policyno) &
            (pa_data['PAIssueDate'] >= start_date) &
            (pa_data['PAIssueDate'] <= end_date),
            ['AvonPaCode', 'EnrolleeName', 'MemberNo', 'Gender', 'MemberType', 'PlanName', 'ProviderName', 'Benefits', 'ApprovedPAAmount', 'PAIssueDate']
        ]
        #calculate the total PA value within the selected client's policy period 
        client_pa_value = round(client_pa_data['ApprovedPAAmount'].sum(),2)
        #Calculate the total number of enrollees who accessed care from the client's pa data
        client_pa_count = client_pa_data['MemberNo'].nunique()
        #calculate the total active lives for each selected client and assign to a variable
        total_active_lives = active_clients.loc[
            active_clients['PolicyNo'] == policyno,
            'MemberNo'
        ].nunique()


        client_active_member_df = active_clients[active_clients['PolicyNo'] == policyno]
        #use the function above to calculate the percentage utilization based on gender, membertype, plan and benefits and assign to variables
        gender_utilization = calc_n_format_percent(client_pa_data, 'Gender', 'ApprovedPAAmount', client_pa_value, 'sum' )
        member_utilization = calc_n_format_percent(client_pa_data, 'MemberType', 'ApprovedPAAmount', client_pa_value, 'sum')
        plan_utilization = calc_n_format_percent(client_pa_data, 'PlanName', 'ApprovedPAAmount', client_pa_value, 'sum')
        benefits_utilization = calc_n_format_percent(client_pa_data, 'Benefits', 'ApprovedPAAmount', client_pa_value, 'sum')
        plan_population = calc_n_format_percent(client_active_member_df, 'PlanType', 'MemberNo', total_active_lives, 'nunique')

        #aggregate client pa data by provider's frequency and PA Value
        client_provider_agg = client_pa_data.groupby(['ProviderName']).agg({
            'AvonPaCode':'nunique',
            'ApprovedPAAmount': 'sum'
            }).sort_values(by=['AvonPaCode','ApprovedPAAmount'], ascending=False)
        #calculate the total number of providers accessed by the client
        client_num_of_providers_used = client_provider_agg.index.nunique()
        #retrieve the top 3 providers and assign to a variable
        top_providers = client_provider_agg.head(3)
        client_top_providers = top_providers.index.tolist()
        #format the list to be displayed on a new line and seperated by comma
        formatted_providers = ',\n'.join(client_top_providers)
        
        #check if the selected client already has a record in the client_renewal_df
        if client in client_renewal_df['client'].values:
            plan_df = plan_renewal_df.loc[plan_renewal_df['client'] == client]
            client_df = client_renewal_df.loc[client_renewal_df['client'] == client]
            sub_date = pd.to_datetime(client_df['date_submitted'].values[0])
            #convert to a date format e.g 21st July, 2021
            sub_date = sub_date.strftime('%d %B, %Y')
            st.subheader(f"{client} Renewal Template has been pre-populated based on the Last Invoice Module Data and Last Submitted Template Module Data on {sub_date}. Kindly review and make necessary changes before submitting the renewal template.")
            # for index, plan in plan_df.iterrows():
            #         st.info(f"Plan Name: {plan['plan_name']}\n\n"
            #                 f"Plan Category: {plan['category']}\n\n"
            #                 f"Plan Upsell: {plan['upsell_last_3yrs']}\n\n"
            #                 f"Plan Repriced Last 3 Years?: {plan['repriced_last_3yrs']}\n\n"
            #                 f"Plan Repriced Percentage: {plan['repriced_percent']}\n\n"
            #                 f"Individual Lives: {'{:,}'.format(plan['individual_lives'])}\n\n"
            #                 f"Current Individual Premium: {'#' + '{:,}'.format(plan['individual_premium'])}\n\n"
            #                 # f"Recommended Individual Renewal Premium: {'#' + '{:,}'.format(plan['rec_renewal_ipremium'])}\n\n"
            #                 f"Family Lives: {'{:,}'.format(plan['family_lives'])}\n\n"
            #                 f"Family Premium: {'#' + '{:,}'.format(plan['family_premium'])}\n\n"
            #                 # f"Recommended Family Renewal Premium: {'#' + '{:,}'.format(plan['rec_renewal_fpremium'])}\n\n"
            #                 f"Total Lives on this plan: {'{:,}'.format(plan['total_lives'])}")
            
        else:
            st.subheader(f"Renewal Template for {client}")
        #assign the unique plans of the selected client to a variable
        unique_plan = active_clients.loc[active_clients['PolicyName'] == client, 'PlanType'].unique()
        #group the client data by the number of enrollees on each plane and assign to a variable.
        agg = active_clients[active_clients['PolicyName'] == client].groupby(['PlanType']).agg({'MemberNo':'nunique', 'MemberHeadNo': 'nunique'}).sort_values(by=['PlanType',"MemberNo"], ascending=False)
        #Split each client's unique plan into a new line and display the unique plans
        plan_message = ',\n\n'.join(unique_plan)
        formatted_msg = textwrap.dedent(f'{client} has the following active plan(s): \n\n{plan_message}.\n\nKindly provide below the requested information about each of these active plan(s).')
        st.warning(formatted_msg)

        # Get the number of plans for the selected client 
        num_plans = len(unique_plan)  

        # Call the generate_input_fields function to create input fields for the selected client
        # and the corresponding number of plans and assign the list to a variable 
        # st.write(get_existing_plan_data(client, conn))
        with st.form(key='plan_data_form'):
            existing_data = get_existing_plan_data(client, conn)
            plan_data = generate_input_fields(client, unique_plan, existing_plan_data=existing_data, enable_editing=True)

            # Convert plan_data to a pandas df for easier manipulation
            plan_df = pd.DataFrame(plan_data)
            
            #calculate the total premium and total lives from the inputted info by the client mgr
            #check if the plan_df is not empty
            if not plan_df.empty:
                total_premium = (plan_df['i_num_lives'] * plan_df['i_premium_paid']) + (plan_df['f_num_lives'] * plan_df['f_premium_paid'])
                total_calc_premium = total_premium.sum()
                total_lives = plan_df['total_lives'].sum()
            #create a subheader for the second part of the input fields
            st.subheader('Provide Additional Information Below About the Client')
            #specify the other input fields and assign them to variables as below
            year_joined = st.number_input(label='Client Onboarding Year',min_value=2013, max_value=dt.date.today().year, value=existing_data.get('client_info', {}).get('client_onboarding_yr', 2020))
            shared_portfolio = st.radio(label='Is this a shared portfolio?',options=['No', 'Yes'], index=0 if existing_data.get('client_info', {}).get('shared_portfolio', 'No') == 'No' else 1)
            competitor = st.text_input(label='If Portfolio is Shared, List the HMOs we are sharing with', help='If more than one, seperate the names with comma', value=existing_data.get('client_info', {}).get('Competitor_HMO', ''))
            total_actual_premium = st.number_input(f'Input the actual total premium paid by {client}', value=existing_data.get('client_info', {}).get('total_premium', 0), help='This is the total premium paid by the client for all plans')
            upsell_reprice_doc = st.file_uploader('If Client was Repriced or Upsold, Upload Evidence', accept_multiple_files=True)
            notes = st.text_area(label='Additional Notes/Remarks', value=existing_data.get('client_info', {}).get('AdditionalNotes', ''), help='Provide any additional information about the client that is relevant to the renewal process')

            #create a submit button to submit the inputted data
            submit = st.form_submit_button('Preview Renewal Information')

            if submit:
                repricing_metrics = assign_scores_n_recommendation(plan_data,client_mlr,plan_utilization,total_actual_premium)

                display_info_df = pd.DataFrame(repricing_metrics)


                if not display_info_df.empty:
                    display_info = display_info_df[['plan_name', 'category', 'i_num_lives', 'f_num_lives', 'i_premium_paid', 'f_premium_paid',
                                        'total_lives', 'upsell', 'upsell_yr', 'upsell_notes', 'repriced', 'repriced_yr', 'repriced_percent', 'iBaseRate', 'fBaseRate', 'score',
                                        'recommendation', 'new_ipremium', 'new_fpremium']]
                    display_info = display_info.rename(columns={'plan_name':'Plan Name', 'category':'Category', 'i_num_lives': 'No. of Lives on Individual Plan',
                                        'f_num_lives':'Total No. of Family', 'i_premium_paid':'Premium per Individual Plan',
                                        'f_premium_paid':'Premium Per Family Plan', 'total_lives':'Total No. of Lives', 'upsell':'Plan Upsold in the Last 3Years?',
                                        'upsell_yr':'Year Plan was Upsold', 'upsell_notes':'Upsell Additional Info','repriced':'Plan Repriced in the Last 3 Years?', 'repriced_yr':'Last Repriced Year',
                                        'repriced_percent':'Last Repriced Percentage', 'iBaseRate':'Current Base Rate for Individual Plan',
                                        'fBaseRate':'Current BaseRate for Family Plan', 'score':'Total Reprice Score', 'recommendation':'Recommended Reprice Percentage(%)', 
                                        'new_ipremium':'Recommended Premium for Individual Plan', 'new_fpremium':'Recommended Premium for Family Plan'})

                    plan_info = display_info[['Plan Name', 'Category', 'No. of Lives on Individual Plan', 'Total No. of Family', 'Premium per Individual Plan',
                                            'Premium Per Family Plan', 'Total No. of Lives', 'Plan Upsold in the Last 3Years?', 'Year Plan was Upsold',
                                            'Upsell Additional Info', 'Plan Repriced in the Last 3 Years?', 'Last Repriced Year','Last Repriced Percentage'
                                            ]]
                    # Convert DataFrame to HTML table
                    plan_table = plan_info.to_html(index=False, escape=False)

                    #add styling to the plan_html_table
                    plan_html_table = f"""
                    <h2>{client} Plan(s) Renewal Information</h2>
                    <style>
                    table {{
                            border: 1px solid #1C6EA4;
                            background-color: #EEEEEE;
                            width: 100%;
                            text-align: left;
                            border-collapse: collapse;
                            }}
                            table td, table th {{
                            border: 1px solid #AAAAAA;
                            padding: 3px 2px;
                            }}
                            table tbody td {{
                            font-size: 13px;
                            }}
                            table thead {{
                            background: #59058D;
                            border-bottom: 2px solid #444444;
                            }}
                            table thead th {{
                            font-size: 15px;
                            font-weight: bold;
                            color: #FFFFFF;
                            border-left: 2px solid #D0E4F5;
                            }}
                            table thead th:first-child {{
                            border-left: none;
                            }}
                    </style>
                    <table>
                    {plan_table}
                    </table>
                    """

                    st.markdown(plan_html_table, unsafe_allow_html=True)
        
                    st.success("Kindly Review the Client Information above, confirm it's accuracy and click on the Submit button below")
            
            entry = st.form_submit_button('Submit Renewal Information')
            if entry:
                #add the # sign and thousand seperators to client revenue and medical cost
                f_total_actual_premium = '#' + '{:,}'.format(total_actual_premium)
                f_total_active_lives = '{:,}'.format(total_active_lives)    
                f_client_pa_value = '#' + '{:,}'.format(client_pa_value)  

                repricing_metrics = assign_scores_n_recommendation(plan_data,client_mlr,plan_utilization,total_actual_premium)
                display_info_df = pd.DataFrame(repricing_metrics)

                if not display_info_df.empty:
                    display_info = display_info_df[['plan_name', 'category', 'i_num_lives', 'f_num_lives', 'i_premium_paid', 'f_premium_paid',
                                        'total_lives', 'upsell', 'upsell_yr', 'upsell_notes', 'repriced', 'repriced_yr', 'repriced_percent', 'iBaseRate', 'fBaseRate', 'score',
                                        'recommendation', 'new_ipremium', 'new_fpremium']]
                    display_info = display_info.rename(columns={'plan_name':'Plan Name', 'category':'Category', 'i_num_lives': 'No. of Lives on Individual Plan',
                                        'f_num_lives':'Total No. of Family', 'i_premium_paid':'Premium per Individual Plan',
                                        'f_premium_paid':'Premium Per Family Plan', 'total_lives':'Total No. of Lives', 'upsell':'Plan Upsold in the Last 3Years?',
                                        'upsell_yr':'Year Plan was Upsold', 'upsell_notes':'Upsell Additional Info','repriced':'Plan Repriced in the Last 3 Years?', 'repriced_yr':'Last Repriced Year',
                                        'repriced_percent':'Last Repriced Percentage', 'iBaseRate':'Current Base Rate for Individual Plan',
                                        'fBaseRate':'Current BaseRate for Family Plan', 'score':'Total Reprice Score', 'recommendation':'Recommended Reprice Percentage(%)', 
                                        'new_ipremium':'Recommended Premium for Individual Plan', 'new_fpremium':'Recommended Premium for Family Plan'})
                    rec_info = display_info[['Plan Name', 'Premium per Individual Plan', 'Premium Per Family Plan', 'Recommended Reprice Percentage(%)',
                                            'Current Base Rate for Individual Plan', 'Current BaseRate for Family Plan',
                                            'Recommended Premium for Individual Plan', 'Recommended Premium for Family Plan']]
                    
                    html_table = f"""
                    <h2>Client Utilization Summary and Information Within Policy Year</h2>
                    <table border="1">
                        <tr>
                            <th>Information</th>
                            <th>Value</th>
                        </tr>
                        <tr>
                            <td>Policy Period</td>
                            <td>{start_date} to {end_date}</td>
                        </tr>
                        <tr>
                            <td>Current MLR</td>
                            <td>{client_mlr}%</td>
                        </tr>
                        <tr>
                            <td>Total Portfolio Size</td>
                            <td>{f_total_actual_premium}</td>
                        </tr>
                        <tr>
                            <td>Client Onboarding Year</td>
                            <td>{year_joined}</td>
                        </tr>
                        <tr>
                            <td>Total PA Utilization Within Policy Period</td>
                            <td>{f_client_pa_value}</td>
                        </tr>
                        <tr>
                            <td>Total Medical Cost (Claims + Capitation) within Policy Period as confirmed by Finance</td>
                            <td>{f_client_medicalcost}</td>
                        </tr>
                        <tr>
                            <td>Total Revenue Received from Client within the Policy Period as confirmed by Finance</td>
                            <td>{f_client_revenue}</td>
                        </tr>
                        <tr>
                            <td>Total Number of Lives at Last Renewal</td>
                            <td>{total_lives}</td>
                        </tr>
                        <tr>
                            <td>Total Number of Active Lives on TOSHFA</td>
                            <td>{f_total_active_lives}</td>
                        </tr>
                        <tr>
                            <td>Total Number of Enrollees Who Accessed Care Within Policy Period</td>
                            <td>{client_pa_count}</td>
                        </tr>
                        <tr>
                            <td>Total Number of Providers Accessed Within the Policy Period</td>
                            <td>{client_num_of_providers_used}</td>
                        </tr>
                        <tr>
                            <td>Top 3 Providers Accessed</td>
                            <td>{formatted_providers}</td>
                        </tr>
                        <tr>
                            <td>Plan Population Distribution</td>
                            <td>{plan_population}</td>
                        </tr>
                        <tr>
                            <td>Plan Type Utilization Percentage</td>
                            <td>{plan_utilization}</td>
                        </tr>
                        <tr>
                            <td>Gender Utilization</td>
                            <td>{gender_utilization}</td>
                        </tr>
                        <tr>
                            <td>Member Type Utilization Percentage</td>
                            <td>{member_utilization}</td>
                        </tr>
                        <tr>
                            <td>Benefit Utilization</td>
                            <td>{benefits_utilization}</td>
                        </tr>
                        <tr>
                            <td>Is the Client a Shared Portfolio?</td>
                            <td>{shared_portfolio}</td>
                        </tr>
                        <tr>
                            <td>Additional Comments</td>
                            <td>{notes}</td>
                        </tr>
                        <tr>
                            <td>Client Manager</td>
                            <td>{name}</td>
                        </tr>
                    </table><br><br>
                """
                    #add the same styling in the plan_html_table to the html_table
                    html_table = html_table.replace('<table>', '<table style="border: 1px solid #1C6EA4; background-color: #EEEEEE; width: 100%; text-align: left; border-collapse: collapse;">')
                    html_table = html_table.replace('<th>', '<th style="background: #59058D; border-bottom: 2px solid #444444; font-size: 15px; font-weight: bold; color: #FFFFFF; border-left: 2px solid #D0E4F5;">')
                    html_table = html_table.replace('<td>', '<td style="border: 1px solid #AAAAAA; padding: 3px 2px; font-size: 13px;">')

                    html_code = """
                                <div style="width: 1000px; margin: 20px; padding: 10px; border: 1px solid #ccc;">
                                    <h2 style="text-align: center;">Detailed below is the client's current premium and the recommended renewal premium based on their utilization metrics within their policy year</h2>
                                    <table border="1" cellpadding="10" style="width: 100%;">
                                        <thead>
                                            <tr style="text-align: left;">
                                                {}
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {}
                                        </tbody>
                                    </table>
                                </div>
                            """

                            # Format the HTML code with dynamic content
                    table_header = "".join(['<th>{}</th>'.format(col) for col in rec_info.columns])
                    #format the header using the same format as the plan_html_table
                    table_header = table_header.replace('<th>', '<th style="background: #59058D; border-bottom: 2px solid #444444; font-size: 15px; font-weight: bold; color: #FFFFFF; border-left: 2px solid #D0E4F5;">')
                    table_body = "".join(['<tr>{}</tr>'.format("".join(['<td>{}</td>'.format(row[col]) for col in rec_info.columns])) for _, row in rec_info.iterrows()])
                    
                    #present the recommendation information in a table format and format in same style as the other tables
                    #displayed above
                    rec_msg = html_code.format(table_header, table_body)
                
                cursor = conn.cursor()
                try:
                    #iterate the plan_data list and write the information for each plan as below
                    # to the table created on the DB.
                    for plan in repricing_metrics:
                        cursor.execute('insert into [dbo].[tbl_renewal_portal_template_module_plan_data]\
                                    (PlanID, PolicyNo, client, plan_name, category, individual_lives, individual_premium, family_lives,\
                                    family_premium, total_lives, upsell_last_3yrs, upsell_year, upsell_note, repriced_last_3yrs, year_repriced,\
                                    repriced_percent, i_BaseRate, f_BaseRate, i_CirculationRate, f_CirculationRate,\
                                    repricing_score, rec_reprice_percent, rec_renewal_ipremium, rec_renewal_fpremium, date_submitted)\
                                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                                    (plan['plan_id'], policyno, client, plan['plan_name'],plan['category'], plan['i_num_lives'], plan['i_premium_paid'], 
                                    plan['f_num_lives'], plan['f_premium_paid'], plan['total_lives'], plan['upsell'], plan['upsell_yr'], plan['upsell_notes'],
                                    plan['repriced'], plan['repriced_yr'], plan['repriced_percent'],
                                    None if pd.isna(plan['iBaseRate']) else float(plan['iBaseRate']), None if pd.isna(plan['fBaseRate']) else float(plan['fBaseRate']),
                                    None if pd.isna(plan['iCirculationRate']) else float(plan['iCirculationRate']),
                                    None if pd.isna(plan['fCirculationRate']) else float(plan['fCirculationRate']),
                                    plan['score'], plan['recommendation'], plan['new_ipremium'], plan['new_fpremium'], dt.datetime.now())
                                    )
                        
                    #insert all the other required information relating to the selected client into the 
                    #created client table on the DB
                    cursor.execute("insert into [dbo].[tbl_renewal_portal_template_module_client_data]\
                                (PolicyNo, client, total_num_of_plans, total_lives_client_mgr, total_calc_premium, total_actual_premium, client_mlr,\
                                total_pa_value, total_medical_cost, total_revenue, toshfa_active_lives, total_enrollees_accessed_care, num_of_providers, \
                                top_3_providers_accessed, pop_distribution, plan_utilization, gender_utilization, member_type_utilization, client_onboarding_yr,\
                                shared_portfolio, competitor_HMO, policy_start_date, policy_end_date, AdditionalNotes, client_manager, date_submitted)\
                                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                                (policyno, client, int(num_plans), int(total_lives), int(total_calc_premium), int(total_actual_premium),
                                None if pd.isna(client_mlr) else float(client_mlr), None if pd.isna(client_pa_value) else float(client_pa_value),
                                None if pd.isna(client_medicalcost) else float(client_medicalcost), None if pd.isna(client_revenue) else float(client_revenue),
                                total_active_lives, client_pa_count, client_num_of_providers_used, formatted_providers, plan_population, plan_utilization, gender_utilization,
                                member_utilization, int(year_joined), shared_portfolio, competitor, start_date, end_date, notes, name, dt.datetime.now())
                                )
                    #commit to insert the data into respective tables
                    conn.commit()
                    #display the text after the successful writing of the data to the DB.
                    st.success(f'All {client} Renewal Information Submitted Sucessfully')

                    # cc_email_list = ['bi_dataanalytics@avonhealthcare.com', email]
                    renewal_year = dt.datetime.now().year
                    subject = f'{renewal_year} RENEWAL NOTIFICATION for {client}'
                    # Create a table (HTML format) with some sample data
                    msg_befor_table = f'''
                    Dear IARC,<br><br>
                    Trust this mail finds you well.<br><br>
                    This is to notify you that {client}'s policy is due for renewal in <b>{policy_end_month}</b> and their renewal process has been initiated.<br><br>

                    The tables below details their plan information, a summary of their utilization within their policy year and the recommendation for renewal.
                    '''

                    msg_after_tables = f"""
                    Kindly review the information and the recommendation by the portal and revert to the Client Manager - <b>{name}</b> to kickstart renewal negotiation with the client.
                    The Client Manager will be available to provide additional information about the client if required.<br><br>
                    Regards.<br><br>
                    <span style="font-size: 20;"><b>{name}</b></span>
                    """
                    clientmgr_msg = f'''
                    Dear {name},<br><br>
                    Trust this mail finds you well.<br><br>
                    This is to notify you that you have successfully initiated the renewal process for {client}.<br><br>
                    The IARC team has been notified and will reach out to you shortly with a premium advice for the client.<br><br>
                    Kindly follow-up to ensure the renewal process is concluded in due time.<br><br>
                    Regards.<br><br>
                    '''

                    audit_message = msg_befor_table + html_table + rec_msg + msg_after_tables
                    myemail = 'noreply@avonhealthcare.com'
                    password = os.environ.get('emailpassword')
                    audit_email = 'internalauditriskandcontroldept@avonhealthcare.com'
                    bcc_email = 'ademola.atolagbe@avonhealthcare.com'
                    # Dynamically set cc_email based on logged-in user
                    if name and name.strip().lower() == 'emmanuel.adigwe':
                        cc_email = 'opeoluwa.adeola@avonhealthcare.com'
                    else:
                        cc_email = 'ajibola.bakare@avonhealthcare.com'
                
                    recipient_1 = [audit_email, cc_email, bcc_email]
                    recipient_2 = [email, bcc_email]

                    try:
                        server = smtplib.SMTP('smtp.office365.com', 587)
                        server.starttls()

                        #login to outlook account
                        server.login(myemail, password)

                        #create a MIMETesxt object for the email message
                        msg = MIMEMultipart()
                        msg['From'] = 'AVON HMO Client Services'
                        msg['To'] = audit_email
                        msg['Cc'] = cc_email
                        msg['Bcc'] = bcc_email
                        msg['Subject'] = subject
                        msg.attach(MIMEText(audit_message, 'html'))

                        msg1 = MIMEMultipart()
                        msg1['From'] = 'AVON HMO Client Services'
                        msg1['To'] = email
                        msg['Bcc'] = bcc_email
                        msg1['Subject'] = subject
                        msg1.attach(MIMEText(clientmgr_msg, 'html'))

                        for file in upsell_reprice_doc:
                            file.seek(0)
                            file_data = file.read()
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(file_data)
                            encoders.encode_base64(part)
                            part.add_header('Content-Disposition', f'attachment; filename={file.name}')
                            msg.attach(part)

                        # all_email = email + cc_email_list
                        server.sendmail(myemail, recipient_1, msg.as_string())
                        server.sendmail(myemail, recipient_2, msg1.as_string())
                        server.quit()

                        st.success(f"{client}'s Renewal Process has been successfully initiated and Notification Email has been sent to all stakeholders.\n\n\
                                    You are advised to follow-up to ensure the renewal process is concluded in due time")
                    except Exception as e:
                        st.error(f'An error occurred: {e}')
                    #call the function below to clear the plan_data list for a new entry
                    reset_data()
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
                finally:
                    cursor.close()
                    conn.close()
        

            
        #block of codes to be executed if no client is selected.
    else:
        st.error(f"{client} does not have their MLR updated. Kindly click the button below to notify the Finance Team to update the MLR for the Client")
        if st.button('Notify Finance Team'):
            notify_finance_team(client, policyno, name)
            st.success(f"Finance Team has been notified to update {client}'s MLR. Kindly follow up with them to ensure the MLR is updated before proceeding with the renewal process.")

else:
    st.info('Select a Client and Client Manager to proceed')


    
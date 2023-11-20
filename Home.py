import streamlit as st
import pyodbc
import pandas as pd
import numpy as np
import datetime as dt
from PIL import Image
import textwrap
import os


st.set_page_config(page_title= 'Premium Calculator',layout='wide', initial_sidebar_state='expanded')

image = Image.open('Avon.png')
st.image(image, use_column_width=False)

query = 'select * from tblEnrolleePremium'
query1 = 'select * from vw_tbl_client_mlr'

server = os.environ.get('server_name')
database = os.environ.get('db_name')
username = os.environ.get('db_username')
password = os.environ.get('db_password')

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

@st.cache_data(ttl = dt.timedelta(hours=24))
def get_data_from_sql():  
    active_clients = pd.read_sql(query, conn)
    client_mlr = pd.read_sql(query1, conn)
    return active_clients, client_mlr

active_clients, client_mlr = get_data_from_sql()

# Function to dynamically generate input fields based on the number of plans
def generate_input_fields(client, plan_names):
    plan_data = []
    for i, plan_name in enumerate(plan_names):
        st.subheader(f'{plan_name} for {client}')
        i_num_lives = st.number_input(f'Number of Individual Lives on {plan_name}', value=None)
        i_premium_paid = st.number_input(f'Premium for Individual Lives on {plan_name}', value=None)
        f_num_lives = st.number_input(f'Number of Family Lives on {plan_name}', value=None)
        f_premium_paid = st.number_input(f'Premium for Family Lives on {plan_name}', value=None)
        total_lives = st.number_input(f'Input the total number of Lives on {plan_name}', value=None)
        repriced_key = f'{client}_{plan_name}_repriced'
        repriced = st.radio(label='Has this plan been repriced in the Last 3 Years?', key=repriced_key, options=['No', 'Yes'])
        if repriced == 'Yes':
            repriced_yr = st.number_input(label=f'What year was {plan_name} repriced?',min_value=2019, max_value=dt.date.today().year)
            repriced_percent = st.number_input(label=f'By what % was {plan_name} repriced?')
        else:
            repriced_yr = 0
            repriced_percent = 0
        plan_data.append({'client': client, 'plan_name': plan_name, 'i_num_lives': i_num_lives, 'f_num_lives': f_num_lives, 
                           'i_premium_paid': i_premium_paid, 'f_premium_paid': f_premium_paid,'total_lives': total_lives,
                            'repriced':repriced, 'repriced_yr':repriced_yr, 'repriced_percent':repriced_percent})
    return plan_data

def reset_data():
    plan_data.clear()  

client_mlr['TxDate'] = pd.to_datetime(client_mlr['TxDate']).dt.date
client_mlr['PolicyNo'] = client_mlr['PolicyNo'].astype(str)

client = st.sidebar.selectbox(label='Select Client', placeholder='Pick a Client', index=None, options=active_clients['PolicyName'].unique())


# plan = st.sidebar.selectbox(label= 'Select Plan', options=unique_plan)
# client_df = active_clients.loc[
#     (active_clients['PolicyName'] == client) &
#     (active_clients['PlanType'] == plan),
#     ['Name', 'MemberNo', 'Relation','PremiumType','Premium']
#     ].reset_index(drop=True)
if client is not None:
    policyno = str(active_clients.loc[active_clients['PolicyName'] == client, 'PolicyNo'].values[0])

    unique_plan = active_clients.loc[active_clients['PolicyName'] == client, 'PlanType'].unique()


    agg = active_clients[active_clients['PolicyName'] == client].groupby(['PlanType']).agg({'MemberNo':'nunique', 'MemberHeadNo': 'nunique'}).sort_values(by=['PlanType',"MemberNo"], ascending=False)

    active_clients['MemberNo'] = active_clients['MemberNo'].astype(int).astype(str)

    client_mgr = st.sidebar.selectbox(label='Select Client Manager',placeholder='Select Your Name', index=None, options=['Abraham Owunna', 'Adebola Oreagba', 'Adebisi Idowu',
                                                                  'Christiana Etok', 'Chibuzor Emelogu','Jennifer Udofia',
                                                                  'Opeoluwa Adeola', 'Tolulope Palmer', 'Victor Agbelese'
                                                                  ])

    plan_message = ',\n\n'.join(unique_plan)
    formatted_msg = textwrap.dedent(f'{client} has the following active plan(s): \n\n{plan_message}.\n\nKindly provide below the requested information about each of these active plan(s).')
    st.info(formatted_msg)

    # Get the number of plans for the selected client 
    num_plans = len(unique_plan)  

    # Initialize a list to store plan information
    plan_data = generate_input_fields(client, unique_plan)

    # Convert plan_data to a DataFrame for easier manipulation
    plan_df = pd.DataFrame(plan_data)


    total_premium = (plan_df['i_num_lives'] * plan_df['i_premium_paid']) + (plan_df['f_num_lives'] * plan_df['f_premium_paid'])
    total_calc_premium = total_premium.sum()
    total_lives = plan_df['total_lives'].sum()

    st.subheader('Provide Additional Information Below About the Client')

    year_joined = st.number_input(label='Client Onboarding Year',min_value=2013, max_value=dt.date.today().year)
    shared_portfolio = st.radio(label='Is this a shared portfolio?', options=['No', 'Yes'])
    if shared_portfolio == 'Yes':
        competitor = st.text_input(label='List the HMOs we are sharing the portfolio with', help='If more than one, seperate the names with comma')
    else:
        competitor = None
    total_actual_premium = st.number_input(f'Input the actual total premium paid by {client}', value=None)
    notes = st.text_area(label='Additional Notes/Remarks')



    start_date = active_clients.loc[
        active_clients['PolicyName'] == client,
        'FromDate'
        ].dt.date.unique()
    start_date = np.min(start_date)

    end_date = active_clients.loc[
        active_clients['PolicyName'] == client,
        'ToDate'
        ].dt.date.unique()
    end_date = np.max(end_date)

    client_revenue_data = client_mlr.loc[
        (client_mlr['PolicyNo'] == policyno) &
        (client_mlr['TxDate'] >= start_date) &
        (client_mlr['TxDate'] <= end_date) &
        (client_mlr['Category'] == 'Revenue'),
        'Net_Amount'
    ]

    client_medicalcost_data = client_mlr.loc[
        (client_mlr['PolicyNo'] == policyno) &
        (client_mlr['TxDate'] >= start_date) &
        (client_mlr['TxDate'] <= end_date) &
        (client_mlr['Category'] == 'Cost of Sales'),
        'Net_Amount'
    ]

    client_revenue = round(client_revenue_data.sum(),2)
    client_medicalcost = round(client_medicalcost_data.sum(),2)
    # Check if client_revenue is zero before calculating client_mlr
    if client_revenue != 0:
        client_mlr = round((client_medicalcost / client_revenue) * 100, 2)
    else:
    # Set client_mlr to a specific value when client_revenue is zero
        client_mlr = None

        # Button to preview data before submission
    if st.button('Preview Client Renewal Information'):
        st.subheader(f"{client}'s Plan Renewal Information:")
        for plan in plan_data:
            st.info(f"Plan Name: {plan['plan_name']}\n\n Individual Lives: {plan['i_num_lives']}\n\n Individual Premium: {plan['i_premium_paid']}\n\nFamily Lives: {plan['f_num_lives']}\n\n Family Premium: {plan['f_premium_paid']}\n\nTotal Lives on this plan: {plan['total_lives']}")
        st.info(f'Additonal {client} Renewal Information\n\nPolicy Period: {start_date} to {end_date}\n\nCurrent MLR: {client_mlr}%\n\nTotal Premium Paid: {total_actual_premium}\n\nShared Portfolio: {shared_portfolio}\n\nAdditional Comments: {notes}')
        st.success("Kindly Review the Client Information above and confirm it's accuracy and click on the Submit button below")

    if st.button('Submit'):
        cursor = conn.cursor()
        for plan in plan_data:
            cursor.execute('insert into [dbo].[plan_renewal_template_data]\
                        (PolicyNo, client, plan_name, individual_lives, individual_premium, family_lives,\
                            family_premium, total_lives, repriced_last_3yrs, year_repriced,\
                            repriced_percent, date_submitted)\
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)',
                        (policyno, client, plan['plan_name'], plan['i_num_lives'], plan['i_premium_paid'],
                            plan['f_num_lives'], plan['f_premium_paid'], plan['total_lives'], plan['repriced'],
                            plan['repriced_yr'], plan['repriced_percent'], dt.datetime.now())
                        )
        
        cursor.execute("insert into [dbo].[client_renewal_template_data]\
                    (PolicyNo, client, total_num_of_plans, total_lives, total_calc_premium, \
                    total_actual_premium, client_mlr, client_onboarding_yr, shared_portfolio, competitor_HMO,\
                    policy_start_date, policy_end_date, AdditionalNotes, client_manager, date_submitted)\
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                    (policyno, client, int(num_plans), int(total_lives), int(total_calc_premium), int(total_actual_premium),
                        None if pd.isna(client_mlr) else float(client_mlr), int(year_joined), shared_portfolio, competitor, start_date, end_date, notes, client_mgr, dt.datetime.now())
                    )
        conn.commit()
        st.success(f'All {client} Renewal Information Submitted Sucessfully')
        reset_data()

    conn.close()

else:
    st.info('Select a client to proceed')
# st.write(f'The total premium for this plan is {total_premium}')

# st.markdown(f'Policy Start Date: {start_date}<br>Policy End Date: {end_date}', unsafe_allow_html=True)
# st.info(f'Policy Start Date: {start_date}\nPolicy End Date: {end_date}')

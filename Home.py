#import all required libraries
import streamlit as st
import pyodbc
import pandas as pd
import numpy as np
import datetime as dt
from PIL import Image
import textwrap
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os

#set the page configuration
st.set_page_config(page_title= 'Premium Calculator',layout='wide', initial_sidebar_state='expanded')

#add a image header to the page
image = Image.open('Avon.png')
st.image(image, use_column_width=False)

#write the queries to pull data from the DB
query = 'select * from tblEnrolleePremium'
query1 = 'select * from vw_tbl_client_mlr'
query2 = 'select * from premium_calculator_pa_data'

#assign the DB credentials to variables
server = os.environ.get('server_name')
database = os.environ.get('db_name')
database1 = os.environ.get('db1_name')
username = os.environ.get('db_username')
password = os.environ.get('db_password')

#define the DB connection
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
conn1 = pyodbc.connect(
        'DRIVER={ODBC Driver 17 for SQL Server};SERVER='
        + server
        +';DATABASE='
        + database1
        +';UID='
        + username
        +';PWD='
        + password
        )

#define the connection for the DBs
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

# conn1 = pyodbc.connect(
#         'DRIVER={ODBC Driver 17 for SQL Server};SERVER='
#         +st.secrets['server']
#         +';DATABASE='
#         +st.secrets['database1']
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
    pa_data = pd.read_sql(query2, conn1)
    return active_clients, client_mlr, pa_data

#assign the data to variables as below
active_clients, client_mlr, pa_data = get_data_from_sql()

#convert some certain columns from the 3 tables to the required data type
cols_to_convert = ['MemberNo', 'PolicyNo']
pa_data[cols_to_convert] = pa_data[cols_to_convert].astype(str)
pa_data['PAIssueDate'] = pd.to_datetime(pa_data['PAIssueDate']).dt.date

client_mlr['TxDate'] = pd.to_datetime(client_mlr['TxDate']).dt.date
client_mlr['PolicyNo'] = client_mlr['PolicyNo'].astype(str)

active_clients[cols_to_convert] = active_clients[cols_to_convert].astype(int).astype(str)

# Function to dynamically generate input fields based on the number of plans
def generate_input_fields(client, plan_names):
    plan_data = []
    for i, plan_name in enumerate(plan_names):
        st.subheader(f'{plan_name}')
        repriced_key = f'{plan_name}_repriced'
        upsell_key = f'{plan_name}_upsell'
        plan_category_key = f'{plan_name}_category'
        i_num_lives_key = f'{plan_name}_ilives'
        i_premium_key = f'{plan_name}_ipremium'
        f_num_lives_key = f'{plan_name}_flives'
        f_premium_key = f'{plan_name}_fpremium'
        total_lives_key = f'{plan_name}_tlives'

        plan_category = st.radio(label='Is the Plan Customised or Standard?', key=plan_category_key, index=None, options=['Standard', 'Customised'])
        i_num_lives = st.number_input(f'Number of Individual Plans', step=1, key=i_num_lives_key, value=None)
        i_premium_paid = st.number_input(f'Premium for Individual Plan', key=i_premium_key, value=None)
        f_num_lives = st.number_input(f'Number of Family Plans', step=1, key=f_num_lives_key, value=None)
        f_premium_paid = st.number_input(f'Premium for Family Plan', key=f_premium_key, value=None)
        total_lives = st.number_input(f'Input the total number of Lives (Individual Lives + Family Lives)', step=1, key=total_lives_key, value=None)
        upsell = st.radio(label='Was this plan upsold from a Lower Plan in the Last 3 Years?', key=upsell_key, index=None, options=['No','Yes'])
        if upsell == 'Yes':
            upsell_yr = st.number_input(label='Year Plan was Upsold', min_value=2019, max_value=dt.date.today().year)
            upsell_note = st.text_input(label='Provide Relevant Details of the Upgrade(Previous Plan and Revenue etc)')
        else:
            upsell_yr = None
            upsell_note = None
        repriced = st.radio(label='Has this plan been repriced in the Last 3 Years?', key=repriced_key, index=None, options=['No', 'Yes'])
        if repriced == 'Yes':
            repriced_yr = st.number_input(label=f'What year was {plan_name} repriced?',min_value=2019, max_value=dt.date.today().year)
            repriced_percent = st.number_input(label=f'By what % was {plan_name} repriced?')
        else:
            repriced_yr = 0
            repriced_percent = 0
        
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
        plan_data.append({'client': client, 'plan_id': str(policyid) + '-' + str(i+1), 'plan_name': plan_name, 'category': plan_category,
                        'i_num_lives': i_num_lives, 'f_num_lives': f_num_lives, 'i_premium_paid': i_premium_paid,
                        'f_premium_paid': f_premium_paid, 'total_lives': total_lives, 'repriced':repriced, 'repriced_yr':repriced_yr,
                        'upsell':upsell, 'upsell_yr':upsell_yr, 'upsell_notes':upsell_note, 'iBaseRate': i_baserate,
                        'fBaseRate': f_baserate,'iCirculationRate': i_circulationrate, 'fCirculationRate': f_circulationrate,
                        'repriced_percent':repriced_percent})
    return plan_data
#function to clear the plan_data when it has been submitted and written to the DB
def reset_data():
    plan_data.clear()  
    repricing_metrics.clear()

#create a select box on the sidebar to allow users select a client from the active client list ans assigne to a variable
client = st.sidebar.selectbox(label='Select Client', placeholder='Pick a Client', index=None, options=active_clients['PolicyName'].unique())

#a conditional statement to be executed when a client is selected
if client is not None:
    #assign the policyno of the selected client to a variable
    policyno = str(active_clients.loc[active_clients['PolicyName'] == client, 'PolicyNo'].values[0])
    #assign the unique plans of the selected client to a variablr
    unique_plan = active_clients.loc[active_clients['PolicyName'] == client, 'PlanType'].unique()
    #group the client data by the number of enrollees on each plane and assign to a variable.
    agg = active_clients[active_clients['PolicyName'] == client].groupby(['PlanType']).agg({'MemberNo':'nunique', 'MemberHeadNo': 'nunique'}).sort_values(by=['PlanType',"MemberNo"], ascending=False)
    #List all the client managers and assign the selected client manager to a variable
    client_mgr = st.sidebar.selectbox(label='Select Client Manager',placeholder='Select Your Name', index=None, options=['Abraham Owunna', 'Adebola Oreagba', 'Adebisi Idowu',
                                                                  'Christiana Etok', 'Chibuzor Emelogu','Jennifer Udofia',
                                                                  'Opeoluwa Adeola', 'Tolulope Palmer', 'Victor Agbelese'
                                                                  ])
    #Split each client's unique plan into a new line and display the unique plans
    plan_message = ',\n\n'.join(unique_plan)
    formatted_msg = textwrap.dedent(f'{client} has the following active plan(s): \n\n{plan_message}.\n\nKindly provide below the requested information about each of these active plan(s).')
    st.info(formatted_msg)

    # Get the number of plans for the selected client 
    num_plans = len(unique_plan)  

    # Call the generate_input_fields function to create input fields for the selected client
    # and the corresponding number of plans and assign the list to a variable 
    plan_data = generate_input_fields(client, unique_plan)

    # Convert plan_data to a pandas df for easier manipulation
    plan_df = pd.DataFrame(plan_data)

    #calculate the total premium and total lives from the inputted info by the client mgr
    total_premium = (plan_df['i_num_lives'] * plan_df['i_premium_paid']) + (plan_df['f_num_lives'] * plan_df['f_premium_paid'])
    total_calc_premium = total_premium.sum()
    total_lives = plan_df['total_lives'].sum()
    #create a subheader for the second part of the input fields
    st.subheader('Provide Additional Information Below About the Client')
    #specify the other input fields and assign them to variables as below
    year_joined = st.number_input(label='Client Onboarding Year',min_value=2013, max_value=dt.date.today().year)
    shared_portfolio = st.radio(label='Is this a shared portfolio?',index=None, options=['No', 'Yes'])
    if shared_portfolio == 'Yes':
        competitor = st.text_input(label='List the HMOs we are sharing the portfolio with', help='If more than one, seperate the names with comma')
    else:
        competitor = None
    total_actual_premium = st.number_input(f'Input the actual total premium paid by {client}', value=0)
    notes = st.text_area(label='Additional Notes/Remarks')
    #extract the policy start and end date for the selected clients from the
    #active_clients table and assign to a variable.
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

    #extract the selected client revenue data within the policy start and end date
    # from finance data and assign to a variable
    client_revenue_data = client_mlr.loc[
        (client_mlr['PolicyNo'] == policyno) &
        (client_mlr['TxDate'] >= start_date) &
        (client_mlr['TxDate'] <= end_date) &
        (client_mlr['Category'] == 'Revenue'),
        'Net_Amount'
    ]
    #extract the selected client medical cost data within the policy start and end date
    # from finance data and assign to a variable
    client_medicalcost_data = client_mlr.loc[
        (client_mlr['PolicyNo'] == policyno) &
        (client_mlr['TxDate'] >= start_date) &
        (client_mlr['TxDate'] <= end_date) &
        (client_mlr['Category'] == 'Cost of Sales'),
        'Net_Amount'
    ]
    #Sum the selected client revenue and medical cost within the policy period and assign to variables
    client_revenue = round(client_revenue_data.astype(float).sum(),2)
    client_medicalcost = round(client_medicalcost_data.astype(float).sum(),2)
    
    # Check if client_revenue is zero before calculating client_mlr
    if client_revenue != 0:
        client_mlr = round((client_medicalcost / client_revenue) * 100, 2)
    else:
    # Set client_mlr to a specific value when client_revenue is zero
        client_mlr = None
    #add the # sign and thousand seperators to client revenue and medical cost
    f_client_revenue = '#' + '{:,}'.format(client_revenue)
    f_client_medicalcost = '#' + '{:,}'.format(client_medicalcost)
    f_total_actual_premium = '#' + '{:,}'.format(total_actual_premium)
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

    #write a function to groupby required columns and format the result
    def calc_n_format_percent(data, col, value, total, agg_op):
        col_aggregate = data.groupby(col).agg({
        value: agg_op
        }).reset_index()
        #calculate the percentage utilization by gender.
        col_aggregate['Percentage'] = round((col_aggregate[value] / total) * 100, 2)
        # Sort the data by Percentage in decreasing order
        col_aggregate = col_aggregate.sort_values(by='Percentage', ascending=False)
        col_percentage = col_aggregate[[col, 'Percentage']]
        # Format the percentage values
        col_percentage['Percentage'] = col_percentage['Percentage'].map('{:.2f}%'.format)
        # Concatenate each record with dash and commas
        formatted_col_percentage = ',\n'.join(
        col_percentage.apply(lambda x: f'{x[col]} - {x["Percentage"]}', axis=1)
        )
        return formatted_col_percentage
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

    #add the thousand seperator for values above 1000 for easier read
    f_total_active_lives = '{:,}'.format(total_active_lives)    
    f_client_pa_value = '#' + '{:,}'.format(client_pa_value)

    def extract_percentage_utilization(plan_name, plan_utilization):
        # Given that plan_utilization has been converted to a string.
        # Extract percentage utilization for the given plan_name from plan_utilization with this function
        for entry in plan_utilization.split(','):
            parts = entry.strip().split(' - ')
            if len(parts) == 2 and parts[0] == plan_name:
                return float(parts[1][:-1])  # Convert '30.00%' to 30.00 as a float
        return None  # Return None if the plan_name is not found

    #SCORING AND RECOMMENDATION LOGIC.
    def assign_scores_n_recommendation(plan_data, mlr, utili,premium):
        total_utilization_score = 10
        if premium is not None:
            if premium < 5000000:
                premium_score = 5
            elif 5000000 <= premium < 10000000:
                premium_score = 4
            elif 10000000 <= premium < 50000000:
                premium_score = 3
            elif 50000000 <= premium < 100000000:
                premium_score = 2
            elif premium > 100000000:
                premium_score = 1
        else:
            premium_score = 0
        
        recommendations = []

        for plan in plan_data:
            score = 0
            plan_percentage_utilization = extract_percentage_utilization(plan['plan_name'], utili)
            
            if (mlr is not None) and (mlr > 0 and mlr < 50):
                score += 2
            elif (mlr is not None) and (mlr >= 50 and mlr < 70):
                score += 4
            elif (mlr is not None) and (mlr >= 70 and mlr < 100):
                score += 8
            elif (mlr is not None) and (mlr >= 100):
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

            # if (plan['repriced'] == 'Yes') and (plan['repriced_percent'] < 10):
            #     score += 6
            # elif (plan['repriced'] == 'Yes') and (10 <= plan['repriced_percent'] < 25):
            #     score += 4
            # elif (plan['repriced'] == 'Yes') and (plan['repriced_percent'] >= 25):
            #     score += 2
            # elif plan['repriced'] == 'No':
            #     score += 10
            # else:
            #     score += 0

            if plan_percentage_utilization is not None:
                    score += total_utilization_score * (plan_percentage_utilization/100)

            if plan['category'] == 'Standard' and ((plan['i_premium_paid'] is not None) or (plan['f_premium_paid'] is not None)):
                i_score = 0
                if (plan['i_premium_paid'] is not None) and (plan['i_premium_paid'] < plan['iBaseRate']):
                    i_score += 10
                elif (plan['i_premium_paid'] is not None) and (plan['i_premium_paid'] == plan['iBaseRate']):
                    i_score += 6
                elif (plan['i_premium_paid'] is not None) and (plan['i_premium_paid'] > plan['iBaseRate']):
                    i_score += 2

                f_score = 0
                if (plan['f_premium_paid'] is not None) and (plan['f_premium_paid'] < plan['fBaseRate']):
                    f_score += 10
                elif (plan['f_premium_paid'] is not None) and (plan['f_premium_paid'] == plan['fBaseRate']):
                    f_score += 6
                elif (plan['f_premium_paid'] is not None) and (plan['f_premium_paid'] > plan['fBaseRate']):
                    f_score += 2

                            # Check if both i_premium_paid and f_premium_paid are not None
                if (plan['i_premium_paid'] is not None) and (plan['f_premium_paid'] is not None):
                    avg_score = (i_score + f_score) / 2
                    score += avg_score
                elif plan['i_premium_paid'] is not None:
                    score += i_score
                elif plan['f_premium_paid'] is not None:
                    score += f_score

            score += premium_score

            #calculate the recommendation
            if (plan['category'] == 'Standard' and score >= 40) or (plan['category'] == 'Customised' and score >= 30):
                rec = 50
            elif (plan['category'] == 'Standard' and 35 <= score < 40) or (plan['category'] == 'Customised' and 25 <= score < 30):
                rec = 30
            elif (plan['category'] == 'Standard' and 21 <= score < 35) or (plan['category'] == 'Customised' and 11 <= score < 25):
                rec = 20
            elif (plan['category'] == 'Standard' and 10 <= score < 21) or (plan['category'] == 'Customised' and 8 <= score < 11):
                rec = 10
            else:
                rec = 0

            plan['score'] = score
            plan['recommendation'] = rec

            if plan['i_premium_paid'] is not None:
                premium_increase = (plan['recommendation']/100) * plan['i_premium_paid']
                plan['new_ipremium'] = plan['i_premium_paid'] + premium_increase
            else:
                plan['new_ipremium'] = None

            if plan['f_premium_paid'] is not None:
                premium_increase = (plan['recommendation']/100) * plan['f_premium_paid']
                plan['new_fpremium'] = plan['f_premium_paid'] + premium_increase
            else:
                plan['new_fpremium'] = None

            recommendations.append(plan)

        return recommendations                  
    
    repricing_metrics = assign_scores_n_recommendation(plan_data,client_mlr,plan_utilization,total_actual_premium)

    # st.write(repricing_metrics)

    display_info_df = pd.DataFrame(repricing_metrics)


    display_info = display_info_df[['plan_name', 'category', 'i_num_lives', 'f_num_lives', 'i_premium_paid', 'f_premium_paid',
                        'total_lives', 'upsell', 'upsell_yr', 'upsell_notes', 'repriced', 'repriced_yr', 'repriced_percent', 'iBaseRate', 'fBaseRate', 'score',
                        'recommendation', 'new_ipremium', 'new_fpremium']]
    display_info = display_info.rename(columns={'plan_name':'Plan Name', 'category':'Category', 'i_num_lives': 'Total No. of Individual Plan',
                        'f_num_lives':'Total No. of Family Plan', 'i_premium_paid':'Premium per Individual Plan',
                        'f_premium_paid':'Premium Per Family Plan', 'total_lives':'Total No. of Lives', 'upsell':'Plan Upsold in the Last 3Years?',
                        'upsell_yr':'Year Plan was Upsold', 'upsell_notes':'Relevant Details on the Upsell','repriced':'Plan Repriced in the Last 3 Years?', 'repriced_yr':'Last Repriced Year',
                        'repriced_percent':'Last Repriced Percentage', 'iBaseRate':'Current Base Rate for Individual Plan',
                        'fBaseRate':'Current BaseRate for Family Plan', 'score':'Total Reprice Score', 'recommendation':'Recommended Reprice Percentage(%)', 
                        'new_ipremium':'New Recommended Premium for Individual Plan', 'new_fpremium':'New Recommended Premium for Family Plan'})

    # Convert DataFrame to HTML table
    plan_html_table = display_info.to_html(index=False)

    # Display HTML table using st.markdown()
    # st.markdown(plan_html_table, unsafe_allow_html=True)


    html_table = f"""
    <h2>Additonal Client Renewal Information</h2>
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
            <td>Total Number of Active Lives as Inputed by Client Manager</td>
            <td>{total_lives}</td>
        </tr>
        <tr>
            <td>Total Number of Active Lives from TOSHFA</td>
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
            <td>{client_mgr}</td>
        </tr>
    </table>
"""

    # Display the HTML table in Streamlit
    # st.markdown(html_table, unsafe_allow_html=True)

        
    # add a button to preview data before submission
    if st.button('Preview Client Renewal Information'):
        st.subheader(f"{client}'s Plan Renewal Information:")
        #iterate the plan_data and extract the information for each plans for the selected client
        #display the information for review

        st.markdown(plan_html_table, unsafe_allow_html=True)
        
        st.markdown(html_table, unsafe_allow_html=True)

        # for plan in plan_data:
        #     st.info(f"Plan Name: {plan['plan_name']}\n\n"
        #             f"Plan Category: {plan['category']}\n\n"
        #             f"Individual Lives: {plan['i_num_lives']}\n\n"
        #             f"Individual Premium: {plan['i_premium_paid']}\n\n"
        #             f"Family Lives: {plan['f_num_lives']}\n\n"
        #             f"Family Premium: {plan['f_premium_paid']}\n\n"
        #             f"Total Lives on this plan: {plan['total_lives']}")
        #Extract the additional information and display as below for review
        # st.info(f"Additonal {client} Renewal Information\n\n"
        #         f"Policy Period: {start_date} to {end_date}\n\n"
        #         f"Current MLR: {client_mlr}%\n\n"
        #         f"Total Invoiced Premium: {total_actual_premium}\n\n"
        #         f"Total PA Value within Policy Period: {client_pa_value}\n\n"
        #         f"Total Medical Cost (Claims + Capitation) within Policy Period as confirmed by Finance: {client_medicalcost}\n\n"
        #         f"Total Revenue Received from Client within the Policy Period as confirmed by Finance: {client_revenue}\n\n"
        #         f"Total Number of Active Lives from Client Manager: {total_lives}\n\n"
        #         f"Total Number of Active Lives from TOSHFA: {total_active_lives}\n\n"
        #         f"Total Number of Enrollees Who Accessed Care Within Policy Period: {client_pa_count}\n\n"
        #         f"Total Number of Active Plans from TOSHFA: {num_plans}\n\n"
        #         f"Total Number of Providers Accessed Within the Policy Period: {client_num_of_providers_used}\n\n"
        #         f"Top 3 Providers Accessed: \n{formatted_providers}\n\n"
        #         f"Gender Utilization: \n{gender_utilization}\n\n"
        #         f"Member Type Utilization: \n{member_utilization}\n\n"
        #         f"Plan Type Utilization: \n{plan_utilization}\n\n"
        #         f"Plan Population Distribution: \n{plan_population}\n\n"
        #         f"Benefits Utilization: \n{benefits_utilization}\n\n"
        #         f"Shared Portfolio: {shared_portfolio}\n\n"
        #         f"Additional Comments: {notes}")
        st.success("Kindly Review the Client Information above, confirm it's accuracy and click on the Submit button below")
    #add a button to submit the information and execute the write of the data to the DB table
    if st.button('Submit'):
        cursor = conn.cursor()
        try:
            #iterate the plan_data list and write the information for each plan as below
            # to the table created on the DB.
            for plan in repricing_metrics:
                cursor.execute('insert into [dbo].[renewal_portal_plan_data]\
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
            cursor.execute("insert into [dbo].[renewal_portal_client_data]\
                        (PolicyNo, client, total_num_of_plans, total_lives_client_mgr, total_calc_premium, total_actual_premium, client_mlr,\
                        total_pa_value, total_medical_cost, total_revenue, toshfa_active_lives, total_enrollees_accessed_care, num_of_providers, \
                        top_3_providers_accessed, pop_distribution, plan_utilization, gender_utilization, member_type_utilization, client_onboarding_yr,\
                        shared_portfolio, competitor_HMO, policy_start_date, policy_end_date, AdditionalNotes, client_manager, date_submitted)\
                            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                        (policyno, client, int(num_plans), int(total_lives), int(total_calc_premium), int(total_actual_premium),
                            None if pd.isna(client_mlr) else float(client_mlr), None if pd.isna(client_pa_value) else float(client_pa_value),
                            None if pd.isna(client_medicalcost) else float(client_medicalcost), None if pd.isna(client_revenue) else float(client_revenue),
                            total_active_lives, client_pa_count, client_num_of_providers_used, formatted_providers, plan_population, plan_utilization, gender_utilization,
                            member_utilization, int(year_joined), shared_portfolio, competitor, start_date, end_date, notes, client_mgr, dt.datetime.now())
                        )
            #commit to insert the data into respective tables
            conn.commit()
            #display the text after the successful writing of the data to the DB.
            st.success(f'All {client} Renewal Information Submitted Sucessfully')
            #call the function below to clear the plan_data list for a new entry
            reset_data()
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
        finally:
            cursor.close()
            conn.close()

            recipient_email = 'idowu.sanni@avonhealthcare.com'
            subject = 'CLIENT RENEWAL NOTIFICATION'
            # Create a table (HTML format) with some sample data
            msg_befor_table = f'''
            Dear Team,<br><br>
            Trust this mail finds you well.<br><br>
            {client} is due renewal on {end_date}.<br><br>

            Find below the details of their key metrics within the policy year
            '''
            final_message = msg_befor_table + plan_html_table + html_table

            myemail = 'ademola.atolagbe@avonhealthcare.com'
            password = os.environ.get('emailpassword')

            try:
                server = smtplib.SMTP('smtp.office365.com', 587)
                server.starttls()

                #login to outlook account
                server.login(myemail, password)

                #create a MIMETesxt object for the email message
                msg = MIMEMultipart()
                msg['From'] = 'AVON HMO Client Services'
                msg['To'] = recipient_email
                msg['Subject'] = subject
                msg.attach(MIMEText(final_message, 'html'))

                server.sendmail(myemail, recipient_email, msg.as_string())
                server.quit()

                st.success('Email with Client Renewal Information has been sent to all stakeholders.')
            except Exception as e:
                st.error(f'An error occurred: {e}')
        #block of codes to be executed if no client is selected.
else:
    st.info('Select a client to proceed')




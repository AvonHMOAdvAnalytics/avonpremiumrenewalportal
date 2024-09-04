#import all required libraries
import streamlit as st
import pyodbc
import pandas as pd
import datetime as dt
from PIL import Image
import os



#set the page configuration
# st.set_page_config(page_title= 'Premium Calculator',layout='wide', initial_sidebar_state='expanded')

#add a image header to the page
image = Image.open('RenewalPortal.png')
st.image(image, use_column_width=True)

#write the queries to pull data from the DB
query5 = 'select distinct a.PolicyNo, b.PolicyName, a.FromDate, a.ToDate, a.ClassName\
        from tblClassMaster a\
        join tblEnrolleePremium b on a.PolicyNo = b.PolicyNo\
        where convert(date,a.ToDate) >= convert(date, getdate())'
query6 = 'select * from tbl_renewal_portal_template_module_plan_data'
query7 = 'select * from tbl_renewal_portal_template_module_client_data'
query8 = 'select * from vw_tbl_final_client_mlr'
query9 = 'select * from premium_calculator_pa_data'

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


#write a function to read the data from the DBs
@st.cache_data(ttl = dt.timedelta(hours=24))
def Loading_data():  
    active_clients = pd.read_sql(query5, conn)
    plan_renewal_df = pd.read_sql(query6, conn)
    client_renewal_df = pd.read_sql(query7, conn)
    client_finance_df = pd.read_sql(query8, conn)
    pa_df = pd.read_sql(query9, conn)
    return active_clients, plan_renewal_df, client_renewal_df, client_finance_df, pa_df

active_clients, plan_renewal_df, client_renewal_df, client_finance_df, pa_df = Loading_data()
client_finance_df['PolicyNo'] = client_finance_df['PolicyNo'].astype(str)

def extract_percentage_utilization(plan_name, plan_utilization):
    # Given that plan_utilization has been converted to a string.
    # Extract percentage utilization for the given plan_name from plan_utilization with this function
    for entry in plan_utilization.split(','):
        parts = entry.strip().split(' - ')
        if len(parts) == 2 and parts[0] == plan_name:
                return float(parts[1][:-1]) # Convert '30.00%' to 30.00 as a float
    return None  # Return None if the plan_name is not found

client = st.sidebar.selectbox(label='Select Client', placeholder='Pick a Client', index=None, options=active_clients['PolicyName'].unique())

st.subheader(f"{client}'s Premium Repricing Recommendation\n\n"
            f"#### Below are the renewal details of the selected client and the recommendation for repricing the premium\n\n")

if client is not None:
    if client in client_renewal_df['client'].unique():
        policyno = str(active_clients.loc[active_clients['PolicyName'] == client, 'PolicyNo'].values[0])
        unique_plan = plan_renewal_df.loc[plan_renewal_df['client'] == client, 'plan_name'].unique()
        selected_client_plan_df = plan_renewal_df[plan_renewal_df['PolicyNo'] == policyno]
        
        #for each plan in plans_details, retrieve 'plan_name', 'category', 'individual_premium', 'family_premium', 'upsell_last_3yrs', 'repriced_last_3yrs', 'i_BaseRate', 'f_BaseRate'
        for plan in unique_plan:
            plan_details = selected_client_plan_df[selected_client_plan_df['plan_name'] == plan]
            category = plan_details['category'].values[0]
            individual_premium = plan_details['individual_premium'].values[0]
            family_premium = plan_details['family_premium'].values[0]
            upsell_last_3yrs = plan_details['upsell_last_3yrs'].values[0]
            repriced_last_3yrs = plan_details['repriced_last_3yrs'].values[0]
            repriced_percent = plan_details['repriced_percent'].values[0]
            i_BaseRate = plan_details['i_BaseRate'].values[0]
            f_BaseRate = plan_details['f_BaseRate'].values[0]
            client_finance_df['client_mlr'] = round((client_finance_df['TOTAL_MEDICAL']/client_finance_df['PREMIUM'])*100,2)
            mlr = client_finance_df.loc[client_finance_df['PolicyNo'] == policyno, 'client_mlr'].values[0]
            premium = client_renewal_df.loc[client_renewal_df['PolicyNo'] == policyno, 'total_actual_premium'].values[0]
            plan_utilization = client_renewal_df.loc[client_renewal_df['PolicyNo'] == policyno, 'plan_utilization'].values[0]
            plan_utilization = extract_percentage_utilization(plan, plan_utilization)

        def assign_scores_n_recommendations():
            score = 5 *(plan_utilization/100)
            if premium < 5000000:
                score += 5
            elif 5000000 <= premium < 10000000:
                score += 4
            elif 10000000 <= premium < 50000000:
                score += 3
            elif 50000000 <= premium < 100000000:
                score += 2
            elif premium > 100000000:
                score += 1


            if mlr > 0 and mlr < 50:
                score += 2
            elif mlr >= 50 and mlr < 70:
                score += 4
            elif mlr >= 70 and mlr < 100:
                score += 8
            elif mlr >= 100:
                score += 16
            else:
                score += 0    

            if upsell_last_3yrs == 'No':
                if repriced_last_3yrs == 'No':
                    score += 12
                elif (repriced_last_3yrs == 'Yes') and (repriced_percent < 10):
                    score += 8
                elif (repriced_last_3yrs == 'Yes') and (10 <= repriced_percent < 25):
                    score += 6
                elif (repriced_last_3yrs == 'Yes') and (repriced_percent >= 25):
                    score += 4
            else:
                score += 2     

            if category == 'Standard':
                i_score = 0
                if individual_premium > 0:
                    if individual_premium < i_BaseRate:
                        i_score += 10                    
                    elif individual_premium == i_BaseRate:
                        i_score += 6
                    elif individual_premium > i_BaseRate:
                        i_score += 2
                else:
                    i_score += 0

                f_score = 0
                if family_premium > 0:
                    if family_premium < f_BaseRate:
                        f_score += 10
                    elif family_premium == f_BaseRate:
                        f_score += 6
                    elif family_premium > f_BaseRate:
                        f_score += 2
                else:
                    f_score += 0

                # Check if both i_premium_paid and f_premium_paid are not None
                if (individual_premium > 0) and (family_premium > 0):
                    avg_score = (i_score + f_score) / 2
                    score += avg_score
                elif individual_premium > 0:
                    score += i_score
                elif family_premium > 0:
                    score += f_score
            elif category == 'Customised':
                score += 0

            #calculate the recommendation
            if (category == 'Standard' and score >= 35) or (category == 'Customised' and score >= 25):
                rec = 40
            elif (category == 'Standard' and 30 <= score < 35) or (category == 'Customised' and 20 <= score < 25):
                rec = 30
            elif (category == 'Standard' and 25 <= score < 30) or (category == 'Customised' and 15 <= score < 20):
                rec = 25
            elif (category == 'Standard' and 20 <= score < 25) or (category == 'Customised' and 10 <= score < 15):
                rec = 20
            elif (category == 'Standard' and 15 <= score < 20) or (category == 'Customised' and 5 <= score < 10):
                rec = 15
            elif (category == 'Standard' and score < 15) or (category == 'Customised' and  score < 5):
                rec = 10
            else:
                rec = 0

            if individual_premium > 0:
                premium_increase = (rec/100) * individual_premium
                new_ipremium = individual_premium + premium_increase
            else:
                new_ipremium = None

            if family_premium > 0:
                premium_increase = (rec/100) * family_premium
                new_fpremium = family_premium + premium_increase
            else:
                new_fpremium = None
            return score, rec, new_ipremium, new_fpremium

            #display all the information above for each plan
        score, rec, new_ipremium, new_fpremium = assign_scores_n_recommendations()
        st.info(f'## {plan} Details\n\n'
                f'#### Category: {category}\n\n'
                f"#### Individual Premium: #{'{:,.2f}'.format(individual_premium) if individual_premium is not None else 'N/A'}\n\n"
                f"#### Family Premium: #{'{:,.2f}'.format(family_premium) if family_premium is not None else 'N/A'}\n\n"
                f"#### Upsell Last 3 Years: {upsell_last_3yrs}\n\n"
                f"#### Repriced Last 3 Years: {repriced_last_3yrs}\n\n"
                f"#### Individual Base Rate: #{'{:,.2f}'.format(i_BaseRate)}\n\n"
                f"#### Family Base Rate: #{'{:,.2f}'.format(f_BaseRate)}\n\n"
                f'#### MLR: {mlr}%\n\n'
                f"#### Premium: #{'{:,.2f}'.format(premium)}\n\n"
                f'#### Plan Utilization: {plan_utilization}%\n\n'
                f'#### Recommendation Score: {score}\n\n'
                f'#### Recommended Reprice Percentage: {rec}%\n\n'
                f"#### Recommended Renewal Premium for Individual Package: #{'{:,.2f}'.format(new_ipremium) if new_ipremium is not None else 'N/A'}\n\n"
                f"#### Recommended Renewal Premium for Family Package: #{'{:,.2f}'.format(new_fpremium) if new_fpremium is not None else 'N/A'}\n\n"
                )

    else:
        st.info('The Portfolio details of the selected Client have not been uploaded by the Client Manager. Kindly contact the Client Manager to Upload')
else:
    st.info('Kindly Select a Client to Proceed')

import streamlit as st
import pandas as pd
import pyodbc
import datetime as dt
import os
 
st.set_page_config(page_title='Client Renewal', layout='wide', initial_sidebar_state='expanded')
 
# # Database connection
# try:
#     conn = pyodbc.connect(
#         'DRIVER={ODBC Driver 17 for SQL Server};SERVER='
#         + st.secrets['server']
#         + ';DATABASE='
#         + st.secrets['database']
#         + ';UID='
#         + st.secrets['username']
#         + ';PWD='
#         + st.secrets['password']
#     )
# except pyodbc.Error as ex:
#     st.error(f"Database connection failed: {ex}")

# assign the DB credentials to variables
server = os.environ.get('server_name')
database = os.environ.get('db_name')
username = os.environ.get('db_username')
password = os.environ.get('db_password')

# define the DB connection
try:
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
except pyodbc.Error as ex:
    st.error(f"Database connection failed: {ex}")

def login_user(username,password):
    with conn.cursor() as cursor:
        cursor.execute("SELECT * FROM tbl_client_renewal_portal_users WHERE UserName = ?", username)
        user = cursor.fetchone()
        if user:
            if password:
                return user[1], user[2], user[3], user[4], user[5], user[6]
            else:
                return None, None, None, None, None, None
        else:
            return None, None, None, None, None, None
    
def create_home_widgets(options):
    choice = st.sidebar.radio("Select Module", options, key='home')
    return choice

# Initialize session state variables if they don't exist
if 'authentication_status' not in st.session_state:
    st.session_state['authentication_status'] = None
if 'name' not in st.session_state:
    st.session_state['name'] = None
if 'username' not in st.session_state:
    st.session_state['username'] = None
if 'password' not in st.session_state:
    st.session_state['password'] = None
if 'user_role' not in st.session_state:
    st.session_state['user_role'] = None
if 'department' not in st.session_state:
    st.session_state['department'] = None
if 'email' not in st.session_state:
    st.session_state['email'] = None

if st.session_state['authentication_status']:
    st.sidebar.title("Home Page")
    st.sidebar.write("Welcome to the Client Renewal Portal!!")
    st.sidebar.write(f"You are currently logged in as {st.session_state['name']} ({st.session_state['username']})")

    #sidebar navigation
    st.sidebar.title("Navigation")
    if st.session_state['department'] == 'Retention and Growth':
        choice = create_home_widgets(["Renewal Template", "Invoice Module","Reconcillation Module"])
    elif st.session_state['department'] == 'Internal Audit':
        choice = create_home_widgets(['Premium Calculator'])
    elif st.session_state['department'] == 'BI and Data Analytics':
        choice = create_home_widgets(['Renewal Template', 'Invoice Module', 'Reconcillation Module', 'Premium Calculator'])
    else:
        st.error('Access Denied: Invalid User Role/Access.')

    # Dynamically load the module based on the username
    try:
        # Define a function to execute a module in a separate namespace
        def execute_module(module_name):
            with open(module_name) as file:
                module_code = file.read()
            module_namespace = {'conn': conn, 'st': st, 'pd': pd, 'dt': dt}
            exec(module_code, module_namespace)

        if choice == "Renewal Template":
            execute_module("RenewalTemplate.py")
        elif choice == "Invoice Module":
            execute_module("Invoice Module.py")
        elif choice == "Reconcillation Module":
            execute_module("ReconcillationModule.py")
        elif choice == "Premium Calculator":
            execute_module("PremiumCalculator.py")
        
    except FileNotFoundError as e:
        st.error(f"Error: {e}")
    except Exception as e:
        st.error(f"An unexpected error(s) occurred: {e}")

    # Add the logout button in the sidebar
    with st.sidebar:
        if st.button('Logout'):
            st.session_state['name'] = None
            st.session_state['authentication_status'] = None
            st.session_state['username'] = None
            st.rerun()
else:
    # Display the login page
    st.sidebar.title("Home Page")
    st.sidebar.write("Welcome to the Portal!")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    
    if st.button("Login"):
        login_username, name, email, user_role, department, login_password = login_user(username, password)
        if username == login_username and password == login_password: 
            st.session_state['authentication_status'] = True
            st.session_state['name'] = name
            st.session_state['username'] = login_username
            st.session_state['password'] = login_password
            st.session_state['department'] = department
            st.session_state['user_role'] = user_role
            st.session_state['email'] = email
            st.rerun()
        else:
            st.error("Username/password is incorrect")



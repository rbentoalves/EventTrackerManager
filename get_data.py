
import pandas as pd 
import numpy as np 
import snowflake.connector
import requests
import pickle
import json

# Snowflake options
ACCOUNT = "lt76872.west-europe.azure"
WAREHOUSE = "ANALYTICS_PRD_WH"
DATABASE = "DATA_LAKE_PRD"
SCHEMA = "CURATED"
ROLE = "POWER_BI_GLOBAL_PERFORMANCE_AND_MONITORING"

# oAuth
AUTH_CLIENT_ID = "d75ec0d2-a525-4807-a6e7-260ed0643853"
AUTH_CLIENT_SECRET = "tvs7Q~asgK_JnLcbV0NEL4VX-DFMQIpC6LZp3"

AUTH_GRANT_TYPE = "client_credentials" #"password", client_credentials

#SCOPE_URL = "https://lightsourcebp.com/session:role-any offline-access"  # mine
SCOPE_URL = "https://lightsourcebp.com/.default"
TOKEN_URL = "https://login.microsoftonline.com/ed5f664a-8752-4c95-8205-40c87d185116/oauth2/v2.0/token"

PAYLOAD = "client_id={clientId}&" \
        "client_secret={clientSecret}&" \
        "grant_type={grantType}&" \
        "scope={scopeUrl}".format(clientId=AUTH_CLIENT_ID, clientSecret=AUTH_CLIENT_SECRET, grantType=AUTH_GRANT_TYPE, scopeUrl=SCOPE_URL)


queries_folder = "C:/Users/ricardo.bento/OneDrive - Lightsource BP/Desktop/Snowflake Queries/"
file_query = "test_query.txt"
custom_query = queries_folder + file_query


def get_data():
    response = requests.post(TOKEN_URL, data=PAYLOAD)
    json_data = json.loads(response.text)
    print(json_data)
    TOKEN = json_data['access_token']
    print("Token obtained")

    # Snowflake connection
    print("connecting to Snowflake")
    conn = snowflake.connector.connect(
                    # user=USER,
                    account=ACCOUNT,
                    role=ROLE,
                    authenticator="oauth",
                    token=TOKEN,
                    warehouse=WAREHOUSE,
                    #database=DATABASE,
                    #schema=SCHEMA,
                    client_session_keep_alive=True,
                    max_connection_pool=20
                    )

    cur = conn.cursor()
    print("connected to snowflake")

    try:
        print("running command")
        cur.execute("SELECT current_version()")
        ret = cur.fetchone()
        print(ret)
        print('Connection successful')
    except snowflake.connector.errors.ProgrammingError as e:
        print(e)

    # Set Warehouse *** choose your warehouse here ***
    sql = '''
    use warehouse TEAM_STRATEGY_WH
    '''
    pd.read_sql(sql, conn)

    # Set Warehouse Size *** choose your warehouse size here ***

    warehouse_size = 'LARGE' #MEDIUM

    sql = f'''
    ALTER WAREHOUSE TEAM_STRATEGY_WH SET WAREHOUSE_SIZE = {warehouse_size}
    '''
    pd.read_sql(sql, conn)

    # get queries

    SQL_QUERY = custom_query

    print("Queries generated")
    data_df = pd.read_sql(SQL_QUERY, conn)

    print("Data retrieved")
    return data_df
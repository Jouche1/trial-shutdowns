import snowflake.connector
import pandas as pd
from config import *

def get_snowflake_connection():
    try:
        conn = snowflake.connector.connect(
            user=SNOWFLAKE_USER,
            password=SNOWFLAKE_PASSWORD,
            account=SNOWFLAKE_ACCOUNT,
            warehouse=SNOWFLAKE_WAREHOUSE,
            database=SNOWFLAKE_DATABASE,
            schema=SNOWFLAKE_SCHEMA
        )
        return conn
    except Exception as e:
        print(f"Error connecting to Snowflake: {e}")
        return None
    
def query_snowflake(document_numbers):
    conn = get_snowflake_connection()
    if not conn:
        return pd.DataFrame()
    
    query = f"""
    WITH serials_cte AS (
      SELECT serial, account_id
      FROM MERAKIDW.FACT.TRIAL_DETAIL_SERIAL
      WHERE document_number IN ({','.join("'" + num + "'" for num in document_numbers)})
    ),
    document_numbers_cte AS (
      SELECT 
        s.serial, 
        COALESCE(hl.document_number, td.document_number) AS document_number,
        CASE 
            WHEN hl.serial IS NOT NULL AND hl.ORGANIZATION_ID NOT IN (55321, 805018433392476227, 589408601232114365) THEN 'Yes'
            WHEN hl.ORGANIZATION_ID IN (55321, 805018433392476227, 589408601232114365) THEN 'ShutDown'
            ELSE 'No' 
        END AS in_cx_network,
        CASE 
            WHEN hl.serial IS NOT NULL THEN hl.ORGANIZATION_ID 
            ELSE NULL 
        END AS current_org_ID,
        hl.ORGANIZATION_NAME AS org_name
    FROM serials_cte s
    LEFT JOIN MERAKIDW.FACT.HARDWARE_LINKAGE hl ON s.serial = hl.serial
    LEFT JOIN MERAKIDW.FACT.TRIAL_DETAIL_SERIAL td ON s.serial = td.serial
    )
    SELECT DISTINCT
        dnc.document_number,
        td.SERIAL,
        td.PROD_CODE,
        CASE
            WHEN td.FREE_TRIAL_EXPIRATION_DATE < CURRENT_DATE THEN 'Expired'
            ELSE 'Not Expired'
        END AS trial_status,
        td.FREE_TRIAL_EXPIRATION_DATE,
        dnc.in_cx_network,
        acc.SUPPORT_SCORE,
        acc.MCN,
        td.ORGANIZATION_ID as original_org_id,
        dnc.current_org_id,
        dnc.org_name,
        CASE
            WHEN edo.ORGANIZATION_ID IS NOT NULL THEN 'Yes'
            ELSE 'No'
        END AS is_ea_org
        
    FROM 
        MERAKIDW.FACT.TRIAL_DETAIL_SERIAL td
    JOIN 
        document_numbers_cte dnc ON td.SERIAL = dnc.serial
    LEFT JOIN
        MERAKIDW.FACT.EA_DASHBOARD_ORGANIZATIONS edo ON td.ORGANIZATION_ID = edo.ORGANIZATION_ID
    JOIN
        MERAKIDW.FACT.ACCOUNT acc ON td.ACCOUNT_ID = acc.account_id;
    """
    try:
        df = pd.read_sql(query, conn)
        conn.close()
        print("Connection closed")
        return df
    except Exception as e:
        print(f"Error executing query: {e}")
        return pd.DataFrame()
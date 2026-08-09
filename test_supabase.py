import streamlit as st
from supabase import create_client

url = st.secrets["SUPABASE_URL"]
key = st.secrets["SUPABASE_KEY"]

supabase = create_client(url, key)

response = (
    supabase
    .table("daily_reports")
    .select("*")
    .limit(1)
    .execute()
)

print("CONEXÃO COM SUPABASE OK")
print(response.data)
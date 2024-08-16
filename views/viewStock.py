import os
import pickle
from pathlib import Path
import streamlit_authenticator as stauth
from office365.sharepoint.files.file import File
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
import streamlit as st
import numpy as np
import pandas as pd
import plotly.express as px
import json
import requests
from PIL import Image
from streamlit_option_menu import option_menu
import plotly.graph_objects as go
from dotenv import load_dotenv


st.set_page_config(
    page_title="Stock Dashboard",
    page_icon=":material/trending_up:",
    layout="wide",
)

# --- USER AUTHENTICATION ---
names = ["Art Huang"]
usernames = ["ericarthuang"]

# load hashed passwords
file_path = Path(__file__).parent/"hashed_pw.pkl"
with file_path.open("rb") as file:
    hashed_passwords = pickle.load(file)

authenticator = stauth.Authenticate(
    names, 
    usernames, 
    hashed_passwords,                                
    "investment_dashboard", 
    "abcdef",
    cookie_expiry_days=0)

name, authentication_status, username = authenticator.login("Login", "main")

if authentication_status == False:
    st.error("Username/password is incorrect")

if authentication_status == None:
    st.warning("Please enter your username and password")

if authentication_status:
    # Load environment variables
    load_dotenv()

    # SharePoint credentials
    sharepoint_site_url = os.getenv("SHAREPOINT_SITE_URL")
    username = os.getenv("SHAREPOINT_USERNAME")
    password = os.getenv("SHAREPOINT_PASSWORD")

    # Authenticate and connect to SharePoint
    ctx_auth = AuthenticationContext(sharepoint_site_url)
    if ctx_auth.acquire_token_for_user(username, password):
        ctx = ClientContext(sharepoint_site_url, ctx_auth)
        web = ctx.web
        ctx.load(web)
        ctx.execute_query()
    else:
        st.error("Failed to authenticate to SharePoint")

    
    # Function to download file from SharePoint
    def download_file_from_sharepoint(file_url):
        response = File.open_binary(ctx, file_url)
        file_name = os.path.basename(file_url)
        with open(file_name, "wb") as local_file:
            local_file.write(response.content)
        return file_name

    @st.cache_data
    def get_data_from_excel(filepath: str, sheet_name: str):
        df = pd.read_excel(
            io=filepath,
            engine="openpyxl",
            sheet_name=sheet_name,
            # skiprows=0,
            # usecols="A:Z",
            header=0,
        )

        return df


    # --- SharePoint and Folder urls ---
    sharepoint_doc = "Shared Documents"
    folder_name = 'investment'
    file_name = 'data_stocks_pure.xlsx'

    folder_url = f"/sites/DashboardDatabase/{sharepoint_doc}/{folder_name}"
    file_url = f"{folder_url}/{file_name}"

    # --- Download the file from SharePoint ---
    file_name = download_file_from_sharepoint(file_url)

    # --- Transform the file to DataFrame --- 
    df = get_data_from_excel(file_name, "Database_Stock")

    # --- Sidebar Welcome ---
    st.sidebar.title(f"Welcome {name}")

    # --- Sidebar Filters ---
    st.sidebar.header("Please Filter Here")

    # --- Selected Dataframe ---
    company = st.sidebar.multiselect(
        label="Select Company",
        options=df["公司名"].unique(),
        default=df["公司名"].unique(),
    )

    stock_name = st.sidebar.multiselect(
        label="Select Stock",
        options=df["股票中文名稱"].unique(),
        default=df["股票中文名稱"].unique(),
    )

    df_selected = df.query(
        '公司名==@company & 股票中文名稱==@stock_name'
    )

    # --- Dataframe Group ---
    groupby_investAmount = df_selected.groupby(
        "公司名")["實際總成本"].sum().reset_index().sort_values(by="實際總成本", ascending=False)

    groupby_marketValue = df_selected.groupby(
        "公司名")["市值估算"].sum().reset_index().sort_values(by="市值估算", ascending=False)

    groupby_stockAmount = df_selected.groupby(
        "股票中文名稱")["實際總成本"].sum().reset_index().sort_values(by="實際總成本", ascending=False)

    groupby_stockValue = df_selected.groupby(
        "股票中文名稱")["市值估算"].sum().reset_index().sort_values(by="市值估算", ascending=False)

    groupby_stockProfitLoss = df_selected.groupby(
        "股票中文名稱")["未實現損益估算"].sum().reset_index().sort_values(by="未實現損益估算", ascending=False)


    # --- Show the KPIs ---
    st.title("📋Investment Overview")

    total_inevstment_amount = int(df_selected["實際總成本"].sum())

    total_market_value = int(df_selected["市值估算"].sum())

    total_profit = total_market_value - total_inevstment_amount

    profit_rate = total_profit / total_inevstment_amount

    left_column, right_column = st.columns(2)
    with left_column:
        st.subheader(f"Total Investment: JP¥  {total_inevstment_amount:,}")
    with right_column:
        st.subheader(f"Total Market Value: JP¥  {total_market_value:,}")

    left_column, right_column = st.columns(2)
    with left_column:
        st.subheader(f"Profit: JP¥  {total_profit:,}")
    with right_column:
        st.subheader(f"Profit Rate:  {profit_rate:.1%}")

    st.markdown("---")

    fig_investAmount = px.bar(
        data_frame=groupby_investAmount,
        x="公司名",
        y="實際總成本",
        text_auto=',.2s',
        color="公司名",
        title="Investment Amount by Company",
        labels={"公司名": "Company", "實際總成本": "Amount"},
    )

    fig_marketValue = px.bar(
        data_frame=groupby_marketValue,
        x="公司名",
        y="市值估算",
        text_auto=',.2s',
        color="公司名",
        title="Market Value by Company",
        labels={"公司名": "Company", "市值估算": "Amount"},
    )

    left_column, right_column = st.columns(2)
    with left_column:
        st.plotly_chart(fig_investAmount)
    with right_column:
        st.plotly_chart(fig_marketValue)

    st.markdown("---")


    fig_stockAmount = px.bar(
        data_frame=groupby_stockAmount,
        x="實際總成本",
        y="股票中文名稱",
        orientation="h",
        text_auto=',.2s',
        color="股票中文名稱",
        title="Investment Amount by Stock",
        labels={"股票中文名稱": "Stock Name", "實際總成本": "Amount"},
        height=800,
    )

    fig_stockValue = px.bar(
        data_frame=groupby_stockValue,
        x="市值估算",
        y="股票中文名稱",
        text_auto=',.2s',
        color="股票中文名稱",
        title="Market Value by Stock",
        labels={"股票中文名稱": "Stock Name", "市值估算": "Amount"},
        height=800,
    )

    left_column, right_column = st.columns(2)
    with left_column:
        st.plotly_chart(fig_stockAmount)
    with right_column:
        st.plotly_chart(fig_stockValue)

    st.markdown("---")

    fig_stockProfitLoss = px.bar(
        data_frame=groupby_stockProfitLoss,
        x="未實現損益估算",
        y="股票中文名稱",
        orientation="h",
        text_auto=',.2s',
        color="股票中文名稱",
        title="Investment Amount by Stock",
        labels={"股票中文名稱": "Stock Name", "未實現損益估算": "Amount"},
        height=800,
    )
    
    st.plotly_chart(fig_stockProfitLoss)

    st.markdown("---")

    # Profit Trend
    st.subheader("The Trend of Profit and Profit Rate")

    df_profit = get_data_from_excel(file_name, "Profit_Trend")
    df_profit['投資報酬率'] = df_profit['投資報酬率'] * 100
    df_profit['日期'] = pd.to_datetime(df_profit['日期'])

    # Set '日期' as index and apply asfreq to ensure daily frequency
    df_profit.set_index('日期', inplace=True)
    df_profit = df_profit.asfreq('D', method='pad')
    df_profit.reset_index(inplace=True)

    startDay = pd.to_datetime(df_profit['日期']).min()
    endDay = pd.to_datetime(df_profit['日期']).max()

    left_column, right_column = st.columns([3,2])
    with left_column:
        col_1, col_2 = st.columns(2)
        with col_1:
            date1 = pd.to_datetime(st.date_input("Start Date", startDay))
        with col_2:
            date2 = pd.to_datetime(st.date_input("End Date", endDay))

    df_profit_selected = df_profit.loc[
        (df_profit['日期'] >= date1) \
        & (df_profit['日期'] <= date2)
    ]

    # --- Create Bar Chart for Profit Trend---
    fig_bar = px.bar(
        df_profit, 
        x='日期', 
        y='未實現損益估算', 
        labels={'未實現損益估算': '未實現損益估算'}
    )

    # Update bar chart color based on profit/loss
    fig_bar.update_traces(
        marker_color=df_profit['未實現損益估算'].apply(lambda x: 'red' if x < 0 else 'blue')
    )

    # --- Create Line Chart for Profit Rate Trend---
    fig_line = px.line(
        df_profit, 
        x='日期', 
        y='投資報酬率', 
        labels={'投資報酬率': '投資報酬率 (%)'}
    )

    # --- Combine Bar and Line Charts ---
    fig_profit = go.Figure()

    # Add bar chart data
    for trace in fig_bar.data:
        fig_profit.add_trace(trace)

    # Add line chart data with secondary y-axis
    for trace in fig_line.data:
        trace.update(yaxis='y2', line=dict(color='white'))
        fig_profit.add_trace(trace)

    # --- Update Layout ---
    fig_profit.update_layout(
        yaxis2=dict(
            title='投資報酬率 (%)',
            overlaying='y',
            side='right'
        ),
        yaxis=dict(
            title='未實現損益估算'
        ),
        title=''
    )

    st.plotly_chart(fig_profit)

    st.markdown("---")

    # --- Show the Selected Data ---
    st.markdown("Selected Stock List")
                
    output_columns = ['股票代碼', '股票名稱', '股票中文名稱', '公司名', '成交股數', '實際總成本','平均持股單價', '股價', '市值估算', '未實現損益估算']

    df_display = df_selected[output_columns]

    df_display_styled = df_display.style.format({
        '成交股數': '{0:,.0f}',
        '實際總成本': 'JP¥{0:,.0f}',
        '平均持股單價': '{0:,.1f}',
        '股價': '{0:,.1f}',
        '市值估算': 'JP¥{0:,.0f}',
        '未實現損益估算': 'JP¥{0:,.0f}',
    })

    st.dataframe(df_display_styled)


# --- Hide Streamlit Style ---
hide_streamlit_style = """
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
    </style>
    """

st.markdown(hide_streamlit_style, unsafe_allow_html=True)
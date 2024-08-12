import pickle
from pathlib import Path
import streamlit_authenticator as stauth
import streamlit as st
import numpy as np
import pandas as pd
import plotly.express as px
from PIL import Image
from streamlit_option_menu import option_menu

st.set_page_config(
    page_title="Dashboard",
    page_icon=":bar_chart:",
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
    # --- Load the Data ---
    left_column, right_column = st.columns([2, 3])
    with left_column:
        image = Image.open('image/Financial planner.jpg')
        st.image(image, use_column_width=True)
    with right_column:
        st.header("Explore Your Data 📈")
        uploaded_file = st.file_uploader("", type=["xlsx", "xls"])

    st.markdown("---")

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

    # --- Explore the Data ---
    if uploaded_file:
        df = get_data_from_excel(uploaded_file, "Database_Bonds")

        # --- Sidebar Welcome ---
        st.sidebar.title(f"Welcome {name}")

        # --- Sidebar Filters ---
        st.sidebar.header("Please Filter Here")

        company = st.sidebar.multiselect(
            label="Select Company",
            options=df["公司名"].unique(),
            default=df["公司名"].unique(),
        )

        bond_type = st.sidebar.multiselect(
            label="Select Type of Bond",
            options=df["債券類別"].unique(),
            default=df["債券類別"].unique(),
        )

        risk_level = st.sidebar.multiselect(
            label="Select Risk Level",
            options=df["信評等級"].unique(),
            default=df["信評等級"].unique(),
        )

        brokers = st.sidebar.multiselect(
            label="Brokers",
            options=df["交易對象"].unique(),
            default=df["交易對象"].unique(),
        )

        # --- Filter the Data ---
        df_selected = df.query(
            '公司名==@company & 債券類別==@bond_type & 信評等級==@risk_level & 交易對象==@brokers'
        )

        # --- Dataframe Group ---
        groupby_ytm = df_selected.groupby(
            "殖利率區間")["交割金額"].sum().reset_index()

        groupby_duration = df_selected.groupby(
            "存續期區間")["交割金額"].sum().reset_index()

        groupby_company = df_selected.groupby(
            "公司名")["交割金額"].sum().reset_index().sort_values(by="交割金額", ascending=False)

        groupby_type = df_selected.groupby(
            "債券類別")["交割金額"].sum().reset_index()

        groupby_risk_level = df_selected.groupby(
            "信評等級")["交割金額"].sum().reset_index().sort_values(by="交割金額", ascending=False)

        groupby_interest = df_selected.groupby(
            "公司名")["利息"].sum().reset_index().sort_values(by="利息", ascending=False)

        groupby_broker = df_selected.groupby(
            "交易對象")["交割金額"].sum().reset_index()

        groupby_issuer = df_selected.groupby(
            "發行機構")["交割金額"].sum().reset_index().sort_values(by="交割金額", ascending=False).head(10)

        # --- Show the KPIs ---
        st.title("📋Investment Overview")

        total_inevstment_amount = int(df_selected["交割金額"].sum())

        total_payback = int(df_selected["本利合計"].sum())

        total_return = total_payback - total_inevstment_amount

        average_duration = (
            sum(df_selected["交割金額"] * df_selected["存續期間"])/total_inevstment_amount)

        retrun_rate_yearly = total_return / total_inevstment_amount / average_duration

        left_column, right_column = st.columns(2)
        with left_column:
            st.subheader(f"Total Investment: US$ {total_inevstment_amount:,}")
        with right_column:
            st.subheader(f"Duration Period(Years): {average_duration:.1f}")

        left_column, right_column = st.columns(2)
        with left_column:
            st.subheader(f"Total Return: US$ {total_return:,}")
        with right_column:
            st.subheader(f"Average Yearly Return Rate: {retrun_rate_yearly:.1%}")

        st.markdown("---")


        # --- Visualize the Analysis ---
        st.title(":bar_chart: Dashboard")
        st.markdown("### Analysis of Investment")

        fig_company = px.bar(
            data_frame=groupby_company,
            orientation="h",
            x="交割金額",
            y="公司名",
            text_auto='$.2s',
            color="公司名",
            title="Investment Amount by Company",
            labels={"公司名": "Company", "交割金額": "Amount"},
        )

        fig_interest = px.bar(
            data_frame=groupby_interest,
            x="公司名",
            y="利息",
            text_auto='$.2s',
            color="公司名",
            title="Interest Amount by Company",
            labels={"公司名": "Company", "利息": "Interest Amount"},
        )

        left_column, right_column = st.columns(2)
        with left_column:
            st.plotly_chart(fig_company)
        with right_column:
            st.plotly_chart(fig_interest)


        fig_ytm = px.bar(
            data_frame=groupby_ytm,
            x="殖利率區間",
            y="交割金額",
            text_auto='$.2s',
            color="殖利率區間",
            title="Investment Amount by YTM",
            labels={"殖利率區間": "Yield to Maturity", "交割金額": "Amount"},
        )

        fig_duration = px.bar(
            data_frame=groupby_duration,
            x="存續期區間",
            y="交割金額",
            text_auto='$.2s',
            color="存續期區間",
            title="Investment Amount by Duration",
            labels={"存續期區間": "Duration(Years)", "交割金額": "Amount"},
        )


        left_column, right_column = st.columns(2)
        with left_column:
            st.plotly_chart(fig_ytm)
        with right_column:
            st.plotly_chart(fig_duration)

        fig_risk = px.bar(
            data_frame=groupby_risk_level,
            orientation="h",
            x="交割金額",
            y="信評等級",
            text_auto='$.2s',
            color="信評等級",
            title="Investment Amount by Risk Level",
            labels={"信評等級": "Risk Level", "交割金額": "Amount"},
        )

        fig_issuer = px.bar(
            data_frame=groupby_issuer,
            x="發行機構",
            y="交割金額",
            text_auto='$.2s',
            color="發行機構",
            title="Top 10 Investment Amount by Issuers",
            labels={"發行機構": "Issuers", "交割金額": "Amount"},
        )

        left_column, right_column = st.columns(2)
        with left_column:
            st.plotly_chart(fig_risk)
        with right_column:
            st.plotly_chart(fig_issuer)


        fig_type = px.pie(
            data_frame=groupby_type,
            names="債券類別",
            values="交割金額",
            title="Investment Amount by Type of Bond",
            hole=0.5,
        )

        fig_type.update_traces(
            textinfo='value+percent', 
            textposition='inside',
        )

        fig_broker = px.pie(
            data_frame=groupby_broker,
            names="交易對象",
            values="交割金額",
            title="Investment Amount by Type of Broker",
            hole=0.5,
        )  

        fig_broker.update_traces(
            textinfo='value+percent', 
            textposition='inside',
        )


        left_column, right_column = st.columns(2)
        with left_column:
            st.plotly_chart(fig_type)
        with right_column:
            st.plotly_chart(fig_broker)

        fig_tree = px.treemap(
            data_frame=df_selected,
            path=["公司名", "債券類別", "信評等級"],
            values="交割金額",
            color='信評等級',
            title="Investment Amount by Company, Type of Bond, and Risk Level",
        )

        fig_tree.update_layout(width=1200, height=600)

        st.plotly_chart(fig_tree, use_column_width=True)

        st.markdown("---")


        # --- Charts for Payback Date---
        st.markdown("### Payback Timeseries Analysis")

        df_payback = get_data_from_excel(uploaded_file, "Database_Payback")

        df_payback_selected = df_payback[
            (df_payback["公司名"].isin(company)) \
            & (df_payback["交易對象"].isin(brokers)) \
            & (df_payback["信評等級"].isin(risk_level)) \
            & (df_payback["債券類別"].isin(bond_type))
        ]

        startDay = pd.to_datetime(df_payback_selected['配息日']).min()
        endDay = pd.to_datetime(df_payback_selected['配息日']).max()

        left_column, right_column = st.columns([3,2])
        with left_column:
            selected = option_menu(
                    menu_title=None,
                    options=["Yearly", "Monthly"],
                    orientation="horizontal",
                )

        left_column, right_column = st.columns([3,2])
        with left_column:
            col_1, col_2 = st.columns(2)
            with col_1:
                date1 = pd.to_datetime(st.date_input("Start Date", startDay))
            with col_2:
                date2 = pd.to_datetime(st.date_input("End Date", endDay))

        df_payback_selected = df_payback_selected[
            (df_payback_selected['配息日'] >= date1) \
            & (df_payback_selected['配息日'] <= date2)
        ].copy()

        if selected == "Yearly":
                df_payback_selected['Year'] = df_payback_selected['配息日'].dt.to_period('Y')

                groupby_interest = pd.DataFrame(df_payback_selected.groupby(
                    df_payback_selected['Year'].dt.strftime('%Y'))["應收利息"].sum()).reset_index()

                groupby_payack = pd.DataFrame(df_payback_selected.groupby(
                    df_payback_selected['Year'].dt.strftime('%Y'))["本利合計"].sum()).reset_index()


                fig_interest = px.bar(
                    data_frame=groupby_interest,
                    x="應收利息",
                    y="Year",
                    orientation="h",
                    text_auto='$.2s',
                    color='Year',
                    title="Interest Payment Schedule",
                    labels={"應收利息": "Total Interest", "Year": "Date"},
                    height=600,
                )

                fig_payback = px.bar(
                    data_frame=groupby_payack,
                    x="本利合計",
                    y="Year",
                    orientation="h",
                    text_auto='$.2s',
                    color='Year',
                    title="Total Payback Schedule",
                    labels={"本利合計": "Total Payback", "Year": "Date"},
                    height=600,
                )

                left_column, right_column = st.columns(2)
                with left_column:
                    st.plotly_chart(fig_interest)
                with right_column:
                    st.plotly_chart(fig_payback)


        if selected == "Monthly":
            df_payback_selected['Year-Month'] = df_payback_selected['配息日'].dt.to_period('M')

            groupby_interest = pd.DataFrame(df_payback_selected.groupby(
                    df_payback_selected['Year-Month'].dt.strftime('%Y/%m'))["應收利息"].sum()).reset_index()

            groupby_payack = pd.DataFrame(df_payback_selected.groupby(
                df_payback_selected['Year-Month'].dt.strftime('%Y/%m'))["本利合計"].sum()).reset_index()

            fig_interest = px.bar(
                data_frame=groupby_interest,
                x="Year-Month",
                y="應收利息",
                text_auto='$.2s',
                color='Year-Month',
                title="Interest Payment Schedule",
                labels={"應收利息": "Total Interest", "Year-Month": "Date"},
            )

            fig_payback = px.bar(
                data_frame=groupby_payack,
                x="Year-Month",
                y="本利合計",
                text_auto='$.2s',
                color='Year-Month',
                title="Total Payback Schedule",
                labels={"本利合計": "Total Payback", "Year-Month": "Date"},
            )


            left_column, right_column = st.columns(2)
            with left_column:
                st.plotly_chart(fig_interest)
            with right_column:
                st.plotly_chart(fig_payback)


        st.markdown("---")

        # --- Show the Selected Data ---
        st.markdown("### Selected Bond List")
        
        output_columns = ['ISIN', '公司名', '債券類別', '交易對象', '發行機構', '信評(M/S/F)', '信評等級', '票面利率', '存續期間', '到期日', '買入價', '殖利率', '投資面額', '交割金額', '利息', '本利合計']

        df_display = df_selected[output_columns]

        df_display['到期日'] = pd.to_datetime(df_display['到期日']).dt.strftime("%Y/%m/%d")

        df_display_styled = df_display.style.format({
            "存續期間": '{:.1f}', 
            "買入價":  '${:.2f}', 
            "票面利率": '{:.2%}', 
            "殖利率":  '{:.2%}',
            "投資面額": '${0:,.0f}',
            "應付本金": '${0:,.0f}',
            "前手息": '${0:,.0f}',
            "利息": '${0:,.0f}',
            "本利合計": '${0:,.0f}',
            "交割金額": '${0:,.0f}',
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
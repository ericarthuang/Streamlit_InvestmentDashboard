import streamlit as st 

# --- Page Setup ---
about_page = st.Page(
    page="views/about_me.py",
    title="About This App",
    icon="👩‍💻",
    default=True,
)

# --- Logo Setup ---
st.logo('image/PSA Logo.jpg')


# --- Page Setup ---
viewSharePoint_page = st.Page(
    page="views/viewSharePoint.py",
    title="View PSA Bonds",
    icon=":material/bar_chart:",
)

viewByCompany_page = st.Page(
    page="views/viewCompany.py",
    title="Your Own Bonds",
    icon=":material/business:",
)

viewstock_page = st.Page(
    page="views/viewStock.py",
    title="View PSA Stocks",
    icon=":material/trending_up:",
)

# --- Navigation Setup ---
pg = st.navigation(
    {
        "Info": [about_page],
        "Projects": [viewSharePoint_page, viewByCompany_page, viewstock_page],
    }
)


# --- Run Navigation ---
pg.run()
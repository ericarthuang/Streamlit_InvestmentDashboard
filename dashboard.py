import streamlit as st 

# --- Page Setup ---
about_page = st.Page(
    page="views/about_me.py",
    title="About This App",
    icon="👩‍💻",
    default=True,
)

viewall_page = st.Page(
    page="views/viewAll.py",
    title="View PSA Bonds",
    icon=":material/bar_chart:",
)

viewbycompany_page = st.Page(
    page="views/viewByCompany.py",
    title="Your Own Bonds",
    icon=":material/business:",
)

# --- Navigation Setup ---
pg = st.navigation(
    {
        "Info": [about_page],
        "Projects": [viewall_page, viewbycompany_page],
    }
)

# --- Logo Setup ---
st.logo('image/PSA Logo.jpg')

# --- Run Navigation ---
pg.run()
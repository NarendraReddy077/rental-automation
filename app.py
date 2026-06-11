import streamlit as st
import importlib

# 1. Streamlit Page Configuration (Must run first!)
st.set_page_config(
    page_title="Siemens - Rental Automation System",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Imports and Hot Reloading for Development
import sheets.main_sheet as main_sheet
import sheets.rent_calculation as rent_calculation
import sheets.lease_size as lease_size
import sheets.lease_rent as lease_rent
import sheets.capex_pm as capex_pm
import sheets.security_deposit as security_deposit
import sheets.input_parser as input_parser

import core.parameters as parameters
import core.excel_compiler as excel_compiler
import ui.styles as styles
import ui.sidebar as sidebar
import ui.dashboard as dashboard

# Reload sheets components
importlib.reload(main_sheet)
importlib.reload(rent_calculation)
importlib.reload(lease_size)
importlib.reload(lease_rent)
importlib.reload(capex_pm)
importlib.reload(security_deposit)
importlib.reload(input_parser)

# Reload core and ui modules
importlib.reload(parameters)
importlib.reload(excel_compiler)
importlib.reload(styles)
importlib.reload(sidebar)
importlib.reload(dashboard)

# 3. Inject Stylesheets & Background
styles.inject_styles_and_bg()

# 4. Render Dynamic Header
header_html = styles.load_template("header.html")
if header_html:
    st.markdown(header_html, unsafe_allow_html=True)
else:
    st.title("🏢 Rental Automation System")

# 5. Render Sidebar & Collect UI Configuration
ui_params, upload_occurred, parse_error_msg = sidebar.render_sidebar()

# 6. Render Dashboard & Excel Compilation/Download Action
dashboard.render_dashboard(ui_params, upload_occurred, parse_error_msg)

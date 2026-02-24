import streamlit as st
import sys

# ==============================================================================
# --- BARE BONES CONNECTION TEST ---
# ==============================================================================

st.set_page_config(page_title="System Check", layout="centered")

st.title("🔧 System Diagnostic")

# Check if we can see the environment
st.write(f"Python Version: {sys.version}")

st.success("The Streamlit script 'cr_CG.py' has loaded successfully.")

st.info("If you are still seeing a 'ModuleNotFoundError' referring to a line that is a comment, please 'Reboot App' in the Streamlit Cloud dashboard to clear the cache.")

# Minimal Interactive Test
st.subheader("Interactivity Test")
val = st.slider("Move this slider to test reactivity", 0, 100, 50)
st.write(f"The slider value is: {val}")

if st.button("Run Balloon Test"):
    st.balloons()
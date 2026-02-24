import streamlit as st

# ==============================================================================
# --- BASE LEVEL CONNECTION TEST ---
# ==============================================================================

st.set_page_config(page_title="Streamlit Test", layout="centered")

st.title("🚀 Streamlit Connection Test")

st.success("If you can see this message, the Streamlit server is running correctly!")

st.markdown("""
### Environment Verification
This is a base-level script designed to bypass all external utility files and complex data processing. 
It confirms that:
1. The `cr_CG.py` entry point is being found.
2. Streamlit is able to render basic UI components.
3. The server is responsive to interactions.
""")

# Test interaction
name = st.text_input("Enter your name to test reactivity:", "User")
if name:
    st.write(f"Hello, **{name}**! The app is responding to your input.")

st.sidebar.title("Sidebar Test")
st.sidebar.info("Sidebar is working.")

if st.button("Click for a Balloon Test"):
    st.balloons()
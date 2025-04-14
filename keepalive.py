import streamlit as st
import time

def run_app():
    st.set_page_config(page_title="Keepalive Monitor", layout="centered")
    params = st.experimental_get_query_params()
    
    if "keepalive" in params:
        st.write("✅ Keepalive ping received at:", time.strftime("%Y-%m-%d %H:%M:%S"))
        st.stop()  # end early to avoid loading full UI
    
    st.title("🔁 Polytex Keepalive")
    st.markdown("This tool is used to keep Streamlit session alive via UptimeRobot.")
    st.markdown("You can ping this endpoint: `/keepalive?keepalive=true`")

if __name__ == "__main__":
    run_app()
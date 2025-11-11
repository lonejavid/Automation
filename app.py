"""
HME Cloud Automation - Web Interface
Simple web interface to start the automation process
"""

import streamlit as st
import subprocess
import sys
from pathlib import Path
import threading
import time

# Page config
st.set_page_config(
    page_title="HME Cloud Automation",
    page_icon="🤖",
    layout="centered"
)

# Custom CSS for better styling
st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stButton>button {
        width: 100%;
        height: 60px;
        font-size: 20px;
        font-weight: bold;
        background-color: #4CAF50;
        color: white;
        border-radius: 10px;
    }
    .stButton>button:hover {
        background-color: #45a049;
    }
    .status-box {
        padding: 20px;
        border-radius: 10px;
        margin: 20px 0;
        background-color: #f0f0f0;
    }
    </style>
""", unsafe_allow_html=True)

# Title and description
st.title("🤖 HME Cloud Automation")
st.markdown("---")
st.markdown("""
    **Welcome to the HME Cloud Automation System**
    
    This tool will automatically:
    - ✅ Login to HME Cloud
    - ✅ Download Raw Car Data Report
    - ✅ Format the Excel file using DT macro
    - ✅ Process the data
    
    Click the button below to start the automation process.
""")

# Initialize session state
if 'automation_running' not in st.session_state:
    st.session_state.automation_running = False
if 'automation_status' not in st.session_state:
    st.session_state.automation_status = "Ready to start"
if 'automation_output' not in st.session_state:
    st.session_state.automation_output = []

def run_automation():
    """Run the automation script in a subprocess"""
    try:
        st.session_state.automation_status = "🔄 Starting automation..."
        st.session_state.automation_output = []
        
        # Get the project root
        project_root = Path(__file__).parent
        script_path = project_root / "scripts" / "test_store_selection.py"
        
        # Run the automation script
        process = subprocess.Popen(
            [sys.executable, str(script_path)],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1,
            universal_newlines=True
        )
        
        # Read output line by line
        output_lines = []
        for line in process.stdout:
            output_lines.append(line.strip())
            st.session_state.automation_output = output_lines[-20:]  # Keep last 20 lines
        
        process.wait()
        
        if process.returncode == 0:
            st.session_state.automation_status = "✅ Automation completed successfully!"
        else:
            st.session_state.automation_status = "❌ Automation completed with errors"
            
    except Exception as e:
        st.session_state.automation_status = f"❌ Error: {str(e)}"
    finally:
        st.session_state.automation_running = False

def start_automation():
    """Start the automation in a background thread"""
    if not st.session_state.automation_running:
        st.session_state.automation_running = True
        thread = threading.Thread(target=run_automation, daemon=True)
        thread.start()

# Status display
status_container = st.container()
with status_container:
    st.markdown("### 📊 Status")
    status_placeholder = st.empty()
    
    # Display current status
    with status_placeholder.container():
        if st.session_state.automation_running:
            st.info(f"⏳ {st.session_state.automation_status}")
            with st.spinner("Automation is running... A browser window will open shortly."):
                pass
        else:
            if "✅" in st.session_state.automation_status:
                st.success(st.session_state.automation_status)
            elif "❌" in st.session_state.automation_status:
                st.error(st.session_state.automation_status)
            else:
                st.info(f"ℹ️ {st.session_state.automation_status}")

# Start button
st.markdown("---")
col1, col2, col3 = st.columns([1, 2, 1])
with col2:
    if st.button("🚀 Start Automation", disabled=st.session_state.automation_running):
        start_automation()
        st.rerun()

# Show output if available
if st.session_state.automation_output:
    st.markdown("---")
    with st.expander("📋 View Automation Log", expanded=False):
        st.code("\n".join(st.session_state.automation_output), language="text")

# Auto-refresh when automation is running
if st.session_state.automation_running:
    time.sleep(2)
    st.rerun()

# Footer
st.markdown("---")
st.markdown("""
    <div style='text-align: center; color: #666; padding: 20px;'>
        <p>HME Cloud Automation System</p>
        <p>Click "Start Automation" to begin the process</p>
    </div>
""", unsafe_allow_html=True)


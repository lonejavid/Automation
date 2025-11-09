"""
KFC Guyana Drive-Thru Automation Platform - INTEGRATED VERSION
Professional automation tool with real HMECloud integration
"""

import streamlit as st
from datetime import datetime, timedelta
import time
import threading
from pathlib import Path
import sys

PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_DIR = PROJECT_ROOT / "src"
if str(SRC_DIR) not in sys.path:
    sys.path.append(str(SRC_DIR))

from automation.hmecloud import download_all_stores, download_single_store
from automation.complete_automation import main as run_complete_automation

# Page configuration
st.set_page_config(
    page_title="KFC Guyana Drive-Thru Automation",
    page_icon="🍗",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for professional styling
st.markdown("""
    <style>
    .main {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    .hero-container {
        background: white;
        padding: 3rem;
        border-radius: 20px;
        box-shadow: 0 20px 60px rgba(0,0,0,0.3);
        margin: 2rem auto;
        max-width: 800px;
        text-align: center;
    }
    .hero-title {
        font-size: 3rem;
        font-weight: 800;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 1rem;
    }
    .hero-subtitle {
        font-size: 1.3rem;
        color: #666;
        margin-bottom: 2rem;
    }
    .feature-box {
        background: #f8f9fa;
        padding: 1.5rem;
        border-radius: 10px;
        margin: 1rem 0;
        border-left: 4px solid #667eea;
    }
    .step-number {
        display: inline-block;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        width: 40px;
        height: 40px;
        border-radius: 50%;
        line-height: 40px;
        text-align: center;
        font-weight: bold;
        margin-right: 10px;
    }
    .success-box {
        background: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
        padding: 1rem;
        border-radius: 10px;
        margin: 1rem 0;
    }
    .info-box {
        background: #d1ecf1;
        border: 1px solid #bee5eb;
        color: #0c5460;
        padding: 1rem;
        border-radius: 10px;
        margin: 1rem 0;
    }
    .warning-box {
        background: #fff3cd;
        border: 1px solid #ffeeba;
        color: #856404;
        padding: 1rem;
        border-radius: 10px;
        margin: 1rem 0;
    }
    .stButton>button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 0.75rem 2rem;
        font-size: 1.1rem;
        border-radius: 50px;
        font-weight: 600;
        transition: all 0.3s;
        width: 100%;
    }
    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 30px rgba(102, 126, 234, 0.4);
    }
    .credentials-form {
        background: white;
        padding: 2rem;
        border-radius: 15px;
        box-shadow: 0 10px 40px rgba(0,0,0,0.1);
        margin: 2rem auto;
        max-width: 600px;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state
if 'page' not in st.session_state:
    st.session_state.page = 'home'
if 'automation_complete' not in st.session_state:
    st.session_state.automation_complete = False

# Store list
STORES = [
    "(Ungrouped) 2 Vlissengen Road – KFC",
    "(Ungrouped) 5 Mandela - KFC",
    "(Ungrouped) Movie Towne - KFC",
    "(Ungrouped) Giftland Mall - KFC",
    "(Ungrouped) Sheriff Street - KFC",
    "(Ungrouped) Providence - KFC"
]

def show_home_page():
    """Landing page with branding and introduction"""
    
    st.markdown("""
        <div class="hero-container">
            <div style="font-size: 5rem; margin-bottom: 1rem;">🍗</div>
            <h1 class="hero-title">KFC Guyana</h1>
            <h2 class="hero-subtitle">Drive-Thru Optimization Platform</h2>
            <p style="color: #888; font-size: 1.1rem;">
                Fully automated data processing from HMECloud to final reports
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    # Features section
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
            <div class="feature-box">
                <h3 style="color: #667eea;">⚡ Fast</h3>
                <p>Complete automation in minutes</p>
            </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
            <div class="feature-box">
                <h3 style="color: #667eea;">🤖 Automated</h3>
                <p>Login & download automatically</p>
            </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
            <div class="feature-box">
                <h3 style="color: #667eea;">📊 Complete</h3>
                <p>Full end-to-end workflow</p>
            </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # What it does section
    st.markdown("""
        <div style="background: white; padding: 2rem; border-radius: 15px; margin: 2rem auto; max-width: 800px;">
            <h2 style="text-align: center; color: #333; margin-bottom: 2rem;">What This Tool Does</h2>
            
            <div style="margin: 1.5rem 0;">
                <span class="step-number">1</span>
                <strong>Logs in</strong> to HMECloud automatically
            </div>
            
            <div style="margin: 1.5rem 0;">
                <span class="step-number">2</span>
                <strong>Downloads</strong> Raw Car Data for selected stores
            </div>
            
            <div style="margin: 1.5rem 0;">
                <span class="step-number">3</span>
                <strong>Transforms</strong> data (unlimited rows!)
            </div>
            
            <div style="margin: 1.5rem 0;">
                <span class="step-number">4</span>
                <strong>Updates</strong> Drive-Thru Optimization template
            </div>
            
            <div style="margin: 1.5rem 0;">
                <span class="step-number">5</span>
                <strong>Refreshes</strong> pivot tables and dates
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    # Get started button
    st.markdown("<br>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🚀 Start Automation", key="get_started"):
            st.session_state.page = 'setup'
            st.rerun()

def show_setup_page():
    """Setup page for automation options"""
    
    # Back button
    if st.button("← Back to Home"):
        st.session_state.page = 'home'
        st.rerun()
    
    st.markdown("""
        <div class="hero-container">
            <div style="font-size: 3rem; margin-bottom: 1rem;">⚙️</div>
            <h1 class="hero-title">Automation Setup</h1>
            <p style="color: #888; font-size: 1.1rem;">
                Configure your automation preferences
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    # Setup form
    with st.container():
        st.markdown('<div class="credentials-form">', unsafe_allow_html=True)
        
        st.markdown("### 🏪 Store Selection")
        
        download_option = st.radio(
            "What would you like to download?",
            ["All Stores (Recommended)", "Single Store Only"],
            horizontal=True
        )
        
        if download_option == "Single Store Only":
            selected_store = st.selectbox(
                "Select Store",
                STORES
            )
        else:
            selected_store = None
        
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown("### 📅 Date Selection")
        
        date_option = st.radio(
            "Select Date",
            ["Yesterday (Default)", "Custom Date"],
            horizontal=True
        )
        
        if date_option == "Custom Date":
            selected_date = st.date_input(
                "Choose Date",
                value=datetime.now() - timedelta(days=1),
                max_value=datetime.now()
            )
        else:
            selected_date = datetime.now() - timedelta(days=1)
            st.info(f"📅 Will download data for: **{selected_date.strftime('%B %d, %Y')}**")
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Credentials info
        st.markdown("""
            <div class="info-box">
                <strong>🔒 Auto-Login Enabled:</strong><br>
                Using saved HMECloud credentials (doudit@hotmail.com)
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Start button
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("✨ Start Full Automation", key="start_automation"):
                # Save settings to session state
                st.session_state.download_all = (download_option == "All Stores (Recommended)")
                st.session_state.selected_store = selected_store
                st.session_state.selected_date = selected_date
                st.session_state.page = 'processing'
                st.session_state.automation_complete = False
                st.rerun()
        
        st.markdown('</div>', unsafe_allow_html=True)

def show_processing_page():
    """Processing/automation page with real execution"""
    
    st.markdown("""
        <div class="hero-container">
            <div style="font-size: 3rem; margin-bottom: 1rem;">⚙️</div>
            <h1 class="hero-title">Processing Automation</h1>
            <p style="color: #888; font-size: 1.1rem;">
                Running automated workflow...
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    # Show settings
    st.markdown(f"""
        <div class="info-box">
            <strong>📋 Automation Settings:</strong><br>
            <strong>Stores:</strong> {'All Stores' if st.session_state.download_all else st.session_state.selected_store}<br>
            <strong>Date:</strong> {st.session_state.selected_date.strftime('%B %d, %Y')}
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    if not st.session_state.automation_complete:
        # Show progress
        with st.spinner("Running automation... This may take a few minutes."):
            
            # Create status placeholders
            status_container = st.empty()
            
            try:
                # Step 1: Download from HMECloud
                status_container.markdown("""
                    <div class="info-box">
                        <strong>🌐 Logging into HMECloud...</strong><br>
                        Opening browser and authenticating...
                    </div>
                """, unsafe_allow_html=True)
                
                time.sleep(2)
                
                status_container.markdown("""
                    <div class="info-box">
                        <strong>📥 Downloading reports...</strong><br>
                        This will take a few minutes. Browser window will open.
                    </div>
                """, unsafe_allow_html=True)
                
                # Run the actual download
                if st.session_state.download_all:
                    download_success = download_all_stores(
                        report_date=st.session_state.selected_date
                    )
                else:
                    download_success = download_single_store(
                        st.session_state.selected_store,
                        report_date=st.session_state.selected_date
                    )
                
                if not download_success:
                    st.error("⚠️ Download encountered issues. Check browser window.")
                    st.stop()
                
                status_container.markdown("""
                    <div class="success-box">
                        <strong>✅ Download complete!</strong>
                    </div>
                """, unsafe_allow_html=True)
                
                time.sleep(1)
                
                # Step 2: Process data
                status_container.markdown("""
                    <div class="info-box">
                        <strong>🔄 Processing data...</strong><br>
                        Transforming and updating template...
                    </div>
                """, unsafe_allow_html=True)
                
                # Run the processing
                processing_success = run_complete_automation()
                
                if not processing_success:
                    st.error("❌ Data processing failed!")
                    st.stop()
                
                status_container.markdown("""
                    <div class="success-box">
                        <strong>✅ Processing complete!</strong>
                    </div>
                """, unsafe_allow_html=True)
                
                # Mark as complete
                st.session_state.automation_complete = True
                st.rerun()
                
            except Exception as e:
                st.error(f"❌ Error during automation: {e}")
                st.stop()
    
    else:
        # Show success
        st.markdown("""
            <div class="success-box" style="padding: 2rem; text-align: center;">
                <h2>🎉 Automation Complete!</h2>
                <p style="font-size: 1.2rem; margin: 1rem 0;">
                    Your Drive-Thru Optimization report has been updated successfully.
                </p>
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Next steps
        st.markdown("""
            <div class="info-box">
                <strong>📋 Next Steps:</strong><br>
                1. Open the template in Excel<br>
                2. Click "Refresh All" to update pivot tables<br>
                3. Verify data looks correct
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Action buttons
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("🔄 Process Another Report"):
                st.session_state.page = 'setup'
                st.session_state.automation_complete = False
                st.rerun()
        
        with col2:
            if st.button("🏠 Back to Home"):
                st.session_state.page = 'home'
                st.session_state.automation_complete = False
                st.rerun()

# Main app logic
def main():
    if st.session_state.page == 'home':
        show_home_page()
    elif st.session_state.page == 'setup':
        show_setup_page()
    elif st.session_state.page == 'processing':
        show_processing_page()

if __name__ == "__main__":
    main()


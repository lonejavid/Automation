"""
KFC Guyana Drive-Thru Automation Platform
Professional automation tool for HMECloud data processing
"""

import streamlit as st
from datetime import datetime, timedelta
import time

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
if 'credentials_saved' not in st.session_state:
    st.session_state.credentials_saved = False
if 'processing' not in st.session_state:
    st.session_state.processing = False

def show_home_page():
    """Landing page with branding and introduction"""
    
    st.markdown("""
        <div class="hero-container">
            <div style="font-size: 5rem; margin-bottom: 1rem;">🍗</div>
            <h1 class="hero-title">KFC Guyana</h1>
            <h2 class="hero-subtitle">Drive-Thru Optimization Platform</h2>
            <p style="color: #888; font-size: 1.1rem;">
                Automated data processing and reporting for drive-thru performance analytics
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    # Features section
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
            <div class="feature-box">
                <h3 style="color: #667eea;">⚡ Fast</h3>
                <p>Process data in seconds, not minutes</p>
            </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
            <div class="feature-box">
                <h3 style="color: #667eea;">🎯 Accurate</h3>
                <p>Eliminates manual errors</p>
            </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
            <div class="feature-box">
                <h3 style="color: #667eea;">📊 Complete</h3>
                <p>Full automation from download to reports</p>
            </div>
        """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # What it does section
    st.markdown("""
        <div style="background: white; padding: 2rem; border-radius: 15px; margin: 2rem auto; max-width: 800px;">
            <h2 style="text-align: center; color: #333; margin-bottom: 2rem;">What This Tool Does</h2>
            
            <div style="margin: 1.5rem 0;">
                <span class="step-number">1</span>
                <strong>Downloads</strong> Raw Car Data from HMECloud
            </div>
            
            <div style="margin: 1.5rem 0;">
                <span class="step-number">2</span>
                <strong>Transforms</strong> data using automated macro (unlimited rows!)
            </div>
            
            <div style="margin: 1.5rem 0;">
                <span class="step-number">3</span>
                <strong>Pastes</strong> into Drive-Thru Optimization template
            </div>
            
            <div style="margin: 1.5rem 0;">
                <span class="step-number">4</span>
                <strong>Refreshes</strong> all pivot tables and analytics
            </div>
            
            <div style="margin: 1.5rem 0;">
                <span class="step-number">5</span>
                <strong>Updates</strong> dates across all report sheets
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    # Get started button
    st.markdown("<br>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🚀 Get Started", key="get_started"):
            st.session_state.page = 'credentials'
            st.rerun()

def show_credentials_page():
    """Credentials input page"""
    
    # Back button
    if st.button("← Back to Home"):
        st.session_state.page = 'home'
        st.rerun()
    
    st.markdown("""
        <div class="hero-container">
            <div style="font-size: 3rem; margin-bottom: 1rem;">🔐</div>
            <h1 class="hero-title">HMECloud Login</h1>
            <p style="color: #888; font-size: 1.1rem;">
                Enter your HMECloud credentials to begin automation
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    # Credentials form
    with st.container():
        st.markdown('<div class="credentials-form">', unsafe_allow_html=True)
        
        st.markdown("### 📝 Login Credentials")
        
        username = st.text_input(
            "Username / Email",
            placeholder="Enter your HMECloud username",
            help="Your HMECloud login email"
        )
        
        password = st.text_input(
            "Password",
            type="password",
            placeholder="Enter your password",
            help="Your HMECloud password"
        )
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        st.markdown("### 🏪 Store Selection")
        
        store = st.selectbox(
            "Select Store",
            [
                "Select a store...",
                "(Ungrouped) 2 Vlissengen Road – KFC",
                "(Ungrouped) 5 Mandela - KFC",
                "(Ungrouped) Movie Towne - KFC",
                "(Ungrouped) Giftland Mall - KFC",
                "(Ungrouped) Sheriff Street - KFC",
                "(Ungrouped) Providence - KFC"
            ],
            help="Choose which store's data to download"
        )
        
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
        
        # Security notice
        st.markdown("""
            <div class="info-box">
                <strong>🔒 Security Notice:</strong><br>
                Your credentials are used only for this session and are not stored permanently.
            </div>
        """, unsafe_allow_html=True)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # Submit button
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("✨ Start Automation", key="start_automation"):
                if not username or not password:
                    st.error("⚠️ Please enter both username and password")
                elif store == "Select a store...":
                    st.error("⚠️ Please select a store")
                else:
                    # Save credentials to session state
                    st.session_state.username = username
                    st.session_state.password = password
                    st.session_state.store = store
                    st.session_state.date = selected_date
                    st.session_state.credentials_saved = True
                    st.session_state.page = 'processing'
                    st.rerun()
        
        st.markdown('</div>', unsafe_allow_html=True)

def show_processing_page():
    """Processing/automation page"""
    
    st.markdown("""
        <div class="hero-container">
            <div style="font-size: 3rem; margin-bottom: 1rem;">⚙️</div>
            <h1 class="hero-title">Processing Automation</h1>
            <p style="color: #888; font-size: 1.1rem;">
                Automating your drive-thru data workflow
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    # Show selected info
    st.markdown(f"""
        <div class="info-box">
            <strong>📋 Automation Details:</strong><br>
            <strong>Store:</strong> {st.session_state.store}<br>
            <strong>Date:</strong> {st.session_state.date.strftime('%B %d, %Y')}<br>
            <strong>Username:</strong> {st.session_state.username}
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Progress steps
    steps = [
        ("🌐 Connecting to HMECloud", "Logging in with your credentials..."),
        ("📥 Downloading Raw Car Data", "Fetching data from HMECloud servers..."),
        ("🔄 Transforming Data", "Running automated macro (all rows)..."),
        ("📋 Pasting into Template", "Updating Drive-Thru Optimization report..."),
        ("📊 Refreshing Pivot Tables", "Updating all analytics and charts..."),
        ("📅 Updating Dates", "Setting dates across all sheets..."),
        ("✅ Finalizing", "Saving and preparing your report...")
    ]
    
    # Create progress container
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # Simulate processing (replace with actual automation later)
    for i, (step_title, step_desc) in enumerate(steps):
        progress_bar.progress((i + 1) / len(steps))
        
        st.markdown(f"""
            <div class="success-box">
                <strong>{step_title}</strong><br>
                {step_desc}
            </div>
        """, unsafe_allow_html=True)
        
        time.sleep(1)  # Simulate work
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Success message
    st.markdown("""
        <div class="success-box" style="padding: 2rem; text-align: center;">
            <h2>🎉 Automation Complete!</h2>
            <p style="font-size: 1.2rem; margin: 1rem 0;">
                Your Drive-Thru Optimization report has been updated successfully.
            </p>
        </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Download/Next steps
    col1, col2 = st.columns(2)
    
    with col1:
        st.download_button(
            label="📥 Download Updated Report",
            data="Placeholder - Report will be here",
            file_name=f"Drive_Thru_Report_{st.session_state.date.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with col2:
        if st.button("🔄 Process Another Store"):
            st.session_state.page = 'credentials'
            st.rerun()
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    if st.button("🏠 Back to Home"):
        st.session_state.page = 'home'
        st.session_state.credentials_saved = False
        st.rerun()

# Main app logic
def main():
    if st.session_state.page == 'home':
        show_home_page()
    elif st.session_state.page == 'credentials':
        show_credentials_page()
    elif st.session_state.page == 'processing':
        show_processing_page()

if __name__ == "__main__":
    main()


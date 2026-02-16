"""
Michael Hill SKU Checker - Professional Streamlit Interface
Clean, lean, and aesthetic design
"""

import streamlit as st
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
import time
import io
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="Michael Hill SKU Checker",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS for professional, lean UI
st.markdown("""
    <style>
    /* Main container */
    .main {
        padding: 1rem 2rem;
        max-width: 1200px;
        margin: 0 auto;
    }
    
    /* Background */
    .stApp {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    }
    
    /* Remove default Streamlit padding */
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    
    /* Header styling */
    h1 {
        color: white !important;
        font-weight: 700;
        text-align: center;
        margin-bottom: 0.3rem !important;
        font-size: 2.5rem !important;
    }
    
    .subtitle {
        text-align: center;
        color: rgba(255, 255, 255, 0.9);
        font-size: 1rem;
        margin-bottom: 2rem;
    }
    
    /* Card styling - lean and clean */
    .upload-card {
        background: rgba(255, 255, 255, 0.95);
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin-bottom: 1rem;
        backdrop-filter: blur(10px);
    }
    
    /* File uploader */
    .stFileUploader {
        background: transparent;
    }
    
    .stFileUploader > div {
        background: white;
        border: 2px dashed #667eea;
        border-radius: 8px;
        padding: 1rem;
    }
    
    /* Buttons */
    .stButton > button {
        width: 100%;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-weight: 600;
        padding: 0.6rem 1.5rem;
        border-radius: 8px;
        border: none;
        font-size: 1rem;
        transition: all 0.2s;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    
    .stButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 4px 8px rgba(102, 126, 234, 0.3);
    }
    
    /* Metrics */
    [data-testid="stMetricValue"] {
        font-size: 1.8rem;
        font-weight: 700;
    }
    
    [data-testid="stMetric"] {
        background: rgba(255, 255, 255, 0.95);
        padding: 0.8rem;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05);
    }
    
    /* Progress bar */
    .stProgress > div > div {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
    }
    
    /* Dataframe */
    .stDataFrame {
        background: white;
        border-radius: 8px;
        overflow: hidden;
    }
    
    /* Info boxes - compact */
    .info-compact {
        background: rgba(255, 255, 255, 0.95);
        border-left: 3px solid #667eea;
        padding: 0.8rem 1rem;
        border-radius: 6px;
        margin: 0.5rem 0;
        font-size: 0.9rem;
    }
    
    .success-compact {
        background: rgba(240, 253, 244, 0.95);
        border-left: 3px solid #10b981;
        padding: 0.8rem 1rem;
        border-radius: 6px;
        margin: 0.5rem 0;
        font-size: 0.9rem;
        color: #065f46;
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background: rgba(255, 255, 255, 0.1);
        border-radius: 6px;
        font-size: 0.9rem;
    }
    
    /* Slider */
    .stSlider {
        padding: 0.5rem 0;
    }
    
    /* Remove extra spacing */
    .element-container {
        margin-bottom: 0.5rem;
    }
    
    /* Compact download button */
    .stDownloadButton > button {
        background: #10b981;
        color: white;
        font-weight: 600;
        border-radius: 8px;
        padding: 0.6rem 1.5rem;
        border: none;
    }
    
    .stDownloadButton > button:hover {
        background: #059669;
        transform: translateY(-1px);
    }
    
    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* Compact section headers */
    h3 {
        color: white !important;
        font-size: 1.3rem !important;
        margin-top: 1rem !important;
        margin-bottom: 0.5rem !important;
        font-weight: 600 !important;
    }
    </style>
""", unsafe_allow_html=True)

def setup_driver():
    """Setup Chrome driver for Selenium"""
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    
    try:
        driver = webdriver.Chrome(options=chrome_options)
        return driver
    except Exception as e:
        st.error(f"❌ Could not start Chrome driver: {str(e)}")
        st.info("📋 Make sure ChromeDriver is installed. See instructions in the sidebar.")
        return None

def check_sku_status(driver, sku):
    """Check if a SKU is listed or delisted"""
    try:
        search_url = f"https://www.michaelhill.com.au/search?q={sku}"
        driver.get(search_url)
        time.sleep(3)
        
        page_source = driver.page_source.lower()
        
        # Check for "No Search Results"
        try:
            no_results = driver.find_element(By.XPATH, "//*[contains(text(), 'No Search Results')]")
            if no_results:
                return "Delisted"
        except NoSuchElementException:
            pass
        
        # Check page text
        if "no products were found" in page_source or \
           "no search results for" in page_source:
            return "Delisted"
        
        # Look for products
        try:
            products = driver.find_elements(By.CSS_SELECTOR, 
                ".product-item, .product-card, .product-tile, a[href*='/product/']")
            if products and len(products) > 0:
                return "Listed"
        except:
            pass
        
        # Check for Add to Cart
        try:
            cart_buttons = driver.find_elements(By.XPATH, 
                "//*[contains(text(), 'Add to Cart') or contains(text(), 'Add to Bag')]")
            if cart_buttons and len(cart_buttons) > 0:
                return "Listed"
        except:
            pass
        
        # Check page title
        if "no search results" in driver.title.lower():
            return "Delisted"
        
        return "Delisted"
        
    except Exception as e:
        return "Error"

def create_sample_excel():
    """Create a sample Excel file for download"""
    sample_data = {
        'SKU': ['23360778', '23402560', '23189867', '22334633', '23360747']
    }
    df = pd.DataFrame(sample_data)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='SKUs')
    
    return output.getvalue()

# Main app
def main():
    # Compact header
    st.markdown("<h1>💎 Michael Hill SKU Checker</h1>", unsafe_allow_html=True)
    st.markdown("<p class='subtitle'>Check product status on Michael Hill Australia website</p>", unsafe_allow_html=True)
    
    # Two column layout for compact design
    col_left, col_right = st.columns([2, 1])
    
    with col_left:
        # File uploader - compact
        st.markdown("### 📤 Upload Excel File")
        uploaded_file = st.file_uploader(
            "SKUs should be in Column A",
            type=['xlsx', 'xls'],
            help="Upload Excel file with SKU numbers in first column",
            label_visibility="collapsed"
        )
    
    with col_right:
        # Quick stats or sample download
        st.markdown("### 📥 Sample")
        st.download_button(
            label="Download Sample",
            data=create_sample_excel(),
            file_name="sample_skus.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    if uploaded_file is not None:
        try:
            # Read Excel file
            df = pd.read_excel(uploaded_file)
            
            if df.empty:
                st.error("❌ The uploaded file is empty!")
                return
            
            # Get SKUs from first column
            skus = df.iloc[:, 0].dropna().astype(str).tolist()
            
            if not skus:
                st.error("❌ No SKUs found in Column A!")
                return
            
            # Compact success message
            st.markdown(f"<div class='success-compact'>✅ Found <b>{len(skus)} SKUs</b> ready to check</div>", unsafe_allow_html=True)
            
            # Compact settings row
            col1, col2, col3 = st.columns([1, 1, 1])
            
            with col1:
                with st.expander("👀 Preview SKUs"):
                    preview_df = pd.DataFrame({'SKU': skus[:10]})
                    st.dataframe(preview_df, use_container_width=True, height=200)
                    if len(skus) > 10:
                        st.caption(f"... and {len(skus) - 10} more")
            
            with col2:
                delay = st.slider(
                    "Delay (sec)",
                    min_value=1,
                    max_value=5,
                    value=2,
                    help="Delay between checks"
                )
            
            with col3:
                st.write("")  # Spacing
                st.write("")  # Spacing
                check_button = st.button("🚀 Start Checking", use_container_width=True, type="primary")
            
            # Start checking
            if check_button:
                results = []
                
                # Compact progress section
                st.markdown("---")
                progress_bar = st.progress(0)
                status_container = st.empty()
                
                # Compact metrics in one row
                metric_cols = st.columns(4)
                total_metric = metric_cols[0].empty()
                listed_metric = metric_cols[1].empty()
                delisted_metric = metric_cols[2].empty()
                error_metric = metric_cols[3].empty()
                
                # Initialize driver
                with status_container.container():
                    st.info("🌐 Starting browser...")
                
                driver = setup_driver()
                
                if driver is None:
                    st.error("❌ Failed to start browser. Check ChromeDriver installation.")
                    return
                
                try:
                    # Check each SKU
                    for i, sku in enumerate(skus):
                        with status_container.container():
                            st.info(f"🔍 Checking {i+1}/{len(skus)}: **{sku}**")
                        
                        status = check_sku_status(driver, sku)
                        results.append({'SKU': sku, 'Status': status})
                        
                        # Update progress
                        progress_bar.progress((i + 1) / len(skus))
                        
                        # Update metrics
                        results_df = pd.DataFrame(results)
                        listed_count = len(results_df[results_df['Status'] == 'Listed'])
                        delisted_count = len(results_df[results_df['Status'] == 'Delisted'])
                        error_count = len(results_df[results_df['Status'] == 'Error'])
                        
                        total_metric.metric("Total", i + 1)
                        listed_metric.metric("✅ Listed", listed_count)
                        delisted_metric.metric("❌ Delisted", delisted_count)
                        error_metric.metric("⚠️ Errors", error_count)
                        
                        if i < len(skus) - 1:
                            time.sleep(delay)
                    
                    status_container.success("🎉 Complete!")
                    
                finally:
                    driver.quit()
                
                # Results
                results_df = pd.DataFrame(results)
                
                st.markdown("---")
                st.markdown("### 📊 Results")
                
                # Compact results display
                def style_status(val):
                    colors = {
                        'Listed': 'background-color: #d1fae5; color: #065f46',
                        'Delisted': 'background-color: #fee2e2; color: #991b1b',
                        'Error': 'background-color: #fef3c7; color: #92400e'
                    }
                    return colors.get(val, '')
                
                styled_df = results_df.style.applymap(style_status, subset=['Status'])
                st.dataframe(styled_df, use_container_width=True, height=300)
                
                # Compact summary
                col1, col2, col3 = st.columns(3)
                col1.metric("✅ Listed", f"{listed_count} ({listed_count/len(results_df)*100:.0f}%)")
                col2.metric("❌ Delisted", f"{delisted_count} ({delisted_count/len(results_df)*100:.0f}%)")
                col3.metric("⚠️ Errors", f"{error_count} ({error_count/len(results_df)*100:.0f}%)")
                
                # Download
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    results_df.to_excel(writer, index=False, sheet_name='Results')
                
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                st.download_button(
                    label="📥 Download Results",
                    data=output.getvalue(),
                    file_name=f"sku_results_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
    
    else:
        # Instructions when no file
        st.markdown("<div class='info-compact'>", unsafe_allow_html=True)
        st.info("📋 **Quick Start:** Upload Excel file with SKUs in Column A → Click Start Checking → Download Results")
        st.markdown("</div>", unsafe_allow_html=True)
    
    # Compact footer
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("""
        <div style='text-align: center; color: rgba(255,255,255,0.7); padding: 0.5rem;'>
            <small>Powered by Selenium & Streamlit</small>
        </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
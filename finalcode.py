import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import os
import time
import streamlit as st
from io import BytesIO

# Setup WebDriver
chrome_options = Options()
chrome_options.add_argument("--headless")  # Run in headless mode
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# Function to search Google and extract LinkedIn profile link
def search_linkedin(first_name, last_name, credential_text, address, state):
    query = f"{first_name} {last_name} {credential_text} {address} {state} LinkedIn"
    driver.get(f"https://www.google.com/search?q={query}")
    
    try:
        # Wait for search results to load and locate LinkedIn profiles
        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'a')))
        links = driver.find_elements(By.CSS_SELECTOR, 'a')
        
        if links:  # Check if links is not empty
            for link in links:
                href = link.get_attribute('href')
                if href and 'linkedin.com/in/' in href:
                    return href
        return None
    except Exception as e:
        print(f"Error during search: {e}")
        return None

# Streamlit UI
st.title("LinkedIn Profile Finder")

# File uploader
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    # Load data from Excel
    try:
        df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"Error reading the Excel file: {e}")
        st.stop()
    
    # Add new columns to DataFrame
    df['linkedin_links'] = None
    df['Not found'] = None
    
    # Input for row range
    start_row = st.number_input("Start row index (0-based)", min_value=0, max_value=len(df)-1, value=0)
    end_row = st.number_input("End row index (exclusive)", min_value=start_row+1, max_value=len(df), value=len(df))
    
    # Validate the row range
    if start_row < 0 or end_row > len(df) or start_row >= end_row:
        st.error("Error: Invalid row range specified.")
        st.stop()
    
    # Process each row in the specified range of the DataFrame
    for index, row in df.iloc[start_row:end_row].iterrows():
        linkedin_link = search_linkedin(
            row.get('Provider First Name', ''),
            row.get('Provider Last Name (Legal Name)', ''),
            row.get('Provider Credential Text', ''),
            row.get('Provider First Line Business Practice Location Address', ''),
            row.get('Provider Business Practice Location Address State Name', '')
        )
        if linkedin_link:
            df.at[index, 'linkedin_links'] = linkedin_link
        else:
            df.at[index, 'Not found'] = 'No LinkedIn profile found'
        
        # Adding a delay to avoid hitting rate limits
        time.sleep(1)  # Adjust the sleep time as necessary
    
    # Save updated DataFrame to BytesIO
    output = BytesIO()
    try:
        df.to_excel(output, index=False)
        output.seek(0)
        st.success("Processing complete!")
        st.download_button(
            label="Download Updated Excel File",
            data=output,
            file_name="updated_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error saving the Excel file: {e}")

# Close WebDriver
driver.quit()

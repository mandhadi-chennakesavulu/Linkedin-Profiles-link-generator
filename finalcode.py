import streamlit as st
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# Function to search LinkedIn profile
def search_linkedin(first_name, last_name, state_name, state, location):
    query = f"{first_name} {last_name} {state_name} {state} {location} LinkedIn"
    driver.get(f"https://www.google.com/search?q={query}")
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'a')))
        links = driver.find_elements(By.CSS_SELECTOR, 'a')
        for link in links:
            href = link.get_attribute('href')
            if href and 'linkedin.com/in/' in href:
                return href
        return None
    except Exception as e:
        print(f"Error during search: {e}")
        return None

# Setup WebDriver
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

st.title("LinkedIn Profile Finder")
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    # Add new columns to DataFrame
    df['linkedin_links'] = None
    df['Not found'] = None

    # Process each row in the DataFrame
    for index, row in df.head(100).iterrows():
        linkedin_link = search_linkedin(
            row.get('Provider First Name', ''),
            row.get('Provider Last Name (Legal Name)', ''),
            row.get('Provider Business Mailing Address State Name', ''),
            row.get('State', ''),
            row.get('Location', '')
        )
        if linkedin_link:
            df.at[index, 'linkedin_links'] = linkedin_link
        else:
            df.at[index, 'Not found'] = 'No LinkedIn profile found'

    # Save updated DataFrame to new Excel file
    df.to_excel("output.xlsx", index=False)

    st.write(df)
    st.download_button(label="Download Updated Excel", data=open("output.xlsx", "rb"), file_name="output.xlsx")

# Close WebDriver
driver.quit()

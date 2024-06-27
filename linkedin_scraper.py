from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
from concurrent.futures import ThreadPoolExecutor

def google_search(query, index):
    driver = webdriver.Chrome()
    driver.get("https://www.google.com")
    search_box = driver.find_element(By.NAME, 'q')
    search_box.send_keys(query)
    search_box.send_keys(Keys.RETURN)
    time.sleep(5)  # Wait for the page to load

    links = driver.find_elements(By.XPATH, "//a[contains(@href, 'linkedin.com/in')]")
    linkedin_url = links[0].get_attribute('href') if links else None

    driver.quit()
    return index, linkedin_url

def main(excel_file_path, output_file_path):
    # Load the Excel file
    df = pd.read_excel(excel_file_path)

    # Prepare search queries
    search_queries = [row[0] + " " + row[1] + " CEO President chief" for _, row in df.iterrows()]

    # Perform searches with multithreading
    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = [executor.submit(google_search, query, index) for index, query in enumerate(search_queries)]

        for future in futures:
            index, linkedin_url = future.result()
            df.at[index, 'LinkedIn URL'] = linkedin_url

    # Save the updated DataFrame to an Excel file
    df.to_excel(output_file_path, index=False)
    print(f"Updated file saved to {output_file_path}")

if __name__ == "__main__":
    excel_file_path = "path/to/your/input_file.xlsx"  # Change this to the path of your input Excel file
    output_file_path = "path/to/your/output_file.xlsx"  # Change this to the path where you want to save the output Excel file

    main(excel_file_path, output_file_path)

""" Import Library """
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
# -------------------------------------------------------------------------#


""" Drivers """
# Set the path to the GeckoDriver executable
firefox_driver_path = "C:/Drivers/chromedriver_win32/chromedriver.exe"

# Create an instance of FirefoxOptions
firefox_options = Options()


# -------------------------------------------------------------------------#

""" Login Part """
# Pass the executable_path directly as an argument to the Firefox constructor
driver = webdriver.Chrome(
    #service=webdriver.Chrome.service.Service(executable_path=firefox_driver_path),
    #options=firefox_options,
)
# Open the desired URL
driver.get(
    "http://localhost/orangehrm-4.5/orangehrm-4.5/symfony/web/index.php/auth/login"
)

# Find elements and send keys
driver.find_element("id", "txtUsername").send_keys("Admin")
driver.find_element("id", "txtPassword").send_keys("Hoangvu#123")
driver.find_element("id", "btnLogin").click()

# -------------------------------------------------------------------------#


""" Edit contact details """
driver.get(
    "http://localhost/orangehrm-4.5/orangehrm-4.5/symfony/web/index.php/pim/contactDetails/empNumber/1"
)
# Turn on edit mode
driver.find_element("id", "btnSave").click()


""" Function edit """

results = []

import pandas as pd

def edit_info(
    Address_Street_01=None, Address_Street_02=None, city=None, State=None, Zip=None, 
    Home_Telephone=None, Mobile=None, Work_Phone=None, Work_Email=None, Other_Email=None
):
    driver.get(
    "http://localhost/orangehrm-4.5/orangehrm-4.5/symfony/web/index.php/pim/contactDetails/empNumber/1"
)
    # active button
    driver.find_element("id", "btnSave").click()  # turn on edit mode

    if pd.notna(Address_Street_01):
        driver.find_element("id", "contact_street1").clear()
        driver.find_element("id", "contact_street1").send_keys(Address_Street_01)

    if pd.notna(Address_Street_02):
        driver.find_element("id", "contact_street2").clear()
        driver.find_element("id", "contact_street2").send_keys(Address_Street_02)

    if pd.notna(city):
        driver.find_element("id", "contact_city").clear()
        driver.find_element("id", "contact_city").send_keys(city)

    if pd.notna(State):
        driver.find_element("id", "contact_province").clear()
        driver.find_element("id", "contact_province").send_keys(State)

    if pd.notna(Zip):
        driver.find_element("id", "contact_emp_zipcode").clear()
        driver.find_element("id", "contact_emp_zipcode").send_keys(Zip)

    if pd.notna(Home_Telephone):
        driver.find_element("id", "contact_emp_hm_telephone").clear()
        driver.find_element("id", "contact_emp_hm_telephone").send_keys(Home_Telephone)

    if pd.notna(Mobile):
        driver.find_element("id", "contact_emp_mobile").clear()
        driver.find_element("id", "contact_emp_mobile").send_keys(Mobile)

    if pd.notna(Work_Phone):
        driver.find_element("id", "contact_emp_work_telephone").clear()
        driver.find_element("id", "contact_emp_work_telephone").send_keys(Work_Phone)

    if pd.notna(Work_Email):
        driver.find_element("id", "contact_emp_work_email").clear()
        driver.find_element("id", "contact_emp_work_email").send_keys(Work_Email)

    if pd.notna(Other_Email):
        driver.find_element("id", "contact_emp_oth_email").clear()
        driver.find_element("id", "contact_emp_oth_email").send_keys(Other_Email)

    driver.find_element("id", "btnSave").click()  # save

    # Kiểm tra xem có thông báo lỗi xuất hiện hay không

    # Chờ cho thông báo thành công xuất hiện trong vòng 2 giây
    timeout = 0.1
    success_message_present = EC.presence_of_element_located(
        (By.CSS_SELECTOR, ".message.success.fadable")
    )
    try:
        success_message_element = WebDriverWait(driver, timeout).until(
            success_message_present
        )
        results.append("Pass")
    except:
        results.append("Fail")



# Read data from 'data.xlsx' file
data_file = "T01_Dataset.xlsx"

data_df = pd.read_excel(data_file)
filtered_df = data_df.drop(columns=["Expected"])

# Edit data
for _, row in filtered_df.iterrows():
    edit_info(**row)

data_df["Pass/Fail"] = results

# Compare value

# So sánh giá trị của hai cột và tạo cột mới 'Result'
data_df["Result"] = data_df.apply(
    lambda row: "Fail" if row["Expected"] != row["Pass/Fail"] else "Pass", axis=1
)

# Save the DataFrame to 'result.xlsx' file
result_file = "T01_result.xlsx"
data_df.to_excel(result_file, index=False)

# Close the WebDriver
driver.quit()

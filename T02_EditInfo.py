""" Import Library """
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
# -------------------------------------------------------------------------#


""" Drivers """
# Set the path to the GeckoDriver executable
firefox_driver_path = "C:\Drivers\geckodriver-v0.33.0-win64\geckodriver.exe"

# Create an instance of FirefoxOptions
firefox_options = webdriver.FirefoxOptions()


# -------------------------------------------------------------------------#

""" Login Part """
# Pass the executable_path directly as an argument to the Firefox constructor
driver = webdriver.Firefox(
    service=webdriver.firefox.service.Service(executable_path=firefox_driver_path),
    options=firefox_options,
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

""" Edit Info """
# Go to edit link profile
driver.get(
    "http://localhost/orangehrm-4.5/orangehrm-4.5/symfony/web/index.php/pim/addEmployee"
)

# Turn on edit mode
driver.find_element("id", "btnSave").click()

# Edit elements
""" Elements 
1. personal_txtEmpFirstName : text input first name . maxlength="30"
2. personal_txtEmpMiddleName :  text input middle name . maxlength="30"
3. personal_txtEmpLastName: text input last name . maxlength="30"
4. personal_txtEmployeeId: text input id employee . maxlength="10"
5. personal_txtLicenNo: text input id License Number . maxlength="30"
6. personal_txtOtherID: text input id other id number . maxlength="30" 
7. personal_txtLicExpDate: text input LicExpDate . format : yyyy-mm-dd 
8. personal_optGender_1 : radion male .
9. personal_optGender_2: radion female . 
10. personal_cmbMarital: list box Marital Status . 
11. personal_cmbNation: personal nationality 
12. personal_DOB: date of birth  

"""
""" Function edit """

results = []


def edit_info(
    firstName=None, middleName=None, lastName=None, EmployeeID=None, license_num=None, date_of_birth = None
):
    driver.get(
        "http://localhost/orangehrm-4.5/orangehrm-4.5/symfony/web/index.php/pim/viewMyDetails"
    )
    # active button
    driver.find_element("id", "btnSave").click()  # turn on edit mode
    driver.find_element("id", "personal_txtEmpFirstName").clear()
    driver.find_element("id", "personal_txtEmpMiddleName").clear()
    driver.find_element("id", "personal_txtEmpLastName").clear()
    driver.find_element("id", "personal_txtEmployeeId").clear()
    driver.find_element("id", "personal_txtLicenNo").clear()
    driver.find_element("id", "personal_DOB").clear()
    if pd.isna(firstName) == False:
        driver.find_element("id", "personal_txtEmpFirstName").send_keys(firstName)
    if pd.isna(middleName) == False:
        driver.find_element("id", "personal_txtEmpMiddleName").send_keys(middleName)
    if pd.isna(lastName) == False:
        driver.find_element("id", "personal_txtEmpLastName").send_keys(lastName)
    if pd.isna(EmployeeID) == False:
        driver.find_element("id", "personal_txtEmployeeId").send_keys(EmployeeID)
    if pd.isna(license_num) == False:
        driver.find_element("id", "personal_txtLicenNo").send_keys(license_num)
    if pd.isna(date_of_birth) == False:
        driver.find_element("id", "personal_DOB").send_keys(date_of_birth)
        driver.send_keys(Keys.RETURN)
    driver.find_element("id", "btnSave").click()  # save
    # Kiểm tra xem có thông báo lỗi xuất hiện hay không

    # Chờ cho thông báo thành công xuất hiện trong vòng 10 giây
    timeout = 1
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
data_file = "T02_Dataset.xlsx"

data_df = pd.read_excel(data_file)


# Edit data
for _, row in data_df.iterrows():
    first_name = row["First Name"]
    last_name = row["Last Name"]
    middle_name = row["Middle Name"]
    employee_id = row["Employee ID"]
    license_num = row["License Number"]
    edit_info(first_name, middle_name, last_name, employee_id, license_num)

data_df["Pass/Fail"] = results

# Compare value

# So sánh giá trị của hai cột và tạo cột mới 'Result'
data_df["Result"] = data_df.apply(
    lambda row: "Fail" if row["Expected"] != row["Pass/Fail"] else "Pass", axis=1
)

# Save the DataFrame to 'result.xlsx' file
result_file = "T02_result.xlsx"
data_df.to_excel(result_file, index=False)

# Close the WebDriver
driver.quit()

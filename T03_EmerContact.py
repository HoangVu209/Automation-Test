""" Import Library """ 
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

#-------------------------------------------------------------------------#


""" Drivers """  
# Set the path to the GeckoDriver executable
firefox_driver_path = "C:\Drivers\geckodriver-v0.33.0-win64\geckodriver.exe"

# Create an instance of FirefoxOptions
firefox_options = webdriver.FirefoxOptions()



#-------------------------------------------------------------------------#

""" Login Part """
# Pass the executable_path directly as an argument to the Firefox constructor
#driver = webdriver.Firefox(service=webdriver.firefox.service.Service(executable_path=firefox_driver_path), options=firefox_options)
driver = webdriver.Edge()
# Open the desired URL
driver.get("http://localhost/orangehrm-4.5/orangehrm-4.5/symfony/web/index.php/auth/login")

# Find elements and send keys
driver.find_element("id", "txtUsername").send_keys("Admin")
driver.find_element("id", "txtPassword").send_keys("Hoangvu#123")
driver.find_element("id", "btnLogin").click()


#-------------------------------------------------------------------------#

""" Add Emergency Contact """

results = [] #array to save result


# Function to add an emergency contact with given information
def add_contact(name = None, relationship = None, homeTelephone = None, mobile = None, workPhone = None):
    driver.get("http://localhost/orangehrm-4.5/orangehrm-4.5/symfony/web/index.php/pim/viewEmergencyContacts/empNumber/1") # Go to add emergency contact
    driver.find_element(By.ID, "btnAddContact").click()
    if pd.isna(name) == False : #not empty
        driver.find_element(By.ID, "emgcontacts_name").send_keys(name)
    if pd.isna(relationship) == False:
        driver.find_element(By.ID, "emgcontacts_relationship").send_keys(relationship)
    if(pd.isna(homeTelephone) == False):
        driver.find_element(By.ID, "emgcontacts_homePhone").send_keys(homeTelephone) 
    if(pd.isna(mobile) == False):
        driver.find_element(By.ID, "emgcontacts_mobilePhone").send_keys(mobile) 
    if(pd.isna(workPhone) == False):
        driver.find_element(By.ID, "emgcontacts_workPhone").send_keys(workPhone) 
    
    driver.find_element(By.ID, "btnSaveEContact").click() #Save

      # Chờ cho thông báo thành công xuất hiện trong vòng 10 giây
    timeout = 1
    success_message_present = EC.presence_of_element_located((By.CSS_SELECTOR, ".message.success.fadable"))
    try:
        success_message_element = WebDriverWait(driver, timeout).until(success_message_present)
        results.append("Pass")
    except:
        results.append("Fail")   

# Read data from 'data.xlsx' file
data_file = 'T03_Dataset.xlsx'
data_df = pd.read_excel(data_file)


#Call function 
# Edit data
for _, row in data_df.iterrows():
    name = row['Name']
    relation = row['Relation']
    telephone = row['Home Telephone']
    mobile = row['Mobile']
    work_phone = row['Work Telephone']
    add_contact(name, relation, telephone, mobile, work_phone)


data_df['Pass/Fail'] = results 

#Compare value 

# So sánh giá trị của hai cột và tạo cột mới 'Result'
data_df['Result'] = data_df.apply(lambda row: 'Fail' if row['Expected'] != row['Pass/Fail'] else 'Pass', axis=1)

# Save the DataFrame to 'result.xlsx' file
result_file = 'T03_result.xlsx'
data_df.to_excel(result_file, index=False)

# Close the WebDriver
driver.quit()








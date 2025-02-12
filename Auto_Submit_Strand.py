import os
import shutil
import fnmatch
import sys
import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains


name = os.getlogin()
bbuser = name + "@betenbough.com"
Main = f"C:/Users/{name}/OneDrive - Betenbough Homes/Desktop/Main"

def main():
    try:
        get_urls()
        login_window()
        get_strand_info()
        get_bid_set()
        fill_strand()

    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(f"\nError on line: {exc_tb.tb_lineno}")
        print(f"Error happened in getAgreementinfo(). This is the error - {e}")
        print("Currently trying to redo main().")
        input("Check out what happened.....")
        global termmain
        termmain += 1
        if termmain > 3:
            print(
                "Failed a total of 3 times, there is an issue with the script or with loading the website at this time."
                "Fix the problem or try later.")
            sys.exit()
        driver.quit()
        s(5)
        main()

    print("\nYour permit is ready to submit.\n")
    print("Make sure the city has the Master Set and the Rescheck for this plan!\n")
    print("Closing browser in 10 minutes.")
    s(600)
    print("Closed browser after 10 minutes.")
    driver.quit()


# Automatically download the correct ChromeDriver version

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
actions = ActionChains(driver)


def s(var):
    time.sleep(var)

def w(var):
    driver.implicitly_wait(var)

def get_urls():
    try:
        # Sets all the urls we're going to work with into a list.
        urls = ["https://homes.betenbough.com/Sales/Search.aspx",
                "https://apps.strandae.com/Security/LogOn?ReturnUrl=%2f"]

        # Opens the urls in order. Not 100% certain on HOW it works, but it does work.
        for posts in range(len(urls)):
            driver.get(urls[posts])
            if posts != len(urls) - 1:
                driver.execute_script("window.open('');")
                chwd = driver.window_handles
                driver.switch_to.window(chwd[-1])

        # Sets the windows to global variables to work with them in other functions.
        global agreements
        agreements = driver.window_handles[0]

        global strand
        strand = driver.window_handles[1]

    except Exception as e:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        print(f"\nError on line: {exc_tb.tb_lineno}")
        print(f"This is the error - {e}")

def login_window():
    # Function to handle login
    def login():
        global address_input
        address_input = address_entry.get()
        global password
        password = password_entry.get()
        root.quit()
        root.destroy()

    # Create the main window (popup)
    root = tk.Tk()
    root.geometry("300x150+700+400")
    root.title("Auto Submit Strand")

    # Create the username and password labels and entries
    address_label = tk.Label(root, text="Address:")
    address_label.grid(row=0, column=0, padx=10, pady=10)
    address_entry = tk.Entry(root)
    address_entry.grid(row=0, column=1, padx=10, pady=10)

    password_label = tk.Label(root, text="Password:")
    password_label.grid(row=1, column=0, padx=10, pady=10)
    password_entry = tk.Entry(root, show="*")  # Show password as stars
    password_entry.grid(row=1, column=1, padx=10, pady=10)

    # Create the login button
    login_button = tk.Button(root, text="Login", command=login)
    login_button.grid(row=2, column=1, padx=10, pady=10)

    # Set focus on the username entry when the popup opens
    address_entry.focus()

    # Start the main event loop
    root.mainloop()
    print(f"Looking for address: {address_input}")

def get_strand_info():
    try:
        driver.switch_to.window(agreements)

        s(2)
        driver.find_element(By.ID, "i0116").send_keys(bbuser)
        driver.find_element(By.ID, "idSIButton9").click()
        driver.implicitly_wait(10)
        driver.find_element(By.ID, "username").send_keys(bbuser + Keys.ENTER)
        driver.implicitly_wait(10)
        driver.find_element(By.ID, "password").send_keys(password + Keys.ENTER)
        input("Press Enter")

        '''
        # Login to BH Agreements
        driver.find_element(By.ID, "ContentPlaceHolder1_txtEmail").send_keys(bbuser)

        s(.25)
        driver.find_element(By.ID, "ContentPlaceHolder1_txtPassword").send_keys(password)

        s(.2)
        driver.find_element(By.ID, "ContentPlaceHolder1_txtPassword").send_keys(Keys.ENTER)
        '''

        # Search for the address
        w(5)
        driver.find_element(By.ID, "ctl00_MainContent_txtAddress").send_keys(address_input)
        s(.2)
        # Hits the search button.
        driver.find_element(By.ID, "ctl00_MainContent_txtAddress").send_keys(Keys.ENTER)

        w(5)
        driver.find_element(By.PARTIAL_LINK_TEXT, "View").click()

        w(10)

        global address
        address = driver.find_element(By.ID, "ctl00_MainContentWithoutPadding_lblAddress").text

        global handing
        handing = driver.find_element(By.ID, "ctl00_MainContentWithoutPadding_AgreementHeader_lblFloorPlan").text

        global floorplan
        floorplan = handing.split()[0]

        handing = handing.split()[1]

        handing = handing.strip("()")

        global upc
        upc = driver.find_element(By.ID, "ctl00_MainContentWithoutPadding_lblFloorPlanFamily").text

        global elevation
        elevation = driver.find_element(By.ID, "ctl00_MainContentWithoutPadding_lblFloorPlanUPC").text

        elevation = elevation.split("-")

        elevation = elevation[2][2:]

        global legalDes
        legal_desc = driver.find_element(By.ID, "ctl00_MainContentWithoutPadding_lblLegalDescription").text

        x = legal_desc.split(",")

        global lot_num
        lot_num = x[0][4:]

        global block_num
        block_num = x[1][7:]

        # SWITCH TO MANAGE HOMESITE TO GET PHASE.............
        s(1)
        driver.find_element(By.ID, "ctl00_MainContentWithoutPadding_AgreementHeader_lnkHomesiteDetails").click()

        # Get Phase info.
        w(5)
        global subdivision
        subdivision = driver.find_element(By.XPATH, "/html/body/form/div[4]/div[2]/div[2]/div[1]/table[1]/tbody/tr[2]/td[2]").text

        print("Address:  " + address)
        print("FloorPlan:  " + floorplan)
        print("UPC:  " + upc)
        print("Lot:  " + lot_num)
        print("Block:  " + block_num)
        print("Elevation:  " + elevation)
        print("Handing:  " + handing)
        print("Phase:  " + subdivision)

    except Exception as e:
        print(f"This is what happened: {e}")
        sys.exit()
        
def get_bid_set():
    def find_files_with_keyword(directory, keyword):
        # List to store the matching filenames
        matching_files = []

        # Loop through the files in the specified directory
        for root, dirs, files in os.walk(directory):
            for file in files:
                if keyword in file:  # Check if the keyword is in the filename
                    matching_files.append(os.path.join(root, file))  # Append full path

        return matching_files

    bid_set_name = find_files_with_keyword(r"V:\Architecture Design\01_PLANS\07_PERMITTING\POST-TENSION BID SETS", floorplan)
    bid_set_path = str(bid_set_name[0])

    def find_dwg_file_with_handing(folder_path, face):
        # Loop through the files in the directory
        for root, dirs, files in os.walk(folder_path):
            # Filter for files that match the pattern "*.dwg" and have 'FL' or 'FR' in the name
            for file in fnmatch.filter(files, f'*{face}*.dwg'):
                # Print or return the full path of the matching file
                file_path = os.path.join(root, file)
                shutil.copy(file_path, Main)
                return file_path  # Stop at the first found file

    folder_path = bid_set_path
    find_dwg_file_with_handing(folder_path)


def fill_strand():

    # Switch to Stand page.
    s(1)
    driver.switch_to.window(strand)
    driver.maximize_window()

    # Logs into Strand.
    s(1)
    driver.find_element(By.ID, "UserName").send_keys("permitting@betenbough.com")

    s(.2)
    driver.find_element(By.ID, "Password").send_keys("PB1113")

    s(.2)
    driver.find_element(By.XPATH, "/html/body/div[1]/div/div[2]/div[1]/div/div/div/div/form/div/div[2]/table/tbody/tr/td[2]/table/tbody/tr/td[1]/table/tbody/tr/td/table/tbody/tr[2]/td[2]/em/button").click()

    # 3002 Settler Ave
    # Red House
    s(1)
    driver.find_element(By.XPATH, "/html/body/div[1]/dl/dt[4]/a").click()

    # Order
    s(2)
    driver.find_element(By.XPATH, "/html/body/div[1]/div[3]/div[2]/div[1]/div/div/div/div/div/div[1]/div/table/tbody/tr/td[1]/table/tbody/tr/td[8]/table/tbody/tr[2]/td[2]/em/button").click()

    # Type Standard to Order Type
    s(1.5)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[1]/div[1]/div/input[2]").send_keys("s")
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[1]/div[1]/div/input[2]").send_keys(Keys.ENTER)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[1]/div[1]/div/input[2]").send_keys(Keys.TAB)


    # Type 'B' to enter BH - Amarillo
    s(1.5)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[2]/div[1]/div/input[2]").send_keys("B")
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[2]/div[1]/div/input[2]").send_keys(Keys.ENTER)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[2]/div[1]/div/input[2]").send_keys(Keys.TAB)

    # Click the subdivision dropdown.
    s(.5)

    if "Homestead 1" in subdivision:
        driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[3]/div[1]/div/div/div[1]/input[2]").send_keys("Ho")
        s(.25)
        driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[3]/div[1]/div/div/div[1]/input[2]").send_keys(Keys.ENTER)
    elif "Homestead 2" in subdivision:
        driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[3]/div[1]/div/div/div[1]/input[2]").send_keys("Ho")
        s(.25)
        driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[3]/div[1]/div/div/div[1]/input[2]").send_keys(Keys.ARROW_DOWN)
        driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[3]/div[1]/div/div/div[1]/input[2]").send_keys(Keys.ENTER)
    elif "Homestead 4" in subdivision:
        driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[3]/div[1]/div/div/div[1]/input[2]").send_keys("Ho")
        s(.25)
        driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[3]/div[1]/div/div/div[1]/input[2]").send_keys(Keys.ARROW_DOWN * 2)
        driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[3]/div[1]/div/div/div[1]/input[2]").send_keys(Keys.ENTER)
    elif "The Meadows 3" in subdivision:
        driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[3]/div[1]/div/div/div[1]/input[2]").send_keys("S")
        s(.25)
        driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[3]/div[1]/div/div/div[1]/input[2]").send_keys(Keys.ENTER)
    elif "The Meadows 4" in subdivision:
        driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[3]/div[1]/div/div/div[1]/input[2]").send_keys("S")
        s(.25)
        driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[3]/div[1]/div/div/div[1]/input[2]").send_keys(Keys.ARROW_DOWN * 2)
        driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[3]/div[1]/div/div/div[1]/input[2]").send_keys(Keys.ENTER)
    elif "Heritage Hills 18" in subdivision:
        driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[3]/div[1]/div/div/div[1]/input[2]").send_keys("H")
        s(.25)
        driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[3]/div[1]/div/div/div[1]/input[2]").send_keys(Keys.ENTER)


    s(2)
    actions.key_down(Keys.SHIFT).send_keys(Keys.TAB).key_up(Keys.SHIFT).perform()

    # Address
    s(2.5)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[5]/div[1]/div/div/input[1]").send_keys(address)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[5]/div[1]/div/div/input[1]").send_keys(Keys.TAB)

    # Lot
    s(.5)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[5]/div[1]/div/div/input[2]").send_keys(lot_num)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[5]/div[1]/div/div/input[2]").send_keys(Keys.TAB)

    # Block
    s(.5)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[5]/div[1]/div/div/input[3]").send_keys(block_num)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[5]/div[1]/div/div/input[3]").send_keys(Keys.TAB)

    # Plan
    s(.5)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[6]/div[1]/div/div/input[1]").send_keys(upc)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[6]/div[1]/div/div/input[1]").send_keys(Keys.TAB)

    # Elevation
    s(.5)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[6]/div[1]/div/div/input[2]").send_keys(elevation)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[6]/div[1]/div/div/input[2]").send_keys(Keys.TAB)

    # Swing
    s(.5)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[6]/div[1]/div/div/div[5]/input[2]").send_keys(handing)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[6]/div[1]/div/div/div[5]/input[2]").send_keys(Keys.ENTER)
    s(.5)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[6]/div[1]/div/div/div[5]/input[2]").send_keys(Keys.TAB * 9)

    # Comments                                  3112 Swenson St
    s(.5)
    driver.find_element(By.XPATH, "/html/body/div[14]/div[2]/div[1]/div/div/div/div/div/form/div[17]/div[1]/textarea").send_keys("PT Foundation Design")
    input("All Done.  Just upload your bid set and you're good to go!!! :)")

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    get_urls()
    login_window()
    get_strand_info()
    # get_bid_set()
    fill_strand()

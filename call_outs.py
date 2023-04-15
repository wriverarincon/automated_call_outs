from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import time
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime
import os

email = os.environ.get("dashboard_email")
password = os.environ.get("dashboard_password")
site = os.environ.get("radnet_dashboard")

# set chrome and page we want to open
opt = Options()
opt.add_argument("--lang=en") # change language
opt.add_experimental_option("detach", True)
driver = webdriver.Chrome(executable_path = r"\AutomatedTasks\chrome_data\chromedriver.exe", chrome_options=opt)
driver.get(site)
driver.maximize_window()

wait_time = WebDriverWait(driver, 10) # Create a variable with 10 seconds wait time for easier calling

# log in into the page
print("\nLogging into the system..")
try:
    email_bar = wait_time.until(
        EC.presence_of_element_located((By.ID, "IDToken1")) # Wait 10 seconds before it continues or until said ID appears
    )

    email_bar.clear() # Clear the bar
    email_bar.send_keys(email) # Insert the email
    email_bar.send_keys(Keys.RETURN) # Click enter

    password_bar = wait_time.until(
        EC.presence_of_element_located((By.ID, "IDToken2")) # Wait 10 seconds before it continues or until said ID appears
    )

    password_bar.clear() # Clear the bar
    password_bar.send_keys(password) # Insert the password
    password_bar.send_keys(Keys.RETURN) # Click enter

except Exception as e:
    print("e")
print("\nLogged in succesfully..")

time.sleep(3)

# select the dashboard and its content
print("\nEntering the dashboard..")
status = "NotDone"
while status == "NotDone":
    try:
        switch_to_frame = wait_time.until(
            EC.frame_to_be_available_and_switch_to_it((By.XPATH, "/html/body/div[1]/div[1]/div/section/div/div/div/div/div[2]/div/div/iframe"))
        )
        # click dropdown menu
        click_dropdown = driver.find_element(By.ID, "select2-dashboardType-container")
        click_dropdown.click()

        time.sleep(3)

        # select dashboard
        click_dashboard = driver.find_element(By.XPATH, "/html/body/span/span/span[2]/ul/li[4]")
        click_dashboard.click()
        
        status = "Done"
        time.sleep(10)

    except Exception as e:
        print("There was an error, trying again..")
        status = "NotDone"
print("\nDone..")

time.sleep(3)

switch_to__second_frame = wait_time.until(
    EC.frame_to_be_available_and_switch_to_it((By.XPATH, "/html/body/div[2]/section/div[4]/div/div/div[5]/iframe[1]"))
    )

time.sleep(3)

# filter teams
print("\nFiltering teams..")
check_all_good = False #check all the elements are available so the browser does not crash
while check_all_good == False:
    try:
        teams_search_bar = driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/form/div[2]/span/span[1]/span/ul/li[2]/input")
        teams_search_bar.send_keys("CCD")
        number_of_teams_found = driver.find_elements(By.CLASS_NAME, "select2-results__option")

        ccd_teams = ["CCD_Arizona", "CCD_Arizona_MR-CT-FL", "CCD_Des-Tem-HD-IE_MR-CT-FL", "CCD_Desert-Temecula", "CCD_Fresno-Bakersfield", "CCD_Fresno-Bakersfield_MR-CT-FL", "CCD_Grove-Riverside", "CCD_OC_LB_MR-CT-FL", "CCD_Orange-LongBeach", "CCD_SFV-LA-SGV", "CCD_SFV-LA-SGV-VEN_MR-CT-FL", "CCD_SoCal_MA_Callback", "CCD_Ventura-Victor_Valley"]

        teams_to_be_selected = [team.text for team in number_of_teams_found]
        # print(teams_to_be_selected)

        for index, team in enumerate(teams_to_be_selected, 0):
            if team in ccd_teams:
                selected = number_of_teams_found[index]
                selected.click()
                # index += 1
        search = driver.find_element(By.ID, "search")
        search.click()
        check_all_good = True

    except Exception as e:
        print(e)
        print("\nThere was an error, reopening script..")
print("\nDone..")

time.sleep(10)
import subprocess
import pyautogui
from keyboard import press, release

# function to sort by largest to smallest
def sort_results():
    sort_button = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/section/div/div[2]/div/div[1]/div[2]/div/table/thead/tr/th[6]")
    sort_button.click() # one click sorts from smallest to largest
    sort_button.click() # two sorts from largest to smallest

def close__error_notifications():
    time_now = datetime.strftime(datetime.now(), "%Y/%M/%D, %H:%M:%S")
    timespan = str(time_now[8:])
    time.sleep(5)
    try:
        print("\n" + str(timespan) +": Closing notifications..")
        close_error = driver.find_elements(By.XPATH, "/html/body/div[3]/div/button")
        for notification in close_error:
            notification.click()
    except:
        pass

def clear_and_search(auxiliar):
    time_now = datetime.strftime(datetime.now(), "%Y/%M/%D, %H:%M:%S")
    timespan = str(time_now[8:])

    # get the search bar's reference
    done = False
    while done == False:
        try:
            print("\n" + str(timespan) + f": Searching {auxiliar}..")
            time.sleep(5)
            search_aux_bar = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/section/div/div[2]/div/div[1]/div[1]/div[2]/div/label/input")
            # clear and type
            search_aux_bar.clear()
            search_aux_bar.send_keys(auxiliar)
            done = True
        except: 
            print("\n" + str(timespan) +": Couldn't find element, trying again..")
            time.sleep(5)

def search_button():
        time_now = datetime.strftime(datetime.now(), "%Y/%M/%D, %H:%M:%S")
        timespan = str(time_now[8:])

        # refresh the dashboard
        done = False
        while done == False:
            try:
                print("\n" + str(timespan) +": Refreshing..")
                search = driver.find_element(By.ID, "search")
                search.click()
                done = True
            except:
                print("\n" + str(timespan) +": Something got in the way, trying again..")

def search_acw():
    #close error message if there is any
    close__error_notifications()

    # refresh results
    search_button()

    # search for aux
    clear_and_search("After Call Work")

    # sort results
    sort_results()

    time.sleep(2)

    time_now = datetime.strftime(datetime.now(), "%Y/%M/%D, %H:%M:%S")
    timespan = str(time_now[8:])

    # get the search bar's element again so we can check
    # if it was cleared before continuing
    search_aux_bar2 = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/section/div/div[2]/div/div[1]/div[1]/div[2]/div/label/input")
    search_aux_bar2_text = search_aux_bar2.get_attribute("value")
    if search_aux_bar2_text != "After Call Work":
        print("\n" + timespan +": Stopping function, search bar was cleared..")
        return

    # get how much time the agents have in X aux
    results_elements = driver.find_elements(By.XPATH, "/html/body/div[2]/div[2]/div/section/div/div[2]/div/div[1]/div[2]/div/table/tbody/tr/td[6]/div/div/span") # returns the elements
    results_elements_text = [element.text for element in results_elements] # returns the element's text

    # get the agent's names that matches the results
    agent_names = driver.find_elements(By.XPATH, "/html/body/div[2]/div[2]/div/section/div/div[2]/div/div[1]/div[2]/div/table/tbody/tr/td[1]")
    agent_names_texts = [name.text for name in agent_names]

    # get the sign out button's reference
    sign_out = driver.find_elements(By.XPATH, "/html/body/div[2]/div[2]/div/section/div/div[2]/div/div[1]/div[2]/div/table/tbody/tr/td[8]/span")

    # goes through each name that matches the search
    # then checks their duration on said aux
    # and kicks them out if it matches the limit
    if results_elements_text == "":
        print("\n" + timespan +": Stopping function, no results found..")
        return
    for index in range(0, len(results_elements) + 1, 1):
            # last agent logged out
            last_logged_out = ""
            try:
                limit = datetime.strptime("00:02:00", "%H:%M:%S")
                element = results_elements_text[index]
                duration = datetime.strptime(element, "%H:%M:%S")
                if last_logged_out == agent_names_texts[index]:
                        print("\n" + timespan +": Ignoring agent, already called out")
                        continue # ignore this agent if it was already called out
                                 # since we could be logging someone else out by accident
                if duration >= limit:
                    print("\n" + timespan + ": " + agent_names_texts[index] + " logged out..")
                    last_logged_out += agent_names_texts[index]
                    sign_out[index].click() # log out

            except Exception as e:
                print(e)
                break
    time.sleep(2)

def search_auxiliars(aux):
    #close error message if there is any
    close__error_notifications()

    # refresh results
    search_button()

    # search for aux
    clear_and_search(aux)

    # sort by largest to smallest
    sort_results()

    time.sleep(2)

    time_now = datetime.strftime(datetime.now(), "%Y/%M/%D, %H:%M:%S")
    timespan = str(time_now[8:])

    # get the search bar's element again so we can check
    # if it was cleared before continuing
    search_aux_bar2 = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/section/div/div[2]/div/div[1]/div[1]/div[2]/div/label/input")
    search_aux_bar2_text = search_aux_bar2.get_attribute("value")
    if search_aux_bar2_text != aux:
        print("\n" + timespan +": Stopping function, search bar was cleared..")
        return
    
    # get how much time the agents have in X aux
    results_elements = driver.find_elements(By.XPATH, "/html/body/div[2]/div[2]/div/section/div/div[2]/div/div[1]/div[2]/div/table/tbody/tr/td[6]/div/div/span") # returns the elements
    results_elements_text = [element.text for element in results_elements] # returns the element's texts

    names_found = False
    while names_found == False:
        try:
            # get the agent's names that matches the results
            agent_names = driver.find_elements(By.XPATH, "/html/body/div[2]/div[2]/div/section/div/div[2]/div/div[1]/div[2]/div/table/tbody/tr/td[1]")
            agent_names_texts = [name.text for name in agent_names]
            names_found = True
        except:
            print("\n" + timespan +": Names not found, trying again..")
            time.sleep(5)

    call_out = "" # agents getting logged out
    limit_time = "" # threshold before logging out
    message = "" # personalized message for each aux

    # get the sign out button's reference
    sign_out = driver.find_elements(By.XPATH, "/html/body/div[2]/div[2]/div/section/div/div[2]/div/div[1]/div[2]/div/table/tbody/tr/td[8]/span")

    # get the search bar's element again so we can check
    # if it was cleared before continuing
    search_aux_bar2 = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/section/div/div[2]/div/div[1]/div[1]/div[2]/div/label/input")

    # goes through each name that matches the search
    # then checks their duration on said aux
    # and kicks them out if it matches the limit
    if results_elements_text == "":
        return
    for index in range(0, len(results_elements) + 1, 1):
            # last agent logged out
            last_logged_out = ""
            try:
                if aux == "Not Responding":
                    limit_time = "00:00:01"
                    message = "Agents not responding logged out: "
                elif aux == "Unavailable":
                    limit_time = "00:09:30"
                    message = f"Agents with more than 10 minutes in {aux} logged out: "
                elif aux == "Break":
                    limit_time = "00:16:30"
                    message = f"Agents with more than 17 minutes in {aux} logged out: "
                elif aux == "Lunch":
                    limit_time = "00:31:30"
                    message = f"Agents with more than 32 minutes in {aux} logged out: "

                limit = datetime.strptime(limit_time, "%H:%M:%S")
                element = results_elements_text[index]
                duration = datetime.strptime(element, "%H:%M:%S")
                if last_logged_out == agent_names_texts[index]: 
                    continue # ignore this agent if it was already called out
                             # since we could be logging someone else out by accident
                if duration > limit:
                    print("\n" + timespan + ": " + agent_names_texts[index] + " logged out..")
                    call_out += "\n" + (agent_names_texts[index])
                    last_logged_out += agent_names_texts[index]
                    sign_out[index].click() # log out
            except:
                break
    time.sleep(2)

    if call_out != "":
        print("\n" + timespan +": Sending call outs..")
        subprocess.call(r"C:\Users\wrr20\AppData\Local\CiscoSparkLauncher\CiscoCollabHost") # opening the messaging app to send the call outs
        time.sleep(1)
        press("ctrl+a")
        press("backspace")
        release("ctrl+a")
        release("backspace")
        pyautogui.typewrite(message + f"{call_out}") # paste call outs
        press("enter")

# module to schedule code
import schedule

# schedule each function to run periodically
schedule.every(2).to(3).minutes.do(search_acw)
schedule.every(5).minutes.do(search_auxiliars, aux = "Unavailable")
schedule.every(30).seconds.do(search_auxiliars, aux= "Not Responding")
schedule.every(3).to(5).minutes.do(search_auxiliars, aux = "Break")
schedule.every(4).to(7).minutes.do(search_auxiliars, aux ="Lunch")

while True:
    schedule.run_pending()
    time.sleep(5)

exit()
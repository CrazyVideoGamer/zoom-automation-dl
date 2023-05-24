from pathlib import Path
from time import sleep

import pyautogui
from pywinauto.application import Application
# from pywinauto import Desktop
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

DOWNLOAD_LOCATION = r"C:\Users\lianj\Documents\Coding\selenium-zoom-test"
FILENAME = "Recording"
FILEEXT = ".mp4"

path = Path(DOWNLOAD_LOCATION) / (FILENAME + FILEEXT)
path.unlink(missing_ok=True)

def download_complete(driver):
  def inner(driver):
    if not driver.current_url.startswith("chrome://downloads"):
      driver.get("chrome://downloads/")
      WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "downloads-manager")))
    return driver.execute_script("""
      let items = document.querySelector('downloads-manager')
        .shadowRoot.getElementById('downloadsList').items;
      if (items[0] === undefined) {
        return false;
      }
      if (items[0].state === "COMPLETE")
        return true;
      """)
  return WebDriverWait(driver, 1800).until(inner)

prefs = {
  "download.prompt_for_download": False,
  "download.default_directory": DOWNLOAD_LOCATION,
  "profile.default_content_setting_values.automatic_downloads": True,
  "profile.default_content_settings.popups": False
}
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_experimental_option("prefs", prefs)
# options.add_argument('ignore-certificate-errors')

driver = webdriver.Chrome(options=options)

try:
  driver.get("https://us02web.zoom.us/rec/play/ToU9UQTiwY2LZvToaELslGRf_kHL7NSlEwz1a64kx65MSyh83nUz0g0g-QbsW7Jnj6LkzQssfbO89D82.FCGTjlattB3-tWki?canPlayFromShare=true&from=share_recording_detail&continueMode=true&componentName=rec-play&originRequestUrl=https%3A%2F%2Fus02web.zoom.us%2Frec%2Fshare%2Fj3jK92NgIOpeiCmu8egt4FyW3cXHit-9CutpRhqnYNPhlwsMyzLITSHyfiIBp8dV.DdD_-BtP79gANHXw")
  WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#vjs_video_3")))
  driver.execute_script("document.addEventListener('contextmenu', function(e) { e.stopPropagation(); }, true);") # the magic

  # sleep(15)
  video = driver.find_element(By.CSS_SELECTOR, "#vjs_video_3")
  ActionChains(driver).context_click(video).perform()

  sleep(0.1)
  pyautogui.press('down')
  pyautogui.press('down')
  pyautogui.press('down')
  pyautogui.press('down')
  pyautogui.press('enter')
  
  sleep(0.4)

  app = Application(backend="win32").connect(path=r"C:\Program Files (x86)\Google\Chrome\Application", timeout=10)
  dialog = app.window(title="Save As")
  pyautogui.typewrite(FILENAME)
  address = dialog.child_window(title="Address: Downloads", class_name="ToolbarWindow32")
  address.click()
  pyautogui.write(DOWNLOAD_LOCATION)
  pyautogui.press("enter")

  save_button = dialog["Save"]
  save_button.click()

  download_complete(driver)
  print("Download Complete.")

  driver.quit()

finally:
  driver.quit()
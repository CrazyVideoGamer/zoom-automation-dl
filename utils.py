import pathlib
import sys
from time import sleep

import pyautogui
import win32com.client
import win32gui
from pywinauto.application import Application
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


def focus_window(window_handle):
  win32com.client.Dispatch("WScript.Shell").SendKeys("%")
  win32gui.SetForegroundWindow(window_handle)

def replace_existing(path: pathlib.Path): 
  path.unlink(missing_ok=True)

def download_complete(driver: WebDriver):
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
  WebDriverWait(driver, 1800).until(inner)

def download_recording(driver: WebDriver, window_handle, links: str, download_dir: pathlib.Path, prefix: list):
  for index, link in enumerate(links):
    driver.get(link)
    WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#vjs_video_3")))
    script = """
    document.addEventListener('contextmenu', function(e) { 
      e.stopPropagation(); 
    }, true);
    """
    driver.execute_script(script) # the magic (disables any right click overwrites)

    video = driver.find_element(By.CSS_SELECTOR, "#vjs_video_3")

    focus_window(window_handle)
    ActionChains(driver).context_click(video).perform()

    sleep(0.1)
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('down')
    pyautogui.press('enter')
    
    sleep(0.4)

    filename = prefix + " " + str(index + 1)

    app = Application(backend="win32").connect(path=r"C:\Program Files (x86)\Google\Chrome\Application", timeout=10)
    dialog = app.window(title="Save As")
    pyautogui.typewrite(filename)
    sleep(0.6)
    address = dialog.child_window(title="Address: Downloads", class_name="ToolbarWindow32")
    address.click()
    pyautogui.write(str(download_dir.parent.resolve()))
    pyautogui.press("enter")

    save_button = dialog["Save"]
    save_button.click()

    download_complete(driver)
    print(f'"{filename}" download complete.')

  driver.quit()
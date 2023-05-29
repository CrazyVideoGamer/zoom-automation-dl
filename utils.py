import pathlib
import sys
from time import sleep

import pyautogui
#from pywinauto.application import Application
import uiautomation as auto
import win32com.client
import win32gui
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


def focus_window(window_handle, chrome=True):
  win32com.client.Dispatch("WScript.Shell").SendKeys("%")
  pyautogui.press("esc")
  win32gui.SetForegroundWindow(window_handle)

def replace_existing(path: pathlib.Path): 
  path.unlink(missing_ok=True)
  (path.parent / f"{path.name}.crdownload").unlink(missing_ok=True)

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
    
    filename = prefix + " " + str(index + 1)

    # time = 0
    # while True:
    #   if time > 10:
    #     raise TimeoutError("Timeout expired (10s): Save As window did not open.")

    #   title = pyautogui.getActiveWindowTitle()
    #   if title == "Save As":
    #     break
    #   sleep(0.4)
    #   time += 0.4

    dialog = auto.WindowControl(searchDepth=2, Name="Save As")
    if not dialog.Exists(10, 0.4):
      raise TimeoutError("Timeout expired (10s): Save As window did not open.")

    focus_window(dialog.NativeWindowHandle, False)
    pyautogui.typewrite(filename)

    address = dialog.ToolBarControl(searchDepth=6, RegexName="Address: [a-zA-Z]+")
    focus_window(dialog.NativeWindowHandle, False)
    pyautogui.click(address.BoundingRectangle.left, address.BoundingRectangle.top)
    pyautogui.typewrite(str(download_dir.resolve()))
    pyautogui.press("enter")

    save = dialog.ButtonControl(searchDepth=1, Name="Save")
    focus_window(dialog.NativeWindowHandle, False)
    pyautogui.click(save.BoundingRectangle.xcenter(), save.BoundingRectangle.ycenter())

    download_complete(driver)
    print(f'"{filename}" download complete.')

  driver.quit()

import pathlib
from time import sleep
import sys
import typing

import pyautogui
import win32con
import win32api
import win32gui
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

def set_focus(window_handle: int):
  if window_handle != win32gui.GetForegroundWindow():
    pyautogui.keyDown("alt") # this method avoids possible shortcuts in other windows
    pyautogui.press("tab")
    pyautogui.press("esc")
    pyautogui.keyUp("alt")

    win32gui.SetForegroundWindow(window_handle)

def replace_existing(path: pathlib.Path): 
  path.unlink(missing_ok=True)
  (path.parent / f"{path.name}.crdownload").unlink(missing_ok=True)

class DownloadInterruptedError(Exception):
  pass

class DownloadMissingTimeoutError(TimeoutError):
  pass

class DownloadsPageTimeoutError(TimeoutError):
  pass

def download_complete(driver: WebDriver, timeout: float = 1800):
  def inner(driver: WebDriver):
    if not driver.current_url.startswith("chrome://downloads"):
      driver.get("chrome://downloads/")
      try:
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, "downloads-manager")))
      except:
        raise DownloadsPageTimeoutError("chrome://downloads is not loading (timeout: 30 sec) (perhaps connection/wifi issues?)")
    def wait_until_not_missing(driver: WebDriver):
      return driver.execute_script("""
        let items = document.querySelector('downloads-manager')
          .shadowRoot.getElementById('downloadsList').items;
        return items[0] !== undefined
      """)
    try:
      WebDriverWait(driver, 5).until(wait_until_not_missing)
    except:
      raise DownloadMissingTimeoutError('Download missing on chrome://downloads (timeout: 5 sec) (perhaps user cancelled Save As dialog box?)')

    state = driver.execute_script("""
      let items = document.querySelector('downloads-manager')
        .shadowRoot.getElementById('downloadsList').items;

      return items[0].state
      """)
    if state == "IN_PROGRESS":
      return False
    elif state == "COMPLETE":
      return True
    elif state == "PAUSED":
      driver.find_element(By.TAG_NAME, "downloads-manager") \
        .shadow_root.find_element(By.ID, "frb0") \
        .shadow_root.find_element(By.ID, "pauseOrResume").click()
      pass

    elif state == "CANCELLED" or state == "PAUSED":
      raise DownloadInterruptedError("Download was unexpectedly paused/cancelled (perhaps by user?)")
    else:
      raise Exception(f"Unexpected state: \"{state}\"")
  
  WebDriverWait(driver, timeout).until(inner)

def navigate_file_explorer_pywinauto(filename: str, download_dir: pathlib.Path, activate_windows: bool = False):
  from pywinauto.application import Application

  app = Application(backend="win32").connect(path=r"C:\Program Files (x86)\Google\Chrome\Application", timeout=10)

  dialog = app.window(class_name="#32770", title="Save As")
  if not dialog.exists(10, 0.4):
    raise TimeoutError("Timeout expired (10s): Save As dialog box did not open.")

  activate_windows and dialog.set_focus()
  pyautogui.write(filename)

  address = dialog.child_window(class_name="ToolbarWindow32", title_re="Address: [a-zA-Z]+")
  activate_windows and dialog.set_focus()
  address.click()
  pyautogui.write(str(download_dir.parent.resolve()))
  pyautogui.press("enter")

  save_button = dialog["Save"]
  activate_windows and dialog.set_focus()
  save_button.click()

def navigate_file_explorer_uiautomation(filename: str, download_dir: pathlib.Path, activate_windows: bool = False):
  import uiautomation as auto

  dialog = auto.WindowControl(searchDepth=2, ClassName="#32770", Name="Save As")
  if not dialog.Exists(10, 0.4):
    raise TimeoutError("Timeout expired (10s): Save As window did not open.")

  activate_windows and set_focus(dialog.NativeWindowHandle)

  pyautogui.write(filename)

  address = dialog.ToolBarControl(searchDepth=6, ClassName="ToolbarWindow32", RegexName="Address: [a-zA-Z]+")
  activate_windows and set_focus(dialog.NativeWindowHandle)
  pyautogui.click(address.BoundingRectangle.left, address.BoundingRectangle.top)
  pyautogui.write(str(download_dir.resolve()))
  pyautogui.press("enter")

  save = dialog.ButtonControl(searchDepth=1, Name="Save")
  activate_windows and set_focus(dialog.NativeWindowHandle)
  pyautogui.click(save.BoundingRectangle.xcenter(), save.BoundingRectangle.ycenter())

def navigate_zoom_link(driver: WebDriver, link: str, window_handle: (int | None) = None):
  driver.get(link)
  WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CSS_SELECTOR, "#vjs_video_3")))
  script = """
  document.addEventListener('contextmenu', function(e) { 
    e.stopPropagation(); 
  }, true);
  """
  driver.execute_script(script) # the magic (disables any right click overwrites)

  video = driver.find_element(By.CSS_SELECTOR, "#vjs_video_3")

  window_handle != None and set_focus(window_handle)

  ActionChains(driver).context_click(video).perform()

  sleep(0.1)
  pyautogui.press('down')
  pyautogui.press('down')
  pyautogui.press('down')
  pyautogui.press('down')
  pyautogui.press('enter')

class DownloadOptions(typing.TypedDict):
  window_handle: int | None
  download_dir: pathlib.Path
  prefix: str
  download_timeout: float
  backend: typing.Literal["pywinauto", "uiautomation"]

def download_recording(driver: WebDriver, links: list[str], options: DownloadOptions):
  options = { 
    "window_handle": None, 
    "download_dir": pathlib.Path("."),
    "prefix": "Recording",
    "download_timeout": 1800,
    "backend": "pywinauto"
  } | options

  for index, link in enumerate(links):
    navigate_zoom_link(driver, link, options["window_handle"])
    
    filename = options["prefix"] + " " + str(index + 1)

    if options["backend"] == "pywinauto":
      navigate_file_explorer_pywinauto(filename, options["download_dir"], options["window_handle"] != None)
    elif options["backend"] == "uiautomation":
      navigate_file_explorer_uiautomation(filename, options["download_dir"], options["window_handle"] != None)
    else:
      print(options["backend"])

    try:
      download_complete(driver, options["download_timeout"])
    except Exception as e:
      raise e
    else:
      print(f'"{filename}" download complete.')
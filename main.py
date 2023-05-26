import argparse
import sys
from pathlib import Path
from time import sleep
from urllib.parse import urlparse

import win32gui
from selenium import webdriver
from selenium.common import exceptions as selenium_exceptions
from selenium.webdriver.chrome.service import Service

from utils import download_recording, replace_existing

# from selenium.webdriver.chrome.options import Options

def URL(arg):
  url = urlparse(arg)
  if all((url.scheme, url.netloc)) and "zoom" in url.netloc.lower():  # possibly other sections?
    return arg  # return url in case you need the parsed object
  raise argparse.ArgumentTypeError('Invalid Zoom URL')

parser = argparse.ArgumentParser(
  prog="zoom-dl",
  description="Automates the zoom recording downloading process"
)

parser.add_argument("links", type=URL, nargs="+", help="links to download")
parser.add_argument("--name", "-n", type=str, nargs=1, default="Recording", help="the name prefix for each file (the filename will be appended by a number starting from 1)")
parser.add_argument("--output", "-o", type=Path, nargs=1, default=Path("."), help="Video output directory")

args = parser.parse_args()

for index, _ in enumerate(args.links):
  replace_existing(args.output / Path(args.name + " " + str(index + 1) + ".mp4"))

prefs = {
  "download.prompt_for_download": False,
  "download.default_directory": str(args.output.resolve()),
  "profile.default_content_setting_values.automatic_downloads": True,
  "profile.default_content_settings.popups": False
}
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
options.add_experimental_option("prefs", prefs)
options.add_experimental_option('excludeSwitches', ['enable-logging'])
# options.add_argument('ignore-certificate-errors')

service = Service(executable_path="./ChromeDriver/chromedriver.exe")

driver = webdriver.Chrome(options=options, service=service)
window_handle = win32gui.GetForegroundWindow()

try:
  download_recording(driver, window_handle, args.links, args.output, args.name)
except selenium_exceptions.NoSuchWindowException:
  print("Error: browser was closed during execution.", file=sys.stderr)
finally:
  driver.quit()
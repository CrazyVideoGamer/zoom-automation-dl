import argparse
import sys
from pathlib import Path
from urllib.parse import urlparse

import win32gui
from selenium import webdriver
from selenium.common import exceptions as selenium_exceptions
from selenium.webdriver.chrome.service import Service

from utils import download_recording, replace_existing
from utils import DownloadInterruptedError, DownloadMissingTimeoutError, DownloadsPageTimeoutError

# from selenium.webdriver.chrome.options import Options

def URL(arg):
  url = urlparse(arg)
  if all((url.scheme, url.netloc)) and "zoom" in url.netloc.lower():  # possibly other sections?
    return arg  # return url in case you need the parsed object
  raise argparse.ArgumentTypeError('Invalid Zoom URL')

parser = argparse.ArgumentParser(
  prog="zoom-dl",
  description="Automates the Zoom recording download process",
  formatter_class=argparse.ArgumentDefaultsHelpFormatter
)

parser.add_argument("links", type=URL, nargs="+", help="links to download")
parser.add_argument("--name", "-n", type=str, default="Recording", help="the name prefix for each file (the filename will be appended by a number starting from 1)")
parser.add_argument("--output", "-o", type=Path, default=Path("."), help="Video output directory")
parser.add_argument("--timeout", "-t", type=float, default=1800, help="Maximum duration (seconds) for recording download until timeout")
parser.add_argument("--backend", "-b", type=str, choices={"pywinauto", "uiautomation"}, default="uiautomation", help="Backend to use (choose from 'pywinauto', 'uiautomation')")
parser.add_argument("--window-activation", action=argparse.BooleanOptionalAction, default=True, help="Enable feature to activate unfocused browser/dialogs that need to be in the foreground to run properly")

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

options = {
  "download_dir": args.output,
  "prefix": args.name,
  "download_timeout": args.timeout,
  "backend": args.backend
}

if args.window_activation:
  options["window_handle"] = window_handle

try:
  download_recording(driver, args.links, options)
except selenium_exceptions.NoSuchWindowException:
  print("BrowserError: browser was closed during execution", file=sys.stderr)
except (DownloadInterruptedError, DownloadMissingTimeoutError, DownloadsPageTimeoutError) as e:
  print(f"{type(e).__name__}: {str(e)}", file=sys.stderr)
except TimeoutError as e:
  print(f"TimeoutError: {str(e)}", file=sys.stderr)
except Exception as e:
  print("Unexpected error:", str(e), file=sys.stderr)
else:
  driver.quit()
  for tmp in Path(".").glob("*.tmp"):
    tmp.unlink()
  quit(0)

driver.quit()
for tmp in Path(".").glob("*.tmp"):
  tmp.unlink()
quit(1)

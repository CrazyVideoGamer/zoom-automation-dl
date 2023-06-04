# Zoom Recording Downloader

🤖Automates zoom recording downloading with selenium

- Doesn't need any cookies or any manual searching to find the download link!

```
usage: zoom-dl [-h] [--name NAME] [--output OUTPUT] [--timeout TIMEOUT] [--backend {pywinauto,uiautomation}]
               [--window-activation | --no-window-activation]
               links [links ...]

Automates the Zoom recording download process

positional arguments:
  links                 links to download

options:
  -h, --help            show this help message and exit
  --name NAME, -n NAME  the name prefix for each file (the filename will be appended by a number starting from 1) (default:    
                        Recording)
  --output OUTPUT, -o OUTPUT
                        Video output directory (default: .)
  --timeout TIMEOUT, -t TIMEOUT
                        Maximum duration (seconds) for recording download until timeout (default: 1800)
  --backend {pywinauto,uiautomation}, -b {pywinauto,uiautomation}
                        Backend to use (choose from 'pywinauto', 'uiautomation') (default: pywinauto)
  --window-activation, --no-window-activation
                        Enable feature to activate unfocused browser/dialogs that need to be in the foreground to run
                        properly (default: True)
```

Backend differences:
- uiautomation: less asethetic focusing capability
- pywinauto: may be a bit slow/inconsistent

from pywinauto.application import Application
import time

# Run a target application
app = Application().start("C:\\Program Files (x86)\\Micro Focus\\Analysis\\bin\\AnalysisUI.exe")
# Select a menu item
app.LoadRunnerAnalysisUI.menu_select("File -> 1")
# Click on a button
'Open Raw Result for the new Analysis Session'
app.OpenRawResultforthenewAnalysisSession.Cancel.click()
# kill app
time.sleep(10)
app.kill()

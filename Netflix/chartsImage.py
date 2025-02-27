import os
import time
import xlwings as xw
import pyautogui
import subprocess
import DBToExcelSummary
from AppKit import NSWorkspace
from Quartz import CGWindowListCopyWindowInfo, kCGWindowListOptionOnScreenOnly, kCGNullWindowID

# Define the Excel file path
excel_path = DBToExcelSummary.excel_path
#"/Users/sandyaudumala/Netflix/NetflixEngagementWBR_2025-02-27.xlsx"

# Extract the directory to save the screenshot
save_dir = os.path.dirname(excel_path)
screenshot_path = os.path.join(save_dir, "SummaryChartsScreenshot.png")

# Open the Excel file using xlwings
app = xw.App(visible=True)  # Open Excel with UI
wb = xw.Book(excel_path)    # Open the workbook

try:
    # Activate the 'Summary Charts' sheet
    sheet = wb.sheets["Summary Charts"]
    sheet.activate()

    # Ensure Excel is the frontmost application using AppleScript
    subprocess.run(["osascript", "-e", 'tell application "Microsoft Excel" to activate'])

    # Allow some time for the UI to update and window to be brought to front
    time.sleep(2)

    # Get the active application (Excel)
    active_app = NSWorkspace.sharedWorkspace().frontmostApplication()
    app_name = active_app.localizedName()  # Name of the frontmost app

    # Get all windows
    window_list = CGWindowListCopyWindowInfo(kCGWindowListOptionOnScreenOnly, kCGNullWindowID)

    # Find Excel's window
    excel_window = None
    for window in window_list:
        if "kCGWindowOwnerName" in window and window["kCGWindowOwnerName"] == app_name:
            excel_window = window
            break

    if excel_window:
        # Get window bounds
        bounds = excel_window.get("kCGWindowBounds", {})
        x, y, width, height = int(bounds.get("X", 0)), int(bounds.get("Y", 0)), int(bounds.get("Width", 0)), int(bounds.get("Height", 0))

        # Take a screenshot of just the Excel window
        screenshot = pyautogui.screenshot(region=(x, y, width, height))
        
        # Save the screenshot
        screenshot.save(screenshot_path)
        print(f"Screenshot saved at: {screenshot_path}")
    else:
        print("Could not detect the Excel window position. Taking full screen screenshot as fallback.")
        screenshot = pyautogui.screenshot()
        screenshot.save(screenshot_path)

finally:
    # Close the workbook and quit Excel
    wb.close()
    app.quit()
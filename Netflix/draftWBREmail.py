import subprocess
import datetime
import chartsImage

# Get current date in YYYY-MM-DD format
current_date = datetime.datetime.now().strftime("%Y-%m-%d")

# Email subject with date
subject = f"Engagement Metrics WBR - {current_date}"

# Placeholder email body with embedded image
body = """Hello Team, 

Happy monday!
Please find the latest business metrics for last week 

Here are the three key takeaways:
1. [Takeaway 1]
2. [Takeaway 2]
3. [Takeaway 3]

"""

# File path to the image
image_path = chartsImage.screenshot_path
#"/Users/sandyaudumala/Netflix/SummaryChartsScreenshot.png"

# AppleScript command to draft email
applescript = f'''
tell application "Mail"
    set newMessage to make new outgoing message with properties {{subject:"{subject}", visible:true}}
    tell newMessage
        set content to "{body}"
        set imgAttachment to make new attachment with properties {{file name:"{image_path}"}} at after last paragraph
    end tell
    activate
end tell
'''

# Execute AppleScript
subprocess.run(["osascript", "-e", applescript])

print("Draft email created in Mail app with embedded image.")
#Import all the modules necessary to run this honker
import subprocess
import re
import json
from github import Github
from datetime import date
import smtplib
from email.mime.text import MIMEText

#variables that can be set right now
today = date.today()
date_formatted = today.strftime("%m/%d/%y")
git_json_file = '/path/to/mmgit.json'
mmctljson = '/path/to/mmctl.json'
bash_command = "mmctl version"
repo_path = "mattermost/mattermost"
token = None # Put your GitHub API token here if you want to access a private repo.



    
#Define the Github version check process
def process_git_data(git_json_data, latest, git_json_file, git_mm_json):
    # Reading the JSON file
    with open(git_json_file, "r") as f:
        git_json_data = json.load(f)
    if __name__ == "__main__":
        g = Github(token)
        repo = g.get_repo(repo_path)
        latest = repo.get_latest_release()
    if git_json_data['latest'] == latest.title:
        print(git_json_data['latest'])
        #setting json format and data, also including date last checked in case I ever need to troubleshoot. Why lose sanity earlier than necessary?
        git_mm_json = {'latest':latest.title,'datelastchecked':date_formatted}
        with open(git_json_file, "w") as f:
            json.dump(git_mm_json, f)
    

def send_email():
    msg = MIMEText(body)
    #Email variables, could probably set these earlier but .. eh.
    subject = "MatterMost Upgrade " + latest.title
    body = "This is a notification to inform you that the latest release for MatterMost has been updated to: " + latest.title
    sender = "email@yourdomain.com"
    recipients = ["email@yourdomain.com"]
    password = "apasswordgoeshere"
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ', '.join(recipients)
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp_server:
        smtp_server.login(sender, password)
        smtp_server.sendmail(sender, recipients, msg.as_string())
                
        send_email(subject, body, sender, recipients, password)
        print("Message sent!")
        

#Define the mmctl version check function
#This code block runs the "mmctl version" command on the local device, grabs the output, formats it to json, prepends "v" to the version number, writes to json. 
#The output json file will be used for comparing against the Mattermost Github "lastest" release channel
def get_mmctl_info(completed_process, mmctljson):
    # Check if the command was successful
    if completed_process.returncode == 0:
        # Define a regular expression pattern to match key-value pairs
        pattern = re.compile(r'(\w+):\s*(.*)')

        # Initialize a dictionary to store the extracted information
        result_dict = {}

        # Iterate over the lines in the command output and extract key-value pairs
        for line in completed_process.stdout.split('\n'):
            match = pattern.match(line)
            if match:
                key, value = match.groups()
                # Modify the version value to include "v" prefix
                if key == 'Version':
                    value = 'v' + value

                result_dict[key] = value.strip()

        # Convert the dictionary to JSON
        mmctl_json_output = json.dumps(result_dict, indent=2, separators=(',', ': '), ensure_ascii=False)
        
        with open(mmctljson, 'w') as f:
            json.dump(result_dict, f)
        
        with open(mmctljson, 'r') as f:
            mmctl_test = json.load(f)
            
        #Testing the above
        #print(mmctl_test['Version'])
    else:
        print(f"Error running the BASH command. Return code: {completed_process.returncode}")
        print(completed_process.stderr)
        
globals()['get_mmctl_info'(completed_process, mmctljson)]()
process_git_data(git_json_data, latest, git_json_file, git_mm_json)
if mmctl_test['Version']!=latest.title:
    send_email()
else:
    print("The other thing happened!")    

# dcloud-sessions

**dcloud VPN Session Report Generator**

This Python script automates the process of retrieving VPN details (username, password, and server) for a specified range of dcloud sessions using the dcloud API. It then compiles this information into a formatted Word document, with each session's details on a separate page.

Prerequisites
To run this script, you must have the following Python libraries installed:

requests for making API calls.

python-docx for creating and managing Word documents.

You can install them using pip:

**pip install requests python-docx**

**How to Use
**

Obtain a Bearer Token: You need a valid Bearer Token from the dcloud API. This token is required for authentication and has a limited lifespan. You will need to get a new one if it expires.

Run the Script: Execute the script from your terminal. It will prompt you for the necessary information.

**python user-inputs-dcloud.py**

Provide Input: When prompted, enter the following information:

Bearer Token: Paste your dcloud API bearer token.

Starting Session ID: Enter the first session ID in the range you want to query (e.g., 1250385).

Ending Session ID: Enter the last session ID in the range you want to query (e.g., 1250434). The script will include this session ID in the report.

Output Filename: Enter the desired name for the Word document (e.g., vpn_report.docx).

Example of User Input
--- VPN Session Report Configuration ---
Enter your Bearer Token: your_long_bearer_token_here
Enter the starting session ID: 1250385
Enter the ending session ID: 1250434
Enter the output filename (e.g., vpn_report.docx): my_vpn_sessions.docx

After you provide the input, the script will begin querying the API and creating the Word document. A confirmation message will be printed to the console once the file is saved.

# msTeamsUserDetailReport
This script uses Microsoft Graph to obtain Microsoft Teams usage details per user. Statistics are stored in a SharePoint Online list, and a daily summary is posted to a Teams channel using the generic incoming webhook. 

# Usage
Register an Azure application. Log into http://portal.azure.com and [create a new application](http://lisa.rushworth.us/?p=3945) with application permissions to use Reports.Read.All

[Create a generic webhook URL for your Teams channel.](http://lisa.rushworth.us/?p=3992)

In the root directory, copy key.sample to key.py and generate a new key:

from cryptography.fernet import Fernet
# Put this somewhere safe!
key = Fernet.generate_key()
print("The key is %s" % key)


Encrypt credentials with access to your SPO list:
from cryptography.fernet import Fernet
token = f.encrypt(b"uid@example.com")
print("The crypted version is %s" % token)

token = f.encrypt(b"R3Al|yG0.dPwdG03s!|-|3rE")
print("The crypted version is %s" % token)

In each subfolder, copy config.sample to config.py and modify with *your* application registration information, tenant auth URL, SPO credentials, and webhook URL.


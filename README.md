# MS Teams User and Group Statistics

This script uses Microsoft Graph to obtain Microsoft Teams usage details per user. Statistics are stored in SharePoint Online lists.

# User Stats
**postTeamsStatsToSPO.py** -- updates individual usage statistics and posts a daily summary to a Teams channel using a webhook.

**updateUserData.py** -- Updates account status and reporting information

# Group Stats
**postGroupStatsToSPO.py** -- Inserts newly created Teams spaces, updates visibility, and creates a searchable HTML table of  public Teams

**getGroupActivity.py** -- Updates Team owner, member count, and date of latest Channel activity


# Usage
Data is stored in SharePoint Online lists using SPO's REST API. TableCrossReference.xlsx provides the list definitions. When creating lists, be aware that [internal column names may not match column display names.](http://lisa.rushworth.us/?p=4572)

Register an Azure application. Log into [http://portal.azure.com](http://portal.azure.com) and [create a new application](http://lisa.rushworth.us/?p=3945) with application permissions to use Reports.Read.All

[Create a generic webhook URL for your Teams channel.](http://lisa.rushworth.us/?p=3992)

In the root directory, copy key.sample to key.py and generate a new key:

    from cryptography.fernet import Fernet
    # Put this somewhere safe!
    key = Fernet.generate_key()
    print("The key is %s" % key)


Encrypt credentials with access to your SPO list:

    from cryptography.fernet import Fernet
    f = Fernet(strKey)
    token = f.encrypt(b"uid@example.com")
    print("The crypted version is %s" % token)
     
    token = f.encrypt(b"R3Al|yG0.dPwdG03s!|-|3rE")
    print("The crypted version is %s" % token)

In each subfolder, copy config.sample to config.py and modify with *your* application registration information, tenant auth URL, SPO credentials, and webhook URL.


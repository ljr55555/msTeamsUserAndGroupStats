import sys
from cryptography.fernet import Fernet
from base64 import b64encode, b64decode
import datetime
import requests
from requests_toolbelt.utils import dump
import json
from config import strListDataURI, strContextURL, strConnectURL, strUsername, strPassword, strListInfoURI,  strClientID, strClientSecret, strGraphAuthURL, strWebhookURL

# Use https://github.com/ljr55555/sharepy/tree/develop instead of PIP package
import sharepy

sys.path.append('../')
from key import strKey

################################################################################
# Function definitions
################################################################################
# This function creates a new record in a SharePoint list
# Input: s -- connection to  SharePoint REST API
#        strBody -- dictionary of data to POST
# Output: integer HTTP response
################################################################################
def writeNewRecord(s, strBody ):
    strContentType = "application/json;odata=verbose"
    
    # Get digest value for use in POST
    r = s.post(strContextURL)
    jsonDigestRaw = json.loads(r.text)
    jsonDigestValue = jsonDigestRaw['d']['GetContextWebInformation']['FormDigestValue']
    
    strBody  = json.dumps(strBody)

    postRecord = s.post(strListDataURI,headers={"Content-Length": str(len(json.dumps(strBody))), 'accept': strContentType, 'content-Type': strContentType, "X-RequestDigest": jsonDigestValue}, data=strBody)
    #data = dump.dump_all(postRecord)
    #print("Session data:\t%s" % data.decode('utf-8'))
    #print("HTTP Status Code:\t%s\nResult code content:\t%s" % (postRecord.status_code, postRecord.content))
    print("HTTP Status Code:\t%s" % postRecord.status_code)
    return postRecord.status_code

# Research references
#https://sharepoint.stackexchange.com/questions/105380/adding-new-list-item-using-rest?newreg=70a88b49ad694022a867ac3a6e434380
#https://docs.microsoft.com/en-us/sharepoint/dev/sp-add-ins/working-with-lists-and-list-items-with-rest

################################################################################
# This function finds the ID of a record
# Input: s -- connection to  SharePoint REST API
#        strAttr -- attribute on which to search
#        strValue -- attribute value for search
# Output: integer item ID
################################################################################
def findSPRecord(s, strAttr, strValue):
    strListContentURI = ("%s?&$filter=%s eq '%s'" % (strListDataURI, strAttr, strValue))
    r3 = s.get(strListContentURI)
    jsonReply = json.loads(r3.text)
    jsonListContent = jsonReply['d']
    #print("Length is %s" % len(jsonListContent['results']))
    if len(jsonListContent['results']) == 0:
        return None
    else:
        iItemID = jsonListContent['results'][0].get('ID')
        return iItemID

################################################################################
# This function updates an existing record in SharePoint
# Input: s -- connection to  SharePoint REST API
#        strListItemURI -- URI for list item to be updated
#        strBody -- dictionary of data to POST
# Output: integer HTTP response
################################################################################
def updateRecord(s, strListItemURI, strBody):
    strContentType = "application/json;odata=verbose"

    # Get digest value for use in POST
    r = s.post(strContextURL)
    jsonDigestRaw = json.loads(r.text)
    jsonDigestValue = jsonDigestRaw['d']['GetContextWebInformation']['FormDigestValue']

    strBody  = json.dumps(strBody)

    postRecord = s.post(strListItemURI,headers={"Content-Length": str(len(json.dumps(strBody))), 'accept': strContentType, 'content-Type': strContentType, "X-RequestDigest": jsonDigestValue, "IF-MATCH": "*", "X-HTTP-Method": "MERGE"}, data=strBody)
    #data = dump.dump_all(postRecord)
    #print("Session data:\t%s" % data.decode('utf-8'))
    #print("HTTP Status Code:\t%s\nResult code content:\t%s" % (postRecord.status_code, postRecord.content))
    print("HTTP Status Code:\t%s" % postRecord.status_code)
    return postRecord.status_code

################################################################################
# This function posts usage stats to Teams via webhook
# Requirements: json
# Input: strURL -- webhook url
#        iPrivateMessages, iTeamMessages, iCalls, iMeetings -- integer usage stats
#        strReportDate - datetime date for stats
# Output: BOOL -- TRUE on 200 HTTP code, FALSE on other HTTP response
################################################################################
def postStatsToTeams(strURL,iPrivateMessages,iTeamMessages,iCalls,iMeetings,strReportDate):
    try:
        strCardContent = '{"title": "Teams Usage Statistics","sections": [{"activityTitle": "Usage report for ' + yesterday.strftime('%Y-%m-%d') + '"}, {"title": "Details","facts": [{"name": "Private messages","value": "' + str(iPrivateMessages) + '"}, {"name": "Team messages","value": "' + str(iTeamMessages) + '"}, {"name": "Calls ","value": "' + str(iCalls) + '"}, {"name": "Meetings","value": "' + str(iMeetings) + '"}]}],"summary": "Teams Usage Statistics","potentialAction": [{"name": "View web report","target": ["https://csgdirsvcs.windstream.com:1977/o365Stats/msTeams.php"],"@context": "http://schema.org","@type": "ViewAction"}]}'
        jsonPostData = json.loads(strCardContent)

        if postDataToURL(strURL, json.dumps(jsonPostData),'application/json'):
            print("POST successful")
            return True
        else:
            print("POST failed")
            return False
    except Exception as e:
        print("ERROR Unexpected error: %s" % e)
    return False

################################################################################
# This function POSTs to a URL
# Requirements: requests
# Input: strURL -- url to which data is posted
#        strBody -- content to be sent as data
#        strContentType -- Content-Type definition
# Output: BOOL -- TRUE on 200 HTTP code, FALSE on other HTTP response
################################################################################
def postDataToURL(strURL, strBody, strContentType):
    if strURL is None:
        print("POST failed -- no URL provided")
        return False
    print("Sending POST request to strURL=%s" % strURL)
    print("Body: %s" % strBody)
    try:
        dictHeaders = {'Content-Type': strContentType}
        res = requests.post(strURL, headers=dictHeaders,data=strBody)
        print(res.text)
        if 200 <= res.status_code < 300:
            print("Receiver responded with HTTP status=%d" % res.status_code)
            return True
        else:
            print("POST failed -- receiver responded with HTTP status=%d" % res.status_code)
            return False
    except ValueError as e:
        print("POST failed -- Invalid URL: %s" % e)
    return False

################################################################################
# This function updates an existing record in SharePoint
# Input: s -- connection to  SharePoint REST API
# Output: integer HTTP response
################################################################################
# Before we get the statistics, determine if historic data maintenance is needed
################################################################################
def tableMaintenance(s, dateYesterday):
    # If this is the first of a year, we roll current year to previous year, 0 yearly data, 0 monthly data, and 0 daily data
    if (dateYesterday.day == 1) and (dateYesterday.month == 1):
        print("First of the year, performing data maintenance")
        # get all records & roll stuff
        strListContentURI = (strListDataURI)
        while strListContentURI:
            r3 = s.get(strListContentURI)
            jsonReply = json.loads(r3.text)
            jsonListContent = jsonReply['d']
            strListContentURI = jsonListContent.get("__next")
            for listRecord in jsonListContent['results']:
                iFoundItemID = listRecord['ID']
                strItemURI = ("%s(%s)" % (strListDataURI, iFoundItemID))

                iYearlyTeamChat = listRecord['t0gh']
                iYearlyPrivateChat = listRecord['eb7w']
                iYearlyCalls = listRecord['rjbz']
                iYearlyMeetings = listRecord['s4xd']

                dictRecordPatch = {"__metadata": { "type": strItemTypeName}, 'dailyTeamChat': '0', 'monthlyTeamChat': '0', 't0gh': '0', 'mjnm': '0', 'eacr': '0', 'rjbz': '0', 'l69z': '0', 'vymh': '0', 's4xd': '0', 'g1y8': '0', 'o0ru': '0', 'eb7w': '0', 'mrkm': iYearlyTeamChat, 'vt2k': iYearlyCalls , 'lu6i': iYearlyMeetings, 'ro6o': iYearlyPrivateChat}
                updateRecord(connectionSP, strItemURI, dictRecordPatch)
        print("Completed data maintenance -- gathering stats")

    # Else if this is the first of the month, we should zero monthly data and daily data
    elif (dateYesterday.day == 1):
        print("First of the month, performing data maintenance")
        # get all records & roll stuff
        strListContentURI = (strListDataURI)
        while strListContentURI:
            r3 = s.get(strListContentURI)
            jsonReply = json.loads(r3.text)
            jsonListContent = jsonReply['d']
            strListContentURI = jsonListContent.get("__next")
            for listRecord in jsonListContent['results']:
                iFoundItemID = listRecord['ID']
                strItemURI = ("%s(%s)" % (strListDataURI, iFoundItemID))

                iYearlyTeamChat = listRecord['t0gh']
                iYearlyPrivateChat = listRecord['eb7w']
                iYearlyCalls = listRecord['rjbz']
                iYearlyMeetings = listRecord['s4xd']

                dictRecordPatch = {"__metadata": { "type": strItemTypeName}, 'dailyTeamChat': '0', 'monthlyTeamChat': '0', 'mjnm': '0', 'eacr': '0', 'l69z': '0', 'vymh': '0', 'g1y8': '0', 'o0ru': '0'}
                updateRecord(connectionSP, strItemURI, dictRecordPatch)
        print("Completed data maintenance -- gathering stats")
################################################################################
# End of functions
################################################################################
f = Fernet(strKey)

strUID = f.decrypt(strUsername)
strUID = strUID.decode("utf-8")

strPass = f.decrypt(strPassword)
strPass = strPass.decode("utf-8")

# I frequently get all 0's for yesterdays data, so getting -2 to ensure we don't have bad info
yesterday = datetime.date.today() - datetime.timedelta(3)
print("Getting stats from %s" % yesterday)

connectionSP = sharepy.connect(strConnectURL,strUID,strPass)

## Get ListItemEntityTypeFullName from list
r = connectionSP.get(strListInfoURI)
jsonReply = json.loads(r.text)
strItemTypeName = jsonReply['d']['ListItemEntityTypeFullName']

tableMaintenance(connectionSP, yesterday)

postData = {"grant_type": "client_credentials","client_id" : strClientID,"client_secret": strClientSecret,"scope": "https://graph.microsoft.com/.default"}

r = requests.post(strGraphAuthURL, data=postData)

strJSONResponse = r.text
if len(strJSONResponse) > 5:
    jsonResponse = json.loads(strJSONResponse)

    strAccessToken = jsonResponse['access_token']

    getHeader = {"Authorization": "Bearer " + strAccessToken}

    print("Report call is https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(date={0})".format(yesterday.strftime('%Y-%m-%d')))
    r2 = requests.get("https://graph.microsoft.com/v1.0/reports/getTeamsUserActivityUserDetail(date={0})".format(yesterday.strftime('%Y-%m-%d')), headers=getHeader)
    strUsageReport = r2.text

    if 'Report Refresh Date' in strUsageReport:
        strUsageStats = []
        strUsageStats = strUsageReport.splitlines()
        strUsageHeader = strUsageStats[0].split(",")
        if (strUsageHeader[1] == 'User Principal Name') and (strUsageHeader[6] == 'Team Chat Message Count') and (strUsageHeader[7] == 'Private Chat Message Count') and (strUsageHeader[8] == 'Call Count') and (strUsageHeader[9] == 'Meeting Count'):
            i = 1
            iAllPrivateMessages = 0
            iAllTeamMessages = 0
            iAllCalls = 0
            iAllMeetings = 0
            while i < len(strUsageStats):
                strUserUsage = strUsageStats[i].split(",")
                strUser = strUserUsage[1]
                iTeamChat = int(strUserUsage[6])
                iPrivateChat = int(strUserUsage[7])
                iCalls = int(strUserUsage[8])
                iMeetings = int(strUserUsage[9])

                iAllPrivateMessages = iPrivateChat + iAllPrivateMessages
                iAllTeamMessages = iTeamChat + iAllTeamMessages
                iAllCalls = iCalls + iAllCalls
                iAllMeetings = iMeetings + iAllMeetings

                # remove domain component from user ID as we have a single tree/forest/domain
                strUserBreakout = strUser.split('@')
                strUser = strUserBreakout[0]
                # We have eight character user IDs, so db column uid has character limit.
                strUser = strUser[0:8]
#                print("%s:\t%s\t%s\t%s\t%s\n\n" % (strUser, iTeamChat, iPrivateChat, iCalls, iMeetings))
                iFoundItemID = findSPRecord(connectionSP,"Title", strUser)

                if iFoundItemID is not None:
                    strItemURI = ("%s(%s)" % (strListDataURI, iFoundItemID))
                    r4 = connectionSP.get(strItemURI)
                    jsonUserRecord = json.loads(r4.text)
                    # Increment monthly and yearly stats with current day stats
                    iMonthlyTeamChat = jsonUserRecord['d'].get('monthlyTeamChat') + iTeamChat
                    iMonthlyPrivateChat = jsonUserRecord['d'].get('o0ru') + iPrivateChat
                    iMonthlyCalls = jsonUserRecord['d'].get('eacr') + iCalls
                    iMonthlyMeetings = jsonUserRecord['d'].get('vymh') + iMeetings

                    iYearlyTeamChat = jsonUserRecord['d'].get('t0gh') + iTeamChat
                    iYearlyPrivateChat = jsonUserRecord['d'].get('eb7w') + iPrivateChat
                    iYearlyCalls = jsonUserRecord['d'].get('rjbz') + iCalls
                    iYearlyMeetings = jsonUserRecord['d'].get('s4xd') + iMeetings

                    dictRecordPatch = {"__metadata": { "type": strItemTypeName}, 'dailyTeamChat': iTeamChat, 'monthlyTeamChat': iMonthlyTeamChat, 't0gh': iYearlyTeamChat, 'mjnm': iCalls, 'eacr': iMonthlyCalls, 'rjbz': iYearlyCalls, 'l69z': iMeetings, 'vymh': iMonthlyMeetings, 's4xd': iYearlyMeetings, 'g1y8': iPrivateChat, 'o0ru': iMonthlyPrivateChat, 'eb7w': iYearlyPrivateChat}
                    updateRecord(connectionSP, strItemURI, dictRecordPatch)

                else:
                    # insert record, daily/monthly/yearly stats are all current day stats
                    print("Creating record for %s" % strUser)
                    dictRecord = {"__metadata": { "type": strItemTypeName}, 'Title': strUser, 'dailyTeamChat': iTeamChat, 'monthlyTeamChat': iTeamChat, 't0gh': iTeamChat, 'mjnm': iCalls, 'eacr': iCalls, 'rjbz': iCalls, 'l69z': iMeetings, 'vymh': iMeetings, 's4xd': iMeetings, 'g1y8': iPrivateChat, 'o0ru': iPrivateChat, 'eb7w': iPrivateChat}
                    writeNewRecord(connectionSP, dictRecord)
                i += 1
            postStatsToTeams(strWebhookURL,iAllPrivateMessages,iAllTeamMessages,iAllCalls,iAllMeetings,yesterday.strftime('%Y-%m-%d'))
        else:
            print("Header changed, need to verify code")
    else:
        print("User detail report does not have sufficient data: %s\n" % strUsageReport)
else:
    print("Auth failed: %s" % strJSONResponse)
# Format of CSV output from getTeamsUserActivityUserDetail
# 0     Report Refresh Date
# 1     User Principal Name
# 2     Last Activity Date
# 3     Is Deleted
# 4     Deleted Date
# 5     Assigned Products
# 6     Team Chat Message Count
# 7     Private Chat Message Count
# 8     Call Count
# 9     Meeting Count
# 10    Has Other Action
# 11    Report Period

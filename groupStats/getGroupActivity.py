import sys
from cryptography.fernet import Fernet
from base64 import b64encode, b64decode
from csv import reader
import io
import datetime
import requests
from requests_toolbelt.utils import dump
import json
from config import strGroupListDataURI, strContextURL, strConnectURL, strUsername, strPassword, strGroupListInfoURI,  strClientID, strClientSecret, strGraphAuthURL

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

    postRecord = s.post(strGroupListDataURI,headers={"Content-Length": str(len(json.dumps(strBody))), 'accept': strContentType, 'content-Type': strContentType, "X-RequestDigest": jsonDigestValue}, data=strBody)
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
    strGroupListContentURI = ("%s?&$filter=%s eq '%s'" % (strGroupListDataURI, strAttr, strValue))
    r3 = s.get(strGroupListContentURI)
    jsonReply = json.loads(r3.text)
    jsonListContent = jsonReply.get('d')
    if jsonListContent:
        #print("Length is %s" % len(jsonListContent['results']))
        if len(jsonListContent['results']) == 0:
            return None
        else:
            iItemID = jsonListContent['results'][0].get('ID')
            return iItemID
    else:
        return None

################################################################################
# This function updates an existing record in SharePoint
# Input: s -- connection to  SharePoint REST API
#        strGroupListItemURI -- URI for list item to be updated
#        strBody -- dictionary of data to POST
# Output: integer HTTP response
################################################################################
def updateRecord(s, strGroupListItemURI, strBody):
    strContentType = "application/json;odata=verbose"

    # Get digest value for use in POST
    r = s.post(strContextURL)
    jsonDigestRaw = json.loads(r.text)
    jsonDigestValue = jsonDigestRaw['d']['GetContextWebInformation']['FormDigestValue']

    strBody  = json.dumps(strBody)

    postRecord = s.post(strGroupListItemURI,headers={"Content-Length": str(len(json.dumps(strBody))), 'accept': strContentType, 'content-Type': strContentType, "X-RequestDigest": jsonDigestValue, "IF-MATCH": "*", "X-HTTP-Method": "MERGE"}, data=strBody)
    #data = dump.dump_all(postRecord)
    #print("Session data:\t%s" % data.decode('utf-8'))
    #print("HTTP Status Code:\t%s\nResult code content:\t%s" % (postRecord.status_code, postRecord.content))
    print("HTTP Status Code:\t%s" % postRecord.status_code)
    return postRecord.status_code
################################################################################
# End of functions
################################################################################
f = Fernet(strKey)

strUID = f.decrypt(strUsername)
strUID = strUID.decode("utf-8")

strPass = f.decrypt(strPassword)
strPass = strPass.decode("utf-8")

connectionSP = sharepy.connect(strConnectURL,strUID,strPass)

## Get ListItemEntityTypeFullName from list
r2 = connectionSP.get(strGroupListInfoURI)
jsonReply = json.loads(r2.text)
strItemTypeName = jsonReply['d']['ListItemEntityTypeFullName']

postData = {"grant_type": "client_credentials","client_id" : strClientID,"client_secret": strClientSecret,"scope": "https://graph.microsoft.com/.default"}

r = requests.post(strGraphAuthURL, data=postData)

strJSONResponse = r.text
if len(strJSONResponse) > 5:
    jsonResponse = json.loads(strJSONResponse)

    strAccessToken = jsonResponse['access_token']

    getHeader = {"Authorization": "Bearer " + strAccessToken}

    strGraphNext = "https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period='D180')"
    r2 = requests.get(strGraphNext, headers=getHeader)
    readerGroupUsage = reader(io.StringIO(r2.text))
    listUsageReader = next(readerGroupUsage)
    if listUsageReader[3] == 'Owner Principal Name':
        listGroupUsage = next(readerGroupUsage)
        while listGroupUsage is not None:
                iGroupLastActivity = 9999
                iFoundItemID = findSPRecord(connectionSP,"Title",listGroupUsage[1])
                if iFoundItemID:
                    strItemURI = ("%s(%s)" % (strGroupListDataURI, iFoundItemID))
                    if len(listGroupUsage[4]) > 2:
                        d1 = datetime.datetime.strptime(listGroupUsage[0], "%Y-%m-%d")
                        d2 = datetime.datetime.strptime(listGroupUsage[4], "%Y-%m-%d")
                        iGroupLastActivity = abs((d2 - d1).days)
#                    print("Updating URL %s\n\t%s  with %s members is owned by %s: last activity %s days ago, %s items in Exchange." % (strItemURI, listGroupUsage[1],listGroupUsage[6],listGroupUsage[3],iGroupLastActivity,listGroupUsage[13]))
                    dictRecordPatch = {"__metadata": { "type": strItemTypeName}, "ownerUID": listGroupUsage[3], "memberCount":  listGroupUsage[6], "lastActivity": iGroupLastActivity, "externalMemberCount": listGroupUsage[7]}
                    updateRecord(connectionSP, strItemURI, dictRecordPatch)
                try:
                    listGroupUsage = next(readerGroupUsage)
                except StopIteration:
                    listGroupUsage = None
    else:
        print(listUsageReader)
else:
    print("Auth failed: %s" % strJSONResponse)


#0	'Report Refresh Date'
#1	'Group Display Name'
#2	'Is Deleted'
#3	'Owner Principal Name'
#4	'Last Activity Date'
#5	'Group Type'
#6	'Member Count'
#7	'External Member Count'
#8	'Exchange Received Email Count'
#9	'SharePoint Active File Count'
#10	'Yammer Posted Message Count'
#11	'Yammer Read Message Count'
#12	'Yammer Liked Message Count'
#13	'Exchange Mailbox Total Item Count'
#14	'Exchange Mailbox Storage Used (Byte)'
#15	'SharePoint Total File Count'
#16	'SharePoint Site Storage Used (Byte)'
#17	'Report Period'

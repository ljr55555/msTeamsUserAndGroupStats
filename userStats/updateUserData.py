import sys
from cryptography.fernet import Fernet
from base64 import b64encode, b64decode
import datetime
import requests
from requests_toolbelt.utils import dump
import json
from config import strListDataURI, strContextURL, strConnectURL, strUsername, strPassword, strListInfoURI,  strClientID, strClientSecret, strGraphAuthURL, strWebhookURL

import sharepy
#https://github.com/JonathanHolvey/sharepy

sys.path.append('../')
from key import strKey

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
# This function updates an existing record in SharePoint
# Input: s -- connection to  SharePoint REST API
# Output: integer HTTP response
################################################################################
# Before we get the statistics, determine if historic data maintenance is needed
################################################################################
def updateUserData(s, getHeaderData):
    # get all records & roll stuff
    strListContentURI = strListDataURI
    while strListContentURI:
        r1 = s.get(strListContentURI)
        jsonReply = json.loads(r1.text)
        jsonListContent = jsonReply['d']
        strListContentURI = jsonListContent.get("__next")
        print("Next URL is now %s" % strListContentURI)
        i = 0
        for listRecord in jsonListContent['results']:
            i = i + 1
            if (i % 100) is 0:
                print("On cycle %s, refreshing access token" % i)
                postData = {"grant_type": "client_credentials","client_id" : strClientID,"client_secret": strClientSecret,"scope": "https://graph.microsoft.com/.default"}

                r = requests.post(strGraphAuthURL, data=postData)

                strJSONResponse = r.text
                if len(strJSONResponse) > 5:
                    jsonResponse = json.loads(strJSONResponse)
                    strAccessToken = jsonResponse['access_token']
                    getHeaderData = {"Authorization": "Bearer " + strAccessToken}
                    print("getHeaderData is %s" % getHeaderData)

                connectionSP = sharepy.connect(strConnectURL,strUID,strPass)

            iFoundItemID = listRecord['ID']
            strItemURI = ("%s(%s)" % (strListDataURI, iFoundItemID))
            strUserUID = listRecord.get("Title")
            if strUserUID is not None:
                r2 = requests.get("https://graph.microsoft.com/v1.0/users/%s@windstream.com/?$select=displayName,Department" % strUserUID, headers=getHeaderData)
                strUserData = r2.text
                jsonUserData = json.loads(strUserData)
                if len(jsonUserData) > 1:
                    strDepartment = jsonUserData.get("department")

                    r3 = requests.get("https://graph.microsoft.com/v1.0/users/%s@windstream.com/manager" % strUserUID, headers=getHeaderData)
                    strManagerData = r3.text
                    jsonManagerData = json.loads(strManagerData)
                    print("%s in %s reports to %s" % (strUserUID, strDepartment, jsonManagerData.get("userPrincipalName")))
                    dictRecordPatch = {"__metadata": { "type": strItemTypeName}, 'department': strDepartment, 'manager': jsonManagerData.get("userPrincipalName"), 'Active': 1}
                    updateRecord(s, strItemURI, dictRecordPatch)
                else:
                    print("%s is inactive in ADFS" % strUserUID)
                    dictRecordPatch = {"__metadata": { "type": strItemTypeName}, 'Active': 0}
                    updateRecord(s, strItemURI, dictRecordPatch)

    print("Completed data maintenance -- user department and managers are updated")
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
r = connectionSP.get(strListInfoURI)
jsonReply = json.loads(r.text)
strItemTypeName = jsonReply['d']['ListItemEntityTypeFullName']

postData = {"grant_type": "client_credentials","client_id" : strClientID,"client_secret": strClientSecret,"scope": "https://graph.microsoft.com/.default"}

r = requests.post(strGraphAuthURL, data=postData)

strJSONResponse = r.text
if len(strJSONResponse) > 5:
    jsonResponse = json.loads(strJSONResponse)

    strAccessToken = jsonResponse['access_token']
    getHeader = {"Authorization": "Bearer " + strAccessToken}
    print("getHeader is %s" % getHeader)
    updateUserData(connectionSP, getHeader)

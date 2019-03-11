import sys
from cryptography.fernet import Fernet
from base64 import b64encode, b64decode
import datetime
import requests
from requests_toolbelt.utils import dump
import json
from config import strGroupListDataURI, strContextURL, strConnectURL, strUsername, strPassword, strGroupListInfoURI,  strClientID, strClientSecret, strGraphAuthURL, strWebhookURL

import sharepy
#https://github.com/JonathanHolvey/sharepy

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
r = connectionSP.get(strGroupListInfoURI)
jsonReply = json.loads(r.text)
strItemTypeName = jsonReply['d']['ListItemEntityTypeFullName']

postData = {"grant_type": "client_credentials","client_id" : strClientID,"client_secret": strClientSecret,"scope": "https://graph.microsoft.com/.default"}

r = requests.post(strGraphAuthURL, data=postData)

strJSONResponse = r.text
if len(strJSONResponse) > 5:
    jsonResponse = json.loads(strJSONResponse)

    strAccessToken = jsonResponse['access_token']

    getHeader = {"Authorization": "Bearer " + strAccessToken}

    fileWebOutput = open("./winPublicTeams.html","w")
    fileWebOutput.write("<head><title>Windstream Public Team Spaces</title>\n")
    fileWebOutput.write("<link rel=\"stylesheet\" type=\"text/css\" href=\"formatting.css\"></head>\n<body>\n")
    fileWebOutput.write("<form action=\"#\">\n")
    fileWebOutput.write("\t<fieldset>\n")
    fileWebOutput.write("\t\t<input type=\"text\" name=\"search\" value=\"\" id=\"search_input\" placeholder=\"Search\" autofocus />\n")
    fileWebOutput.write("\t</fieldset>\n")
    fileWebOutput.write("</form>\n")
    fileWebOutput.write("<table border=0 padding=1>\n")
    fileWebOutput.write("<thead><tr><th>Group Name</th><th>Description</th></tr></thead><tbody>\n")
    strGraphNext = "https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')&$top=65"
    strGroupRecords = []
    while strGraphNext is not None:
        r2 = requests.get(strGraphNext, headers=getHeader)
        jsonGroupData = json.loads(r2.text)
        strGraphNext = jsonGroupData.get("@odata.nextLink")
        listGroupRecords = jsonGroupData.get("value")
        for dictGroupRecord in listGroupRecords:
            iFoundItemID = findSPRecord(connectionSP,"GroupID", dictGroupRecord.get("id"))

            if iFoundItemID is not None:
                strItemURI = ("%s(%s)" % (strGroupListDataURI, iFoundItemID))
                #print("Found %s matching id %s, updating record on |%s| Team" % (iFoundItemID, dictGroupRecord.get("id"), dictGroupRecord.get("visibility")))
                if dictGroupRecord.get("visibility") == "Public":
                    strGroupRecords.append("<!--%s--><tr><td><a href=\"https://teams.microsoft.com/l/team/conversations/General?groupId=%s&tenantId=2567b4c1-b0ed-40f5-aee3-58d7c5f3e2b2\">%s</a></td><td>%s</td></tr>\n" % (dictGroupRecord.get("displayName"), dictGroupRecord.get("id"), dictGroupRecord.get("displayName"), dictGroupRecord.get("description")))

                dictRecordPatch = {"__metadata": { "type": strItemTypeName}, "Title": dictGroupRecord.get("displayName"), "Visibility": dictGroupRecord.get("visibility"), "CreatedOn": dictGroupRecord.get("createdDateTime"), "RenewedOn": dictGroupRecord.get("renewedDateTime"), "DeletedOn": dictGroupRecord.get("deletedDateTime")}
                updateRecord(connectionSP, strItemURI, dictRecordPatch)

            else:
                #print("No matching record found, creating ...")
                if dictGroupRecord.get("visibility") is "Public":
                    strGroupRecords.append("<!--%s--><tr><td><a href=\"https://teams.microsoft.com/l/team/conversations/General?groupId=%s&tenantId=2567b4c1-b0ed-40f5-aee3-58d7c5f3e2b2\">%s</a></td><td>%s</td></tr>\n" % (dictGroupRecord.get("displayName"), dictGroupRecord.get("id"), dictGroupRecord.get("displayName"), dictGroupRecord.get("description")))

                dictRecord = {"__metadata": { "type": strItemTypeName}, "Title": dictGroupRecord.get("displayName"), "GroupID": dictGroupRecord.get("id"), "Visibility": dictGroupRecord.get("visibility"), "CreatedOn": dictGroupRecord.get("createdDateTime"), "RenewedOn": dictGroupRecord.get("renewedDateTime"), "DeletedOn": dictGroupRecord.get("deletedDateTime")}
                writeNewRecord(connectionSP, dictRecord)

    strGroupRecords.sort(key=lambda x: x.lower())
    for strItem in strGroupRecords:
        fileWebOutput.write(strItem)
    fileWebOutput.write("</tbody></table>\n")
    fileWebOutput.write("<script src=\"jquery.js\"></script>\n")
    fileWebOutput.write("<script src=\"jquery.quicksearch.js\"></script>\n")
    fileWebOutput.write("<script>\n")
    fileWebOutput.write("\t$('input#search_input').quicksearch('table tbody tr');\n")
    fileWebOutput.write('</script>')
    fileWebOutput.write("</body>")
    fileWebOutput.close()
else:
    print("Auth failed: %s" % strJSONResponse)



# Create a MS Teams Channel in an existing Team for xMatters Flow Designer
This is a collection of custom steps that use the new Microsoft Graph API.  The ultimate aim is to create a new Channel in an existing Team in MS Teams.  It's also possible you may find use for some of these steps to help do other things.

We're going to be putting in 3 steps which each do things that help us create a channel.  These are:
1. **Authenticate** - Use some predefined tokens to authenticate on the API and get an access token that lets us cary out other actions for a short time.
1. **Get Team Info** - Find an existing Team knowing the Group and Team name and bring back the Team ID and other info.
1. **Create Channel** - Create a channel in an existing Team knowing the ID of the existing Team.

The usual use here would be to have xMatters create a new channel dedicated to the resolution of a newly discovered issue, perhaps when xMatters is initiated automatically by some monitoring tool or perhaps when a Major Incident is newly declared.  If you find other uses for this though please let us know!

**FAQ**
1. *Can we create a team dynamically?*

    Yes it's been done but we don't do it here.  If that's something you want to do please ask xMatters Support nicely :o)
1. *I don't want to do this straight into my prod Office 365, where can I test it?*

    You can't test in a Free Teams, it is missing some important elements.  You can create a Office 365 Trial and test there though.  See [Free versions of MS Teams and the Office 365 Trial Subscription](#free-versions-of-ms-teams-and-the-office-365-trial-subscription)

This is repo makes part fo the [xMatters Labs Flow Steps](https://github.com/xmatters/xMatters-Labs-Flow-Steps) parent repo.


---------

<kbd>
  <img src="https://github.com/xmatters/xMatters-Labs/raw/master/media/disclaimer.png">
</kbd>

---------

# Files - to do

* [README.md](README.md) - This file, which has all the code and setup instructions. Just about everything you need!
* [MS Teams - Authenticate.png](/media/MS%20Teams%20-%20Authenticate.png) - Logo for the Authenticate step
* [MS Teams - Get Team.png](/media/MS%20Teams%20-%20Get%20Team.png) - Logo for the Authenticate step
* [MS Teams - Create Channel.png](/media/MS%20Teams%20-%20Create%20Channel.png) - Logo for the Authenticate step

# Microsoft Teams
[Microsoft Teams](https://products.office.com/en-us/microsoft-teams/group-chat-software) is the chat tool that comes as part of the [Microsoft Office 365](https://www.office.com/) suite.  MS Teams has its own (slightly older) API but we're not making use of it here.  Instead we're using the [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/overview) for Office 365 which can do all sorts of things in the MS Office environment, including creating a Channel in an MS Teams Team.

## Free versions of MS Teams and the Office 365 Trial Subscription
This process will not work with a free Teams the like that you'll get by signing up at https://products.office.com/en-us/microsoft-teams/group-chat-software.  It doesn't work because we need the Microsoft Graph API and the Azure Console portal to setup permissions and these don't come with the free Teams.  However, you can sign up for an Office 365 Trial Subscription which will come with Teams, the Graph API and the Azure Console.

To get an Office 365 Trial follow [Phase 2 of Office 365 dev/test environment](https://docs.microsoft.com/en-us/office365/enterprise/office-365-dev-test-environment#phase-2-create-an-office-365-trial-subscription) choosing the "*lightweight Office 365 dev/test environment*" option.  You do not need to do Phase 1, 3 or 4.

When you're done you'll have a login username and password for a global admin on your new Office Environment.  The username will be an email address that ends .onmicrosoft.com.  This account can be used to login to the Teams UI and to the Azure Console as a global admin in the next sections.

# Microsoft Office 365 setup via Azure Console
These steps work by making calls on the [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/overview) for Office 365.  At the time of writing the xMatters inbuilt steps for MS Teams use the Teams API so the authentication for these steps are quite different (though they can be used in a flow along with the built in steps).  We're going to need to access the Microsoft Azure Console portal for your Office 365 environment and add an App Registration which will enable xMatters to use the Graph API to query and make changes to your Teams application inside Office 365.

## Add an App Registration
An app registration is a bit like a user account for xMatters to use to connect to the Graph API

1. Log into the Azure Console with a company account at https://portal.azure.com
    <img width=90% src="media/Create Registration.gif" >

1. In the console search for and click through to **App registrations**
1. Click **New registration**
1. Add a **Name** for this registration and click **Register** at the bottom of the page.
1. Congratulations you've created a new registration.  You'll be presented three long IDs:
 - **Application (client) ID**
 - **Directory (tenant) ID**
 - Object ID

 Make a note of the first 2 which we will refer to in later steps as the **Client ID** and the **Tenant ID**

## Add API Permissions to the App Registration
The App Registration you've just created needs to have some permissions added so that xMatters can do things with it.  (This adds the permissions we've been working with, it's possible this could be rationalised.)

<img width=90% src="media/Add API permissions.gif" >

1. In the App Registration you've created click **API permissions** on the left.
1. You'll see that the App has permissions to read user records only. Click **Add a permission**
1. Towards the bottom of the **Microsoft APIs** tab, under **Supported legacy APIs** click **Azure Active Directory Graph**
1. Click **Application permissions** on the right.
1. Find **Directory** > **Directory.ReadWrite.All** and select it.  Click **Add permissions**
1. Again click **Add a permission** to add more permissions.
1. This time at the top choose the large **Microsoft Graph** box.
1. Again click **Application permissions** on the right.
1. Find and select:
 -  **Directory** > **Directory.ReadWrite.All**
 -  **Group** > **Group.ReadWrite.All**

 Click **Add permissions**

This App registration now has all the permissions it needs but admin consent needs to be granted.  If you are an administrator then click **Grant admin consent for <Company Name>** at the bottom of the page and then confirm granting consent.  If you're not an administrator then you'll need to find a friendly admin to click the button for you.

## Create a Client Secret
A client secret is a bit like the password that xMatters will need to know to be able to use the App Registration you've created.

<img width=90% src="media/Create Client Secret.gif" >

1. Still in the App Registration you've created, click **Certificates & secrets**
1. Under **Client secrets** click **New client secret**
1. Enter a **Description** if you want.  Choose how long you want this secret to last before it expires.  Click **Add**
1. A **Client Secret** will be created and you'll be able to see the value.  Make note of this with the Client ID and Tenant ID, you'll need it later too.


# Flow Designer Steps

Here are the configurations and scripts you will require to create each of the custom steps.  Open a communication plan in xMatters and then open the flow designer for any of the forms/flows.  On the **CUSTOM** tab in right had toolbar you will find the **Create a custom step** button.  For each of these steps click that button and fill in the details as shown here.

Once you have all the custom steps created you can drag them into any flow in that communication plan and configure the step. You can use them over and over again in any flow on that com plan.

## Authenticate
Authenticate with the Microsoft Graph API and get a session token to authorise subsequent steps.

To do anything in MS Teams through the Graph API you need to first get a session token.  Send in your **Tenant ID**, **Client ID** and **Client Secret** retrieved from the [Azure portal with the above steps](#microsoft-office-365-setup-via-azure-console).  In return you will get a Session Token as an output that can be passed to later steps.

### Settings

| Field | Value |
| ----- | ----- |
| Name | MS Teams - Authenticate |
| Description | Run this step at the beginning of your flow to get an session token that can then be used for each subsequent Teams step. For more details on how to get the three inputs for this step see the instructions on the GitHub repo xm-labs-step-msteam-channels  |
| Icon | <kbd> <img src="/media/MS Teams - Authenticate.png"></kbd> |
| Include Endpoint | Yes |
| Endpoint Type |	No Authentication |
| Endpoint Label | MS Graph API |

<img src="media/Authenticate - Settings.png" >

### Inputs

| Name  | Required? | Min | Max | Help Text | Default Value | Multiline |
| ----- | ----------| --- | --- | --------- | ------------- | --------- |
| Tenant ID  | Yes | 0 | 2000 | Login to the Azure portal and create an App Registration, you'll be given the "Directory (tenant) ID" which should then be pasted into here. |  | No |
| Client ID | Yes | 0 | 2000 | Login to the Azure portal and create an App Registration, you'll be given the "Application (client) ID" which should then be pasted into here. |  | No |
| Client Secret | Yes | 0 | 2000 | Login to the Azure portal and create an App Registration, then in Certificates & secrets on that registration create a Client secret. You'll be given a complex secret which should then be pasted into here. |  | No |

<img src="media/Authenticate - Inputs.png" >

### Outputs

| Name |
| ---- |
| Session Authentication Token |

<img src="media/Authenticate - Outputs.png" >

### Script

```javascript
var sessionToken = bearerToken( input['Tenant ID'], input['Client ID'], input['Client Secret'] );

if ( sessionToken )  {
    output['Session Authentication Token'] = sessionToken;
} else {
    console.log('*** An error prevented the token from being created!');
}


/** bearerToken    (RSelby & AWatson - xMatters Consulting)
  * Authenticate and receive a token
  * Returns a token if successful, otherwise false
  */
function bearerToken(tenantId, client_id, client_secret)  {
    console.log("bearerToken starting");

    // Path is https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize
    var request = http.request({
      "endpoint": "MSTeams",
      "path": "/"+ tenantId+"/oauth2/v2.0/token",
      "method": "POST",
      "headers": {
      "Content-type": "application/x-www-form-urlencoded"
      }
    });

    var details = {
      'client_id':  client_id,   
      'grant_type': 'client_credentials',
      'client_secret': client_secret,  
      'scope': "https://graph.microsoft.com/.default"
    };

    var data = [];
    for (var property in details) {
        var encodedKey = encodeURIComponent(property);
        var encodedValue = encodeURIComponent(details[property]);
        data.push(encodedKey + "=" + encodedValue);
    }
    data = data.join("&");

    var response = request.write(data);
    var token;

    if (response.statusCode == 200) {
        json = JSON.parse(response.body);
        token = json.access_token;
        console.log("bearerToken Success. Bearer Token = " + token);
        return token;
    }

    console.log ("bearerToken Didn't work, no bearer token");  
    console.log(JSON.stringify(response));
    return false;
}
```

## Get Team Info
Gets details for a given Team specified by Team Name returning details including the Team ID.

The returned Team ID is required by the **Create Channel** step to address the Team.  Before you are able to call this step your flow will have had to have used the **Authenticate** step to get a session token which is then passed to this step as an input.

(Each Team has a Group, though not all Groups have Teams.  Any Team's ID is the same as it's Group's ID.  To find a Team you actually look for groups with Teams.  It's all a little confusing.)


### Settings

| Field | Value |
| ----- | ----- |
| Name | MS Teams - Get Team Info |
| Description | This step will find a Team and return some information about it including it's ID which can be useful for onward steps.  |
| Icon | <kbd> <img src="/media/MS Teams - Get Team.png"></kbd> |
| Include Endpoint | Yes |
| Endpoint Type | No Authentication |
| Endpoint Label | MS Graph API |

<img src="media/Get Team Info - Settings.png" >

### Inputs

| Name  | Required? | Min | Max | Help Text | Default Value | Multiline |
| ----- | ----------| --- | --- | --------- | ------------- | --------- |
| Session Authentication Token  | Yes | 0 | 2000 | You must use the MS Teams - Authenticate step earlier in the flow and input the returned session token here |  | No |
| Team Name | Yes | 0 | 2000 | The display name of the Team you're looking for |  | No |

<img src="media/Get Team Info - Inputs.png" >

### Outputs

| Name |
| ---- |
| Team ID |
| Team Description |
| Team Web URL |
| Team is Archived |

<img src="media/Get Team Info - Outputs.png" >

### Script

```javascript
//A Group isn't exactly the same thing as a Team but they are related 1 to 1 :
//
//Every team is associated with a group. The group has the same ID as the team - for example, /groups/{id}/team is the same as /teams/{id}
// https://docs.microsoft.com/en-us/graph/api/resources/team?view=graph-rest-1.0
//
//[A group] represents an Azure Active Directory (Azure AD) group, which can be an Office 365 group, or a security group.
// https://docs.microsoft.com/en-us/graph/api/resources/group?view=graph-rest-1.0


var groupID = getGroup(input['Team Name'], input['Session Authentication Token']);
if ( groupID )  {

    var teamObj = getTeam( groupID, input['Session Authentication Token']);
    if ( teamObj )  {
        console.log('Found group and team with display name "' + input['Team Name'] + '"');
        output['Team ID'] = groupID;
        if ( teamObj.description )  {
            output['Team Description'] = teamObj.description;
        } else {
            output['Team Description'] = '';
        }
        output['Team Web URL'] = teamObj.webUrl;
        output['Team is Archived'] = teamObj.isArchived;
        //Add any other outputs you need

    } else {
        console.log('Found the group with display name "' + input['Team Name'] + '" but could not find a team associated with it.');

    }

} else {
    console.log('Could not find a group with the display name "' + input['Team Name'] + '"');

}


/** getGroup  (RSelby & AWatson - xMatters Consulting)
  * Check for existence of a group
  * Return it's ID if we find it. Otherwise false.
  */
function getGroup(groupName, token)
{
    console.log("getGroup - Looking for "+groupName);

    // Gotta LIST all groups with a search filter
    var request = http.request({
    "endpoint": "MS Graph API",
    "path": "/v1.0/groups?$filter=startswith(displayName,'"+ encodeURIComponent(groupName) +"')",
    "method": "GET",
    "headers": {
      "Authorization": "Bearer " + token
    }
  });

  var response = request.write();
  //var msUserId;   //ARW don't think we need this
  var statusCode = response.statusCode;
  if ( statusCode == 200)
  {
        json = JSON.parse(response.body);
        var groupList=[];
        groupList = json.value;
        console.log("getGroup Pulled listing of "+ groupList.length +" groups");  

        for (var i = 0; i< groupList.length; i++)
        {
            var displayName = groupList[i].displayName;
            console.log("getGroup Got group called  " + displayName);
            if (displayName == groupName){
                var groupId = groupList[i].id;
                console.log("getGroup Bingo! Returning id "+ groupId);
                return  groupId;
            }
        }

        console.log("getGroup - Don't see it");
        return false;
    }
    console.log("getGroup - Bad request " + statusCode);
    return false;
}


/** getTeam  (RSelby & AWatson - xMatters Consulting)
  * Look for a Team within a Group
  * Return it as an object if we find it. Otherwise false.
  * The webUrl property is going to be very handy in the returned data.
  *
  * Example Output Object:
{
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams/$entity",
  "id": "15e8b5c8-0223-42c4-be7e-82c362e49a03",
  "displayName": "Penguin Major Incidents",
  "description": null,
  "internalId": "19:2805c6a56a0742fe9f1c81838eded38d@thread.skype",
  "webUrl": "https://teams.microsoft.com/l/team/19:2805c6a56a0742fe9f1c81838eded38d%40thread.skype/conversations?groupId=15e8b5c8-0223-42c4-be7e-82c362e49a03&tenantId=1d129432-b0c3-45f1-8f54-5c691ceed94e",
  "isArchived": false,
  "discoverySettings": {
    "showInTeamsSearchAndSuggestions": true
  },
  "memberSettings": {
    "allowCreateUpdateChannels": true,
    "allowDeleteChannels": true,
    "allowAddRemoveApps": true,
    "allowCreateUpdateRemoveTabs": true,
    "allowCreateUpdateRemoveConnectors": true
  },
  "guestSettings": {
    "allowCreateUpdateChannels": false,
    "allowDeleteChannels": false
  },
  "messagingSettings": {
    "allowUserEditMessages": true,
    "allowUserDeleteMessages": true,
    "allowOwnerDeleteMessages": true,
    "allowTeamMentions": true,
    "allowChannelMentions": true
  },
  "funSettings": {
    "allowGiphy": true,
    "giphyContentRating": "moderate",
    "allowStickersAndMemes": true,
    "allowCustomMemes": true
  }
}
  */
function getTeam( groupId, token)
{
    console.log("getTeam - Retrieving Team in group "+groupId);

    var request = http.request({
      "endpoint": "MS Graph API",
      "path": "/v1.0/teams/"+ encodeURIComponent(groupId),
      "method": "GET",
      "headers": {
      "Authorization": "Bearer " + token
     }
    });
    var response = request.write();
    var statusCode =response.statusCode;
    if (statusCode==200){
        json = JSON.parse(response.body);
        console.log('getTeam - Found it. Returning Team Object');
        return json;
    }

    console.log('getTeam - status '+statusCode+'. No Team in group ' + groupId);
    return false;
}
```


## Create Channel
Creates a new Channel, of any valid given name, within an existing Team in MS Teams.

Before you are able to call this step your flow will have had to have used the **Authenticate** step to get a session token which is then passed to this step as an input.  Most likely you will also want to have used the **Get Team Info** step in the flow to get the unique ID of the Team you want to create the channel in as that also needs to be passed to this step.  (If you want to specify the Team ID in a constant that's fine, but we prefer using a constant for the Team Name.)


### Settings

| Field | Value |
| ----- | ----- |
| Name | MS Teams - Create Channel |
| Description | Specify a channel name and this step will create the channel if it doesn't already exist and return the created channel ID. If the channel does already exist we simply return the ID of the channel.  |
| Icon | <kbd> <img src="media/MS Teams - Create Channel.png"></kbd> |
| Include Endpoint | Yes |
| Endpoint Type | No Authentication |
| Endpoint Label | MS Graph API |

<img src="media/Create Channel - Settings.png" >


### Inputs

| Name  | Required? | Min | Max | Help Text | Default Value | Multiline |
| ----- | ----------| --- | --- | --------- | ------------- | --------- |
| Session Authentication Token  | Yes | 0 | 2000 | You must use the MS Teams - Authenticate step earlier in the flow and input the returned session token here |  | No |
| Channel Name | Yes | 0 | 2000 | The name of the channel you want to create |  | No |
| Channel Description | No | 0 | 2000 | Help others find the right channel by providing a description |  | Yes |
| Team ID | Yes | 0 | 2000 | This is the ID of the Office 365 Group that is associated with the Team we want to create the Channel in. You can use the MS Teams - Get Team IDs step earlier in the flow to find this from the Team name. |  | No |

<img src="media/Create Channel - Inputs.png" >

### Outputs

| Name |
| ---- |
| Channel ID |
| Channel Web URL |
| Channel Already Existed? |

<img src="media/Create Channel - Outputs.png" >

### Script

```javascript
output['Channel Already Existed?'] = 'FALSE';

var existingChannel = getChannel(input['Channel Name'], input['Team ID'], input['Session Authentication Token']);

if ( existingChannel )  {
    console.log('The channel already exists, returning channel ID to flow');
    output['Channel Already Existed?'] = 'TRUE';
    output['Channel ID'] = existingChannel.id;
    output['Channel Web URL'] = existingChannel.webUrl;

} else {

    var newChannel = createChannel(input['Channel Name'], input['Channel Description'], input['Team ID'], input['Session Authentication Token'] );

    if ( newChannel ) {
        console.log('The channel did not exist and so was created.');
        output['Channel ID'] = newChannel.id;
        output['Channel Web URL'] = newChannel.webUrl;

    } else {
        console.log('The channel did not exist but something went wrong trying to create it.');
        output['Channel ID'] = '';
    }

}



/** createChannel  (RSelby & AWatson - xMatters Consulting)
  * Make a Channel inside a Team  
  * Returns Channel object, or false if it can't be created.
  * webUrl is one of the properties of the Channel object, which could be useful
  *
  * Returned object example:
{
  "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#teams('15e8b5c8-0223-42c4-be7e-82c362e49a03')/channels/$entity",
  "id": "19:48da5016f72f4884862e00e545b98d0f@thread.skype",
  "displayName": "My new channel",
  "description": "This is a channel description",
  "email": "",
  "webUrl": "https://teams.microsoft.com/l/channel/19%3a48da5016f72f4884862e00e545b98d0f%40thread.skype/My+new+channel?groupId=15e8b5c8-0223-42c4-be7e-82c362e49a03&tenantId=1d129432-b0c3-45f1-8f54-5c691ceed94e"
}
  */
function createChannel(channelName, channelDesc, groupId, token) {
    console.log("createChannel starting for "+channelName + " in group "+groupId);

    var request = http.request({
      "endpoint": "MS Graph API",
      "path": "/v1.0/teams/"+groupId+"/channels",
      "method": "POST",
      "headers": {
      "Authorization": "Bearer " + token
      }
    });

    // Configure Channel settings
    var data = {
     "displayName":channelName,
     "description":channelDesc
    };

    var response = request.write(data);
    var statusCode = response.statusCode;
    if ( statusCode== 201) {
        json = JSON.parse(response.body);
        console.log('createChannel - Created it. Returning Channel '+channelName);
        return json;
    }
    else     {
        console.log("createChannel statusCode is "+statusCode +". Unable to create channel");
    }
    return false;
}



/** getChannel  (RSelby & AWatson - xMatters Consulting)
  * Look for a Channel within a Group
  * Return it as an object if we find it. Otherwise false.
  *
  * Returned object example:
{
  "id": "19:48da5016f72f4884862e00e545b98d0f@thread.skype",
  "displayName": "My new channel",
  "description": "This is a channel description",
  "email": "",
  "webUrl": "https://teams.microsoft.com/l/channel/19%3a48da5016f72f4884862e00e545b98d0f%40thread.skype/My+new+channel?groupId=15e8b5c8-0223-42c4-be7e-82c362e49a03&tenantId=1d129432-b0c3-45f1-8f54-5c691ceed94e"
}
  */
function getChannel(channelName, groupId, token)
{
    console.log("getChannel - Looking for "+ channelName + " in group "+groupId);

    var request = http.request({
      "endpoint": "MS Graph API",
      "path": "/v1.0/teams/"+groupId+"/channels",
      "method": "GET",
      "headers": {
      "Authorization": "Bearer " + token
     }
    });
    var response = request.write();
    var statusCode =response.statusCode;
    var channels = [];
    if (statusCode==200){
        json = JSON.parse(response.body);
        channels = json.value;    
        channelCount = json.value.length;
        if (!channelCount){
            console.log('getChannel - No channels returned when looking for name ' + channelName + '!');
            return false;
        }

        console.log('getChannel - Found '+channelCount + " channels. Looking for "+channelName );
        for (var j=0;  j < channelCount; j++)
        {
            var channel = channels[j];
            console.log('getChannel - '+j+'. Comparing '+ channel.displayName);

            if (channel.displayName == channelName) {
                console.log('getChannel - Found!');
                return channel;
            }
        }
        console.log('getChannel - Channels returned but no '+channelName+' channel found in the list');
        return false;
    }

    console.log('getChannel - status '+statusCode+'. Not good.');
    return false;
}
```

# Create a MS Teams Channel and post to it with xMatters Flow Designer
This is a collection of custom steps that use the new Microsoft Graph API.  The ultimate aim is to create a new Channel in an existing Team in MS Teams and post a message to it.  It's also possible you may find use for some of these steps to help do other things.

We're going to be adding 5 steps which each do things that help us create a channel.  These are:
1. **Application Authenticate** - Use some predefined tokens to authenticate on the API and get an access token that lets us cary out other actions for a short time.
1. **Delegated Authenticate** - Use some predefined tokens and a username and password to authenticate on the API and get an access token that lets us cary out actions as a user for a short time.
1. **Get Team Info** - Find an existing Team knowing the Group / Team name and bring back the Team ID and other info.
1. **Create Channel** - Create a channel in an existing Team knowing the ID of the existing Team.
1. **Post to Channel** - Post a message to an existing channel knowing name of the channel and the ID of the Team it is in.

The usual use here would be to have xMatters create a new channel dedicated to the resolution of a newly discovered issue, perhaps when xMatters is initiated automatically by some monitoring tool or perhaps when a Major Incident is newly declared.  If you find other uses for this though please let us know!

**FAQ**
1. *Can we create a team dynamically?*

    Yes it's been done but we don't do it here.

1. *I don't want to do this straight into my prod Office 365, where can I test it?*

    You can't test in a Free Teams, it is missing some important elements.  You can create a Office 365 Trial and test there though.  See [Free versions of MS Teams and the Office 365 Trial Subscription](#free-versions-of-ms-teams-and-the-office-365-trial-subscription)

This is repo is part of the [xMatters Labs Flow Steps](https://github.com/xmatters/xMatters-Labs-Flow-Steps) parent repo.


---------

<kbd>
  <img src="https://github.com/xmatters/xMatters-Labs/raw/master/media/disclaimer.png">
</kbd>

---------

# Files

* [README.md](README.md) - This file, setup help and instructions. A great start!
* [MSTeamsManageChannels.zip](MSTeamsManageChannels.zip) - An export of an xMatters Workflow containing all 5 steps and an example flow to get you started. Everything you need!
* [MS Teams - Authenticate.png](/media/MS%20Teams%20-%20Authenticate.png) - Logo for the Application Authenticate step
* [MS Teams - Delegated Authenticate.png](/media/MS%20Teams%20-%20Delegated Authenticate.png) - Logo for the Delegated Authenticate step
* [MS Teams - Get Team.png](/media/MS%20Teams%20-%20Get%20Team.png) - Logo for the Get Team Info step
* [MS Teams - Create Channel.png](/media/MS%20Teams%20-%20Create%20Channel.png) - Logo for the Create Channel step
* [MS Teams - Post Message.png](/media/MS%20Teams%20-%20Post Message.png) - Logo for the Post to Channel step

# Microsoft Teams
[Microsoft Teams](https://products.office.com/en-us/microsoft-teams/group-chat-software) is the chat tool that comes as part of the [Microsoft Office 365](https://www.office.com/) suite.  MS Teams has its own (slightly older) API but we're not making use of it here.  Instead we're using the [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/overview) for Office 365 which can do all sorts of things in the MS Office environment, including creating a Channel in an MS Teams Team and post to a Channel without needing a channel web hook.

## Free versions of MS Teams and the Office 365 Trial Subscription
This process will not work with a free Teams the like you'll get by signing up at products.office.com/en-us/microsoft-teams/group-chat-software.  It doesn't work because we need the Microsoft Graph API and the Azure Console portal to setup permissions and these don't come with the free Teams.  However, you can sign up for an Office 365 Trial Subscription which will come with Teams, the Graph API and the Azure Console.

To get an Office 365 Trial follow Phase 1 of [The lightweight base configuration](https://docs.microsoft.com/en-gb/microsoft-365/enterprise/lightweight-base-configuration-microsoft-365-enterprise). You only need to do **Phase 1: Create your Office 365 E5 subscription**.

When you're done you'll have a login username and password for a global admin on your new Office Environment.  The username will be an email address that ends .onmicrosoft.com.  This account can be used to login to the [Teams WebUI](https://teams.microsoft.com/) and Desktop App as well as the Azure Console as a global admin in the next sections.

# Microsoft Office 365 setup via Azure Console
These steps work by making calls on the [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/overview) for Office 365.  At the time of writing the xMatters inbuilt steps for MS Teams use the Teams API so the authentication for these steps are quite different (though they can be used in a flow along with the built in steps).  We're going to need to access the Microsoft Azure Console portal for your Office 365 environment and add an App Registration which will enable xMatters to use the Graph API to query and make changes to your Teams application inside Office 365.

## Add an App Registration
An app registration is how xMatters connects to the Graph API.

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
1. Click **Delegated permissions** on the left.
1. Find and select:
 - **Group** > **Group.ReadWrite.All**
 - **User** > **User.Read**
1. Now back at the top click **Application permissions** on the right.
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

## Create an xMatters User

<img width=90% src="media/Add xMatters User.gif" >

# Flow Designer Steps

Now for the good stuff!  Download workflow file [MSTeamsManageChannels.zip](MSTeamsManageChannels.zip) and import it to your instance (help importing workflows).  The workflow has an example flow that uses four new custom steps.  If you want to go ahead and try it simply
1. Configure the **Delegated Authenticate** step and add the **Client ID** and **Client Secrete** you've created above along with the **Username** and **Password** of a user you want to cary out actions as.
1. Configure the **Get Team Info** step and add the Team Name of the team you're going to create new channels into.
1. To look as real as possible there is a HTTP Trigger at the beginning of this example flow that is designed to take a call from our **ServiceNow Engage with xMatters** integration.  If you have ServiceNow you can point the engage form to this trigger.  If you don't have a ServiceNow to hand you can quickly try triggering this with [Postman](https://www.postman.com/) or similar to post a packet like this to the trigger URL.  You can get the trigger URL by configuring the HTTP Trigger step.  Authenticate as your own user using basic.  Make sure you use the Header `Content-Type: application/json`
```
{
	"recipients": [
	    {"targetName": "YOUR XMATTERS USERNAME"}
	],
	"properties" : {
		"number":"INC12121234",
		"task_message": "Need some help here",
		"impact": "1 - Critical",
		"task_initiator_display_name": "Incident Manager"
	}
}
```

By importing this workflow the custom steps will be added you your instance in state Development and unshared.  In this state you personally will be able to use them in this and other workflows.  When you're happy with them you should [set them to a Deployed state](TBC) and if you want other people to be able to use them you should [share the steps](TBC) too.

## Authentication Steps

Each flow you execute with Teams actions needs to first authenticate with the Microsoft Online Authentication API and get a token.  The following steps need to use an **endpoint** with the base URL `https://login.microsoftonline.com` [Help adding endpoints](https://help.xmatters.com/ondemand/xmodwelcome/flowdesigner/components.htm#Endpoints)

Both steps will output an authentication token on success, you can then use that token as many times as you like.  However, some steps need different kinds of authentication tokens.  There are two kinds of authenticated tokens available, Application Authentication and Delegated Authentication.

### Application Authenticate
<img width=50px src="media/MS Teams - Authenticate.png" >

Use the Tenant ID, Client ID and Client Secret you've created above to obtain a token that enables you to take anonymous actions in Teams.  Creating and querying channels need simple permissions, you don't need to pretend to be anyone to do it.  When it comes to posting to a channel you need to be someone and this isn't going to work.

This token will work with the following steps:
- Get Team Info
- Create Channel

### Delegated Authenticate
<img width=50px src="media/MS Teams - Delegated Authenticate.png" >

Use the Client ID and Client Secrete you've created above along with the username and password of a user you want to cary out actions as.  A token from this step will allow you to do anything that user is aloud to do.  There are some things, like posting to channels, that require you to do them as a user - so you'll need to use this authentication to do them.

This token will work with the following steps where the delegated user used has permission to cary out the requested action:
- Get Team Info
- Create Channel
- Post to Channel

## Graph API Steps for MS Teams

Once you're authenticated with the steps above you can use the token you've got to cary out actions in Teams with these steps.  The following steps need to use an **endpoint** with the base URL `https://graph.microsoft.com` [Help adding endpoints](https://help.xmatters.com/ondemand/xmodwelcome/flowdesigner/components.htm#Endpoints)

### Get Team Info
<img width=50px src="media/MS Teams - Get Team.png" >

Look up a team by Team Name and get back some details.  This is most useful to get the Team ID which is required for most further steps.

Inputs:
- `Session Authentication Token` Application Authenticated or Delegated Authenticate
- `Team Name` The name of the team as you read it in MS Teams

Outputs:
- `Team ID`
- `Team Description`
- `Team Web URL`
- `Team is Archived`  (is the team archived)


### Create Channel
<img width=50px src="media/MS Teams - Create Channel.png" >

Create a new channel in a team. If the channel already exists the step will complete successfully and return all the normal detail about the channel. In this case `Channel Already Existed?` will show as `TRUE`. If the channel could not be created for any other reason the step will fail.

Inputs:
- `Session Authentication Token` Application Authenticated or Delegated Authenticated
- `Channel Name` Name the new channel
- `Channel Description` Optional
- `Team ID` The internal ID of the team to create the channel in

Outputs:
- `Channel ID`  Internal ID of the created channel
- `Channel Web URL`
- `Channel Already Existed?`  Did the channel already exist


### Post to Channel
<img width=50px src="media/MS Teams - Post Message.png" >

Post a message to a channel. You can use some HTML in your message.  For instance, use \<br/> to get a new line within the message and surround text with \<b> and \<\b> to make it bold.

Inputs:
- `Delegated Session Authentication Token` Must be from the Delegated Authenticate step
- `Team ID` The internal ID of the team that has the channel we're posting to
- `Channel Name` The name of the Channel we want to post a message to
- `Message` The message you want to post to the channel

Outputs:  None yet.

## Other steps in the flow

The steps above are all the MS Teams steps in this repo.  However, we do have two more steps in this workflow which are borrowed from other repos and are worthy of their own mention.  These are only here to help illustrate creating the Teams channel in a real environment.

### ServiceNow Engage with xMatters Trigger

From: ServiceNow Engage with xMatters in Flow Designer (To be created)

We touched on this briefly when triggering the flow.  We have to have a beginning to this and typically it's going to be call in to xMatters from another tool over an HTTP trigger.  Here we're using ServiceNow Engage with xMatters as an example to demonstrate a Major Incident callout process.  The Jira Engage with xMatters would certainly be a straight swap.  But perhaps an automatic callout from some monitoring system would also apply for you?  Certainly hooking this up to an event status change to create a new channel when a new event is sent could be useful too?

### Create Advanced xMatters Event

From: [Create Advanced xMatters Event](https://github.com/xmatters/xm-labs-step-xmatters-advanced-event)

When you import this workflow you'll see that we're going to the built in `xMatters Create Event` step.  That works fine and we get an event notification as defined by the form.  However, it's kind of nice to use the URL redirect feature on the Accept response to push people right into the channel you've created.  However, the URL for the newly created channel isn't known until we create it so to do this we need to modify the response options on the event to include the channel URL in the response redirect as we create the event.  Unfortunately the built in step doesn't support this yet.  No need to worry though as we can do it with the Create Advanced xMatters Event step.

If you'd like to use this step it's actually configured and is nearly ready to go.
<img width=90% src="media/Advanced xMatters Event IB URL.gif" >
1. Briefly come out of the flow canvas and go to the older **Integration Builder** tab still within your workflow.
1. In **Inbound Integrations** find `New event with modified responses` and click into it.
1. At the bottom of the page copy the **Trigger** URL.
1. Now back in your flow, configure the `Create Advanced xMatters Event` step.
1. On the **Endpoint** tab click **Edit Endpoints**
1. Find the endpoint `New event Inbound IB` and paste what you have into the **Base URL** of this endpoint replacing everything that's there.
1. Click **Save**
1. Finaly, back on the canvas, move the final link off the `xMatters Create Event` step and onto `Create Advanced xMatters Event` and click **Save**.

Now when you respond **Accept** to the notification via the App or Email you'll be given the option of going straight into the new channel.

# Get Flowing!
Don't be content with the example flow, get creative and put these steps in all the flows you have when you want to interact with MS Teams.  Remember that when you use these steps in other workflows you will need to add both the authentication and graph endpoints to those workflows. What can you create!?

**Note:**
Please consider using com plan constants for any fixed values that you want to put in the steps, in particular the authentication IDs, Codes, usernames and passwords.   If you want to use them again in other flows it makes it so much easier to quickly change them all at the same time if you have used constants.

There's lots of help on [Flow Designer on the xMatters help pages](https://help.xmatters.com/ondemand/xmodwelcome/flowdesigner/laying-out-flow-designer.htm).

# What more do you want?
This is a community project to make our ecosystem better.  We value your input, and where possible you active participation.  We'd like to add:
- Scrape the channel history so it can be posted to another system
- Archive / Delete a channel
- Create a Team

What else should be added?  Join the discussion on the [xMatters Community](https://support.xmatters.com/hc/en-us/community/topics/200330486-Integrations-Toolchains).

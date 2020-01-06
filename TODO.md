Delegated#Teams functionality required

##Does channel exist
I think I have most of this in the create process,  need to verify

##Create channel
Done

##Add bot to channel?
I'm not sure we actually need this.  There is an add to channel though if needed

##Post to channel
Looks like this could be the API we need
https://docs.microsoft.com/en-us/graph/teams-proactive-messaging

Not really sure what a Proactive message is though.   Looks like a proactive message is direct to a user, not to a channel


Ah - or this looks better
https://docs.microsoft.com/en-us/graph/api/channel-post-messages?view=graph-rest-beta&tabs=http    (Beta)

That works but you have to use a token that is associated with the user you're going to post as,   to get that you have to add
Group.ReadWrite.All as a Delegated permission (as well as Application) on the API permissions.  Then you can use https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth-ropc through login.microsoftonline.com with the users name and password

##Retrieve channel history
Probably this  https://docs.microsoft.com/en-us/graph/api/channel-list-messages?view=graph-rest-beta&tabs=http  (Beta)

This is a protected API,  you use it with delegated permissions but if you want to use application permissions you need to get approval from Microsoft https://docs.microsoft.com/en-us/graph/teams-protected-apis

Then you can get reply data for each message ID with https://docs.microsoft.com/en-us/graph/api/channel-list-messagereplies?view=graph-rest-beta&tabs=http






There's also getting the message delta https://docs.microsoft.com/en-us/graph/api/chatmessage-delta?view=graph-rest-beta&tabs=http  (Beta)


##Archive Channel?

What are you supposed to do with a channel in Teams,  is there an Archive facility?
No, you can't archive a channel, only delete it or rename it.

There's a delete channel call  https://docs.microsoft.com/en-us/graph/api/channel-delete?view=graph-rest-beta&tabs=http  (Beta with issues)





#Interesting ideas

Create tabs in the channel that describe the incident or things to do.

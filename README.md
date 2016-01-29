# Visual Studio Team Services App for Zendesk

Unite your customer support and development teams. Quickly create or link work items to tickets, enable efficient two-way communication, and stop using email to check status.

### Create work items for your engineers right from Zendesk

With the Visual Studio Team Services app for Zendesk, users in Zendesk can quickly create a new work item from a Zendesk ticket.

![img](https://i3-vso.sec.s-msft.com/dynimg/IC729561.png)

### Get instant access to the status of linked work items

Give your customer support team easy access to the information they need. See details about work items linked to a Zendesk ticket, including curr

![img](https://ms-vsts.gallery.vsassets.io/_apis/public/gallery/publisher/ms-vsts/extension/services-zendesk/latest/assetbyname/images/zendesk-linked.png)

## How to install and setup

See [full instructions](https://www.visualstudio.com/get-started/zendesk-and-vso-vs)

### Install the app to Zendesk

1. Download the latest .zip release from this CodePlex project.
1. Sign up for a Zendesk account here: [url:https://www.zendesk.com/register]
1. Once signed in, click the settings icon (gear) in the bottom left hand portion of the screen.
1. Under *Apps* click Manage.
1. Click *Upload App*.
1. Give the app a name.
1. Browse to the location you saved the .zip release and select it.
1. Provide your Visual Studio Team Services name and decide on a work item tag for Zendesk.

### Send updates from Visual Studio Team Services to Zendesk

1. Open the admin page for the team project in Visual Studio Online
2. On the *Service Hooks* tab, run the subscription wizard
3. Select Zendesk from the subscription wizard
4. Pick and configure the Visual Studio Online event which will post to Zendesk
5. Tell Zendesk what to do when the event occurs
6. Test the service hook subscription and finish the wizard


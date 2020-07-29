# Handling authentication easily in custom Teams tabs using Microsoft Graph Toolkit 

If you are looking for ways to build easy authentication in [Microsoft Teams custom tab development](https://cda.ms/1vt), [Microsoft Graph Toolkit](https://cda.ms/1vw) [Login](https://cda.ms/1vx) component provides a button and flyout control to facilitate Microsoft identity platform authentication with couple of lines of code.

## How to build a tab with straightforward authentication flow?

1. Enable [Microsoft Teams Toolkit](https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension) extension for Visual Studio Code
1. Build Microsoft Teams tab
1. Implement [Microsoft Graph Toolkit](https://cda.ms/1tV):
    * [Login component](https://cda.ms/1tX): login button to authenticate a user with the Microsoft Identity platform
    * [Teams provider](https://cda.ms/1tY): Microsoft Teams tab to facilitate authentication
    * [Agenda component](https://cda.ms/1tZ): displays events in a user or group's calendar
1. Setup [ngrok](https://ngrok.com/docs#getting-started-authtoken) for tunneling
1. Register your app in Azure Active Directory    
1. Import your app manifest to Microsoft Teams App Studio for testing    
    
## Enable [Microsoft Teams Toolkit](https://marketplace.visualstudio.com/items?itemName=TeamsDevApp.ms-teams-vscode-extension) extension for Visual Studio Code

Install Microsoft Teams Toolkit from the **Extensions** tab on the left side bar in Visual Studio Code. For more information, [Microsoft Teams Toolkit: Setup and Overview](https://quickbites.dev/2020/06/25/microsoft-teams-toolkit-setup/).

   ![Microsoft Teams Toolkit Extension for Visual Studio Code](/Images/01.png)

## Build Microsoft Teams tab

1. Select Microsoft Teams icon on the left side bar in Visual Studio Code and **Sign in**.

   ![Microsoft Teams Toolkit Extension for Visual Studio Code](/Images/02.png)
   
1. Select **Create a new Teams app**, 
   * Give a name for the app 
   * Choose **Tab** for the capability
   * Select **Next**
   
   ![Microsoft Teams Toolkit Extension for Visual Studio Code](/Images/03.png)
   
   * Select **Personal tab**
   * Select **Finish**
   
   ![Microsoft Teams Toolkit Extension for Visual Studio Code](/Images/04.PNG)
   
   * Open Terminal and run:
   
      `npm install`
   
      `npm start`

## Implement [Microsoft Graph Toolkit](https://cda.ms/1tV)

Add a new file under **src** folder and name as **Auth.js**.

```
import React from 'react';
import ReactDOM from 'react-dom';
import './index.css';
import App from './components/App';
import { Provider, themes } from '@fluentui/react-northstar' //https://fluentsite.z22.web.core.windows.net/quick-start

ReactDOM.render(
    <Provider theme={themes.teams}>
        <App />
    </Provider>, document.getElementById('auth')
);
```

Add a new file under **public** folder and name as **auth.html**. `CTRL+SPACE` for adding HTML Sample. Add below code in `<body></body>`

```
<div id="auth"></div>
<script src="https://unpkg.com/@microsoft/teams-js/dist/MicrosoftTeams.min.js" crossorigin="anonymous"></script>
<script src="https://unpkg.com/@microsoft/mgt/dist/bundle/mgt-loader.js"></script>
<script>
  mgt.TeamsProvider.handleAuth();
</script>
```

Add below code in **index.html** `<body></body>`

```
 <script src="https://unpkg.com/@microsoft/teams-js/dist/MicrosoftTeams.min.js" crossorigin="anonymous"></script>
 <script src="https://unpkg.com/@microsoft/mgt/dist/bundle/mgt-loader.js"></script>
 <mgt-teams-provider client-id="YOUR-CLIENT-ID" auth-popup-url="YOUR-NGROK-URL/auth.html"></mgt-teams-provider> 
 <mgt-login></mgt-login>
 <mgt-agenda></mgt-agenda>
```  

## Setup [ngrok](https://ngrok.com/) for tunneling

1. Open Terminal and run the solution. Default Teams tab will be running `https://localhost:3000`:
   
   `npm start`
   
1. Go to [ngrok](https://ngrok.com/) website and login.

1. Complete the setup and installation guide. 

1. Save Authtoken in the default configuration file `C:\Users\[user name]\.ngrok`:
   
   `ngrok authtoken <YOUR_AUTHTOKEN>`
   
1. Run below script to create ngrok tunnel for `https://localhost:3000`:

   `ngrok http https://localhost:3000`
   
   ![Ngrok Setup](/Images/06.PNG) 
   
   ![Ngrok Setup](/Images/05.PNG) 

1. Go to your project **.publish > Development.env**, replace `baseUrl0` with ngrok url  `https://xxxxxxxxxxxx.ngrok.io`.

   ![Ngrok Setup](/Images/07.PNG) 
   
1. Go to your project **public > index.html**, replace `YOUR-NGROK-URL` with ngrok url  `https://xxxxxxxxxxxx.ngrok.io` in **mgt-teams-provider > auth-popup-url**.

   ![Ngrok Setup](/Images/08.PNG) 
   
## Register your app in Azure Active Directory

1. Go to [Azure Portal](https://portal.azure.com), then **Azure Active Directory > App Registration** and select **New Registration**.

   ![Ngrok Setup](/Images/09.PNG) 

1. Fill the details to register an app:
   * give a name to your application
   * select **Accounts in any organizational directory** as an access level
   * place **auth-popup-url** as the redirect url `https://xxxxxxxxxxxx.ngrok.io/auth.html`
   * select **Register**
   
   ![Ngrok Setup](/Images/10.PNG) 

1. Go to **Authentication** tab and enable **Implicit grant** by selecting `Access tokens` and `ID tokens`

   ![Ngrok Setup](/Images/11.PNG)
   
1. Go to **API permissions** tab, select **Add a permission > Microsoft Graph > Delegated permissions** and add `Calendar.Read`, `Calendar.ReadWrite`.
1. Then, select **Grant admin consent**.

   ![Ngrok Setup](/Images/19.PNG)
   
1. Go to **Overview** tab and copy **Application (client) ID**
1. Then go to your project **public > index.html**, replace `YOUR-CLIENT-ID` with `Application (client) ID` in **mgt-teams-provider > auth-popup-url**.

   ![Ngrok Setup](/Images/12.PNG)
   
   ![Ngrok Setup](/Images/13.png)
   
## Import your app manifest to Microsoft Teams App Studio for testing

1. Go to [Microsoft Teams](https://teams.microsoft.com), open **App Studio > Manifest Editor** and select **Import an existing app**.
   
   ![Ngrok Setup](/Images/14.png)
   
1. Select **Development.zip** under **your project folder > .publish**.

   ![Ngrok Setup](/Images/15.png)

1. Scroll down and select **Test and distribute**, then click **Install** and **Add** your app.

   ![Ngrok Setup](/Images/16.png)
   
   ![Ngrok Setup](/Images/17.png)
   
1. Click on **Sign in** for the authentication and give consent to AAD registered app you created.

   ![Ngrok Setup](/Images/18.png)
   
1. Your profile information `e-mail`, `name` and calendar should appear in your tab after the successful authentication.

   ![Ngrok Setup](/Images/20.PNG)
   

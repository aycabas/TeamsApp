# Using Microsoft Graph Toolkit in Microsoft Teams Tabs

## Content:

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



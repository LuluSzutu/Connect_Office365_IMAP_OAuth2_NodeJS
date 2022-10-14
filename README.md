# How To Connect Office365 with IMAP using OAuth Authentication in NodeJS

Microsoft is ending support for legacy authentication and deactivating it start from October 2022. This means connecting to the IMAP server by using just <b>UserName</b> and <b>Password</b> is no longer working with office365. For NodeJS users, the following code, <span style="color:red"><b>will not work any more</b></span>:

```dash
const Imap = require('node-imap'), inspect = require('util').inspect;
var imap = new Imap({
  user: '******@outlook.com', // or '******@<your-domain>.com'  
  password: '************',
  host: 'outlook.office365.com',
  port: 993,
  tls: true
});
```

This document explain how to use AUTH2 authentication method to obtain a token that can be used to access office 365 with IMAP protocal, and how to use IMAP module to read the emails. 

## 1. Set up client credential grant flow to authenticate IMAP. 

The official instructions offered by microsoft can be found in here: [https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth](https://learn.microsoft.com/en-us/exchange/client-developer/legacy-protocols/how-to-authenticate-an-imap-pop-smtp-application-by-using-oauth)

To be able to setup the client credential, you must have permission to manage applications, and login to Azure Active Directory (Azure AD). This means you must have an <b>Administrator</b> or <b>Application Developer</b> roles for Microsoft Azure Portal. 

### 1.1 Register application and grant API permissions.
1. Sign in to the Azure portal.
2. Select Azure Active Directory. 
3. Under Manage, select App registrations > New registration.
4. When registration finisheds, the Azure portal displays the app registration's overview pane. You need to copy the <b>Application (client) ID</b> and the <b>Directory (tenant)ID</b>, because you will need them for your NodeJS code as [clientId] and [tenantId]. 
5. Under Manage, select  Certificate & secrets > Client secrets > New client secret 
6. Add a new client secret with Description and Expire date. Reminber the Description, you will need it later as [appname]. 
7. After generating the new client secret, you will see the Value and SecretID, copy the <b>Value</b>, you will need this for your NodeJS code as [clientSecret]. 
8. Under Manage, select API permissions > Add a permission, select APIs my organization users, searching for office 365 Exchange Online, then choose for Application permissions, then searching for IMAP, when find it, check the IMAP.AccessAsApp, and then Add permission. 
9. Under API permissions pa, click the line 'Grant admin consent for [username]'

### 1.2 Register service principals in Exchange
1. On a PC computer, open powershell and run as administrator. 
2. Type the following commands:

```dash
Install-Module -Name ExchangeOnlineManagement -allowprerelease
Install-Module -Name AxureAD
Install-Moduel Microsoft.Graph -allowprerelese
Import-module ExchangeOnlineManagement
Connect-ExchangeOnline -Organization [tenantId] // see step 4 on previous session. 
Import-module AzureADPreview
$MyApp = Get-AzureADServiePrincipal -SearchString [appname] // see step 6 on previous session. 
New-ServicePrincipal -AppId $MyApp.AppId -ServiceID $MyApp.objectId -DisplayName "Service Principal for IMAP APP"
Add-MailboxPermission -Identify "your-email-address" -User $MyApp.objectId -AccessRight FullAccess
```

## 2. Read emails via IMAP module using Oauth2 and NodeJS

### 2.1 Get accesstoken from outlook 365

```dash
const msal = require("@azure/msal-node");

 const msalConfig = {
    auth: {
      clientId: [clientId], // see step 4 on previous session.
      authority:
        'https://login.microsoftonline.com/[tenantId]', // see step 4 on previous session.
      clientSecret: [clientSecret], 
      redirectUri: 'http://localhost:***', // the redirect port name, this can be omit. 
      grantType: 'client_credentials',
    },
    cache: {
        cacheLocation: "sessionStorage", // This configures where your cache will be stored
        storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
    },
    system: {	
        loggerOptions: {	
            loggerCallback: (level, message, containsPii) => {	
                if (containsPii) {		
                    return;		
                }		
                switch (level) {		
                    case msal.LogLevel.Error:		
                        console.error(message);		
                        return;		
                    case msal.LogLevel.Info:		
                        console.info(message);		
                        return;		
                    case msal.LogLevel.Verbose:		
                        console.debug(message);		
                        return;		
                    case msal.LogLevel.Warning:		
                        console.warn(message);		
                        return;		
                }	
            }	
        }	
    }
  };
  const cca = new msal.ConfidentialClientApplication(msalConfig);
	
  const tokenRequest = {
    scopes: ['https://outlook.office365.com/.default']  
  };

  const tokenResponse = await cca.acquireTokenByClientCredential(tokenRequest);
```
  
### 2.2 Generate Oauth2 token
  
```dash
  var xoauth2_format = '';
  xoauth2_format += 'user=' + '[your-email-address]';
  xoauth2_format += '\1';
  xoauth2_format += 'auth=Bearer ' + tokenResponse.accessToken;
  xoauth2_format += '\1\1';
  var oauth2 = btoa(xoauth2_format); //base64 encodes
  console.log(oauth2);  
```
  
To check the Oauth2 token works or not, you can use openssl. Go to a LINUX or MAC computer, open the terminal, and type:

```dash
openssl s_client -connect outlook.office365.com:993 -crlf
```

Then copy the generated oauth2 token and paste in the following code

```dash
1 AUTHENTICATE XOAUTH2 <oauth2>
```
If the authentication success, you will see the following message:

```dash
[connection begins]
C: C01 CAPABILITY
S: * CAPABILITY … AUTH=XOAUTH2
S: C01 OK Completed
C: A01 AUTHENTICATE XOAUTH2 dXNlcj1zb21ldXNlckBleGFtcGxlLmNvbQFhdXRoPUJlYXJlciB5YTI5LnZGOWRmdDRxbVRjMk52YjNSbGNrQmhkSFJoZG1semRHRXVZMjl0Q2cBAQ==
S: A01 OK AUTHENTICATE completed.
```

### 2.3 Read email using node-imap module with oauth2

```dash
const Imap = require('node-imap'), inspect = require('util').inspect;

function openInbox(cb) {
  imap.openBox('INBOX', false, cb);
}

imap.once('ready', function() {
  openInbox(function(err, box) {
    if (err) throw {status: 500, message: 'cannot connect to email server!'};
    
    // search messages 
    imap.search(['UNSEEN', ['Subject', '*********']], function(err,results) {
      if (err) {
        throw  {status: 500, message: 'cannot read email messages!'};
      }
       
      var output = [];        
      if (results.length === 0) {
        console.log('No messages founded. Nothing to fetch');          
        res.render('email', {parse: output, logging: dynMessage, project: projectname});
        return imap.end();
      }
           
      var f = imap.fetch(results, { 
        bodies: ['HEADER.FIELDS (FROM SUBJECT DATE)','TEXT'],
        struct: true,
        markSeen: false
      }); 

      f.on('message', function(msg, seqno) {
        console.log('Message #%d', seqno);
      });
      
      f.once('error', function(err) {
        console.log('Fetch error: ' + err);
        res.send("Email Fetching Error!");
      });
    
      f.once('end', function() {
        console.log('Done fetching all messages!');
        res.render('email', {parse: output,logging: dynMessage, project: projectname});
        imap.end();
      });
    });
    
    imap.once('error', function(err) {
    console.log('Imap error: ' + err);
    if(err.message.indexOf('Timed out') > -1) {
      throw {status: 500, message: 'Timed out while authenticating with the email server. Please try again.'};
    }
  });
});

imap.once('end', function() {
  console.log('Connection ended');
  imap.end();
});

imap.connect();

```


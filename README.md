<a name="top"></a>

# OnedriveApp

[![MIT License](http://img.shields.io/badge/license-MIT-blue.svg?style=flat)](LICENCE)

This is a library to use [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/overview) with Google Apps Script. OneDrive and Email can be managed using this library.

## Feature

### [Drive of Microsoft Graph API](#drive)

This library can carry out following functions using OneDrive APIs.

1. [Retrieve access token and refresh token using client_id and client_secret](#authprocess)
1. [Retrieve file list on OneDrive.](#retrievefilelist)
1. [Delete files and folders on OneDrive.](#deletefilesandfolders)
1. [Create folder on OneDrive.](#createfolder)
1. [Download files from OneDrive to Google Drive.](#downloadfiles)
1. [Upload files from Google Drive to OneDrive.](#uploadfiles)

By updating at October 5, 2021, OnedriveApp can get and send Emails using Microsoft Graph API.

### [Utilities](#utilities)

1. [Get access token](#getaccesstoken)

### [Email of Microsoft Graph API](#email)

1. [Get Email message list](#getemaillist)
1. [Get Email messages](#getemails)
1. [Send Email messages](#sendemails)
1. [Reply Email messages](#replyemails)
1. [Forward Email messages](#forwardemails)
1. [Get Email folders](#getemailfolders)
1. [Delete Email messages](#deleteemails)

## Demo

![](images/demo.gif)

In this demonstration, it creates a folder with the name of "SampleFolder" on OneDrive, and then a spreadsheet file is uploaded to the created folder. The spreadsheet is converted to excel file and uploaded. The scripts which was used here is as follows.

```javascript
function createFolder() {
  OnedriveApp.init(PropertiesService.getScriptProperties()).creatFolder(
    "SampleFolder"
  );
}

function uploadFile() {
  var id = DriveApp.getFilesByName("samplespreadsheet").next().getId();
  OnedriveApp.init(PropertiesService.getScriptProperties()).uploadFile(
    id,
    "/SampleFolder/"
  );
}
```

## How to install

- Open Script Editor. And please operate follows by click.
- -> Resource
- -> Library
- -> Input Script ID to text box. Script ID is **`1wfoCE1mCQpGQZZ9CrWFY_NvA9iRxkNbxN_qTGSBkRkmn8I2eguLVwfZs`**.
- -> Add library
- -> Please select latest version
- -> Developer mode ON (If you don't want to use latest version, please select others.)
- -> Identifier is "**`OnedriveApp`**". This is set under the default.

[If you want to read about Libraries, please check this.](https://developers.google.com/apps-script/guide_libraries).

- The method of `downloadFile()` and `uploadFile()` use Drive API v3. But, don't worry. Recently, I confirmed that users can use Drive API by only [the authorization for Google Services](https://developers.google.com/apps-script/guides/services/authorization). Users are not necessary to enable Drive API on Google API console. By the authorization for Google Services, Drive API is enabled automatically.

<a name="authprocess"></a>

# Retrieve access token and refresh token for using OneDrive

**Before you use this library, at first, please carry out as follows.**

## 1. OneDrive side

1. Log in to [Microsoft Azure portal](https://portal.azure.com/).
2. Search "Azure Active Directory" at the top of text input box. And open "Azure Active Directory".
3. Click "App registrations" at the left side bar.
   - In my environment, when I used Chrome as the browser, no response occurred. So in that case, I used Microsoft Edge.
4. Click "New registration"
   1. app name: "sample app name"
   2. Supported account types: "Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)"
   3. Redirect URI (optional): Web
      - URL: here, please do the blank.
   4. Click "Register"
5. Copy **"Application (client) ID"**.
6. Click "Certificates & secrets" at the left side bar.
   1. Click "New client secrets".
   2. After input the description and select "expire", click "Add" button.
   3. Copy **the created secret value**.

By above operation, the preparation is done.

[Ref](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)

## 2. Google side

Please copy and paste following script (`doGet(e)`) on the script editor installed the library, and import your **"Application (client) ID"** and **the created secret value** to `client_id` and `client_secret` in the script.

```javascript
function doGet(e) {
  var prop = PropertiesService.getScriptProperties();
  OnedriveApp.setProp(
    prop,
    "### client id ###", // <--- client_id
    "### client secret ###" // <--- client_secret
  );
  return OnedriveApp.getAccesstoken(prop, e);
}
```

Then, please do the following flow at the script editor.

- On the Script Editor
  - File
  - -> Manage Versions
  - -> Save New Version
  - Publish
  - -> Deploy as Web App
  - -> At Execute the app as, select **"your account"**
  - -> At Who has access to the app, select **"Only myself"**
  - -> Click "Deploy"
  - -> Click **"latest code"**. **At First, Please Do This!**
    - By this click, it launches the authorization process. **The refresh token was automatically saved.**

## 3. OneDrive side

When you click **"latest code"**, new tab on your browser is launched and you can see `Please push this button after set redirect_uri to`https://script.google.com/macros/s/#####/usercallback` at your application.`.

- Please back to [Microsoft Azure portal](https://portal.azure.com/) and open "Azure Active Directory".
  1. Click app you created.
  2. Click "Redirect URIs". You can see this at the right side of "Display name".
  3. Please paste `https://script.google.com/macros/s/#####/usercallback` to "Redirect URIs" with "Web" of the type.
  4. Click "Save" button. You can see this at the top.

## 4. Google side

1. Back to the page with `Get access token` button.
1. Click the button.
1. Authorize.
1. You can see `Retrieving access token and refresh token was succeeded!`. If that is not displayed, please confirm your client_id and client_secret again.
1. Access token and refresh token are shown. And they are automatically saved to your PropertiesService. **You can use OnedriveApp from now.**

This process can be seen at following demonstration.

![](images/demo_auth.gif)

**Please run this process only one time on the script editor installed this library.** By only one time running this, you can use all of this library. After run this process, you can undeploy web apps.

If your OneDrive Application was modified, please run this again.

Or, if you can retrieve refresh token by other script, please check [here](https://gist.github.com/tanaikech/d9674f0ead7e3320c5e3184f5d1b05cc).

# Usage

Also, you can see "[Known issues with Microsoft Graph](https://docs.microsoft.com/en-us/graph/known-issues)".

<a name="drive"></a>

## Drive of Microsoft Graph API

These methods manage OneDrive.

<a name="retrievefilelist"></a>

### 1. Retrieve file list on OneDrive

```javascript
var prop = PropertiesService.getScriptProperties();
var odapp = OnedriveApp.init(prop);
var res = odapp.getFilelist("### folder name ###");
```

Filenames and file IDs are returned.

If `"### folder name ###"` is not inputted (`var res = odapp.getFilelist()`), files and folders on the root directory are retrieved.

<a name="deletefilesandfolders"></a>

### 2. Delete files and folders on OneDrive.

When a file is deleted,

```javascript
var prop = PropertiesService.getScriptProperties();
var odapp = OnedriveApp.init(prop);
var res = odapp.deleteItemByName("### filename ###");
```

When a folder is deleted,

```javascript
var prop = PropertiesService.getScriptProperties();
var odapp = OnedriveApp.init(prop);
var res = odapp.deleteItemByName("/### folder name ###/");
```

In the case of folder, please enclose it in `/`.

If you want to delete files and folders using item ID, please use as follows.

```javascript
var prop = PropertiesService.getScriptProperties();
var odapp = OnedriveApp.init(prop);
var res = odapp.deleteItemById("### item ID ###");
```

#### Note :

**If you delete a folder, the files in the folder are also deleted. Please be careful about this.**

<a name="createfolder"></a>

### 3. Create folder on OneDrive.

```javascript
var prop = PropertiesService.getScriptProperties();
var odapp = OnedriveApp.init(prop);
var res = odapp.creatFolder("### foldername ###", "/### path ###/");
```

Created folder name and ID are returned. If you want to create a folder of `newfolder` in the folder of `/root/sample1/sample2/`, please use as follows. If there is no folder `sample2` on your OneDrive, following script creates `sample2` and `newfolder`, simultaneously.

```javascript
var prop = PropertiesService.getScriptProperties();
var odapp = OnedriveApp.init(prop);
var res = odapp.creatFolder("newfolder", "/sample1/sample2/");
```

<a name="downloadfiles"></a>

### 4. Download files from OneDrive to Google Drive.

```javascript
var prop = PropertiesService.getScriptProperties();
var odapp = OnedriveApp.init(prop);
var res = odapp.downloadFile("### file with path ###", convert from Microsoft to Google (true or false), "### Folder ID on Google Drive ###");
```

Downloaded file name and ID on Google Drive are returned. When a file is downloaded from OneDrive to Google Drive, if the file is Microsoft Office Docs, you can select whether the file is converted to Google Docs. If you want to convert, you can use following sample.

```javascript
var prop = PropertiesService.getScriptProperties();
var odapp = OnedriveApp.init(prop);
var res = odapp.downloadFile(
  "/SampleFolder/sample.xlsx",
  true,
  "### Folder ID on Google Drive ###"
);
```

In this case, Excel file is converted to Google Spreadsheet, and imported to the folder ID. If you don't want to convert, you can use following sample. If the folder ID is not set, the file is created to root directory on your Google Drive.

```javascript
var prop = PropertiesService.getScriptProperties();
var odapp = OnedriveApp.init(prop);
var res = odapp.downloadFile(
  "/SampleFolder/sample.xlsx",
  false,
  "### Folder ID on Google Drive ###"
);
```

In this case, only an Excel file is downloaded to Google Drive.

If you use following simple script, the file `sample.xlsx` is just created to root directory.

```javascript
var prop = PropertiesService.getScriptProperties();
var odapp = OnedriveApp.init(prop);
var res = odapp.downloadFile("/SampleFolder/sample.xlsx");
```

#### Note :

**From my previous experiences, I think that the maximum response size using URL Fetch is about 10 MB. And furthermore, there are the limitations for the download size in 1 day ([URL Fetch data received 100MB / day](https://developers.google.com/apps-script/guides/services/quotas#current_limitations)). So when you use this download method, please be careful the file size.**

<a name="uploadfiles"></a>

### 5. Upload files from Google Drive to OneDrive.

```javascript
var fileid = "### file id ###";
var prop = PropertiesService.getScriptProperties();
var odapp = OnedriveApp.init(prop);
var res = odapp.uploadFile(fileid, "/### folder name on OneDrive ###/");
```

Uploaded filename and ID on OneDrive are returned. In the case of folder, please enclose it in `/`.

When you want to upload a Spreadsheet on Google Drive to a folder of `SampleFolder`, the Spreadsheet is converted to Excel file and uploaded to OneDrive. As a sample, when it uploads Spreadsheet to `/SampleFolder/` on OneDrive, the script is as follows.

```javascript
var fileid = "### file id ###";
var prop = PropertiesService.getScriptProperties();
var odapp = OnedriveApp.init(prop);
var res = odapp.uploadFile(fileid, "/SampleFolder/sample.xlsx");
```

At the following script, a file with the file id is uploaded to the root directory on OneDrive.

```javascript
var fileid = "### file id ###";
var prop = PropertiesService.getScriptProperties();
var odapp = OnedriveApp.init(prop);
var res = odapp.uploadFile(fileid);
```

#### Note :

**About this upload, in this library, [the resumable upload](https://dev.onedrive.com/items/upload_large_files.htm) is used for uploading files. So you can upload files with large size to OneDrive. But the chunk size is 10 MB, because of [the limitation of URL Fetch POST size on Google](https://developers.google.com/apps-script/guides/services/quotas#current_limitations). The file with large size is uploaded by separating by 10 MB. There are no limitations for upload size in one day.**

<a name="utilities"></a>

## Utilities

<a name="getaccesstoken"></a>

### 1. Get access token

```javascript
const accessToken = OnedriveApp.init(
  PropertiesService.getScriptProperties()
).getAccessToken();
console.log(accessToken);
```

The access token is simply returned. When this access token is used, you can also test other methods of Microsoft Graph API.

<a name="email"></a>

## Email of Microsoft Graph API

These methods manage Email.

<a name="getemaillist"></a>

### 1. Get Email message list

```javascript
const obj = {
  numberOfEmails: 1,
  select: ["createdDateTime", "sender", "subject", "bodyPreview"],
  orderby: "createdDateTime",
  order: "desc",
  folderId: "###",
};

const prop = PropertiesService.getScriptProperties();
const odapp = OnedriveApp.init(prop);
const res = odapp.getEmailList(obj);
console.log(res);
```

This method retrieves email list of your Microsoft account using Microsoft Graph API.

About the properties of `obj`, you can see the following explanation.

- `numberOfEmails`: Number of output email in the list. If the properties of `numberOfEmails` and `folderId` are not used, all email messages are retrieved.
- `select: Properties you want to retrieve for each email of list. When you use `["*"]`, all properties are retrieved. But in this case, the process time is longer. Please be careful this.
- `orderby`: When this is used, the list is ordered by the value of `orderby`.
- `order`: `desc` or `asc`.
- `folderId`: When you use this property, you can retrieve the email list from the specific folder of Email. About the method for retrieving the folder ID, please check "[Forward Email messages](#forwardemails)". When you don't use this property, all emails are retrieved as a list.

Sample script: [This is a sample script for retrieving all emails from own emails of Microsoft and put to the Google Spreadsheet.](https://gist.github.com/tanaikech/45a5511cf2a4a42660b52b3409f7b537)

<a name="getemails"></a>

### 2. Get Email messages

This method retrieves email messages of your Microsoft account using message IDs with Microsoft Graph API. As an important point of this method, in this method, the multiple emails can be retrieved using the batch request.

```javascript
const ar = ["messageId1", "messageId2", , ,];

const prop = PropertiesService.getScriptProperties();
const odapp = OnedriveApp.init(prop);
const res = odapp.getEmailMessages(ar);
console.log(res);
```

- The argument of `getEmailMessages(ar)` is an array including the message IDs. You can retrieve the message IDs with [the method of `getEmailList`](#getemaillist).

<a name="sendemails"></a>

### 3. Send Email messages

This method sends email messages using Microsoft Graph API with your Microsoft account. As an important point of this method, in this method, the multiple emails can be sent using the batch request.

```javascript
const obj = [
  {
    to: [{ name: "### name ###", email: "### email address ###" }, , ,],
    subject: "sample subject 1",
    body: "sample text body",
    cc: [{ name: "name1", email: "emailaddress1" }, , ,],
  },
  {
    to: [{ name: "### name ###", email: "### email address ###" }, , ,],
    subject: "sample subject 2",
    htmlBody: "<u><b>sample html body</b></u>",
    attachments: [blob],
    bcc: [{ name: "name1", email: "emailaddress1" }, , ,],
  },
];

const prop = PropertiesService.getScriptProperties();
const odapp = OnedriveApp.init(prop);
const res = odapp.sendEmails(obj);
console.log(res);
```

About the properties of `obj`, you can see the following explanation.

`to`: Email address of recipient. You can set the values as an array.
`subject`: Email subject.
`body`: Text body.
`htmlBody`: HTML body. In this case, as the current specification of Microsoft Graph API, it seems that both the text body and the HTML body cannot be used. So please use one of them.
`attachments`: For example, when you want to use the file on Google Drive, you can use `DriveApp.getFileById("### file ID ###").getBlob()`.
`cc`: Email addresses to CC.
`bcc`: Email addresses to BCC.

<a name="replyemails"></a>

### 4. Reply Email messages

This method replies email messages using Microsoft Graph API with your Microsoft account. As an important point of this method, in this method, the multiple emails can be replied using the batch request.

```javascript
const obj = [
  {
    to: [{ name: "### name ###", email: "### email address ###" }],
    body: "Sample replying message",
    messageId: "###",
  },
];

const prop = PropertiesService.getScriptProperties();
const odapp = OnedriveApp.init(prop);
const res = odapp.replyEmails(obj);
console.log(res);
```

About the properties of `obj`, you can see the following explanation.

The properties of `obj` are the same with [Send Email messages](#sendemails). But, in this case, please include `messageId` to reply to the email message.

<a name="forwardemails"></a>

### 5. Forward Email messages

This method forwards email messages using Microsoft Graph API with your Microsoft account. As an important point of this method, in this method, the multiple emails can be forwarded using the batch request.

```javascript
const obj = [
  {
    to: [{ name: "### name ###", email: "### email address ###" }],
    messageId: "###",
  },
];

const prop = PropertiesService.getScriptProperties();
const odapp = OnedriveApp.init(prop);
const res = odapp.forwardEmails(obj);
console.log(res);
```

About the properties of `obj`, you can see the following explanation.

The properties of `obj` are the same with [Send Email messages](#sendemails). But, in this case, please include `messageId` to reply to the email message.

<a name="getemailfolders"></a>

### 6. Get Email folders

This method retrieves the folders of Email of your Microsoft account.

```javascript
const prop = PropertiesService.getScriptProperties();
const odapp = OnedriveApp.init(prop);
const res = odapp.getEmailFolders();
console.log(res);
```

<a name="deleteemails"></a>

### 7. Delete Email messages

This method deletes the email messages of your Microsoft account. As an important point of this method, in this method, the multiple emails can be deleted using the batch request.

```javascript
const ar = ["messageId1", "messageId2", , ,];

const prop = PropertiesService.getScriptProperties();
const odapp = OnedriveApp.init(prop);
const res = odapp.deleteEmails(ar);
console.log(res);
```

Please set the message IDs you want to delete. You can retrieve the message IDs with [the method of `getEmailList`](#getemaillist).

<a name="licence"></a>

# Licence

[MIT](licence)

<a name="author"></a>

# Author

[Tanaike](https://tanaikech.github.io/about/)

If you have any questions and commissions for me, feel free to tell me.

<a name="updatehistory"></a>

# Update History

- v1.0.0 (August 16, 2017)

  Initial release.

- v1.0.1 (August 21, 2017)

  [Added a method for retrieving access token and refresh token using this library.](#authprocess)

- v1.0.2 (August 21, 2017)

  [Moved the instance of `PropertiesService.getScriptProperties()` to outside of this library. When there is the `PropertiesService.getScriptProperties()` inside the library, it was found that the parameters that users set was saved to the library. So this was modified. I'm sorry that I couldn't notice this situation.](#authprocess)

- v1.1.0 (September 24, 2017)

  [From this version, retrieving access token and refresh token became more easy.](#authprocess)

- v1.1.1 (July 28, 2018)

  A serious bug was removed. I had forgot that it added setProp(). I'm really sorry for this. And thank you so much jesus21282.

- v1.1.1 (January 4, 2020)

  ["Retrieve access token and refresh token for using OneDrive"](#authprocess) of the document was changed.

- v1.1.2 (September 29, 2021)

  A bug of method of `uploadFile` was removed. By this, the files except for Google Docs files can be uploaded to OneDrive.

- v1.2.0 (October 4, 2021)

  [1 method for retrieving the access token](#utilities) and [7 methods for managing emails of Microsoft account](#email) were added. By this, the emails got to be able to be gotten and sent using Microsoft account using OnedriveApp with Google Apps Script.

# Etc

If you want the sample script for uploading the contents using node.js, please check [here](https://gist.github.com/tanaikech/22bfb05e61f0afb8beed29dd668bdce9).

[TOP](#top)

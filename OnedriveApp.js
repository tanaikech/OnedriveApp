/**
 * Set parameters to PropertiesService<br>
 * At first, please set parameters using this. Access token is retrieved by the refresh token.<br>
 *<br>
 * If you don't have refresh token, don't worry. In this case, please set only client_id and client_secret.<br>
 * So please run this 'setProp(client_id, client_secret)'. And please paste following script.<br>
 *<br>
 * function doGet(){<br>
 *   var prop = PropertiesService.getScriptProperties();<br>
 *   return OnedriveApp.getAccesstoken(prop);<br>
 * }<br>
 *<br>
 * On the Script Editor -> Publish -> Deploy as Web App -> Click Test web app for your latest code.<br>
 * Please authorize by above process.<br>
 *<br>
 * @param {Object} PropertiesService.getScriptProperties()
 * @param {String} client_id
 * @param {String} client_secret
 * @param {String} redirect_uri
 * @param {String} refresh_token
 * @param {String} scope
 * @return {Object} return values from PropertiesService
 */
function setProp(prop, client_id, client_secret, redirect_uri, refresh_token, scope){
    return new OnedriveApp(false, prop).setprop(client_id, client_secret, redirect_uri, refresh_token, scope);
}

/**
 * Create OnedriveApp instance<br>
 * @param {Object} PropertiesService.getScriptProperties()
 * @return {OnedriveApp} return instance of OnedriveApp
 */
function init(prop) {
    return new OnedriveApp(true, prop);
}


/**
 * Retrieve access token and refresh token<br>
 * @param {Object} PropertiesService.getScriptProperties()
 * @return {HTML} return HTML code with retrieved access token and refresh token
 */
function getAccesstoken(prop, e) {
    var OA = new OnedriveApp(false, prop);
    if (!e.length) {
        return OA.doGet();
    } else if (e.length > 1) {
        if (e[0] == "doget") {
            OA.saveRefreshtoken(JSON.parse(e[1]).refresh_token);
        }
    }
    return;
}


/**
 * Retrieve authorization code.<br>
 * This is automatically used by this library.<br>
 * @param {Object} JSON data with authorization code
 * @return {HTML} return HTML code with access token
 */
function getCode(e) {
    return new OnedriveApp(false).callback(e);
}
;
(function(r) {
  var OnedriveApp;
  OnedriveApp = (function() {
    var addQuery, batchRequest, convToGoogle, convToMicrosoft, createEmailBody, fetch, getaccesstoken, getparams;

    OnedriveApp.name = "OnedriveApp";

    function OnedriveApp(d, p) {
      this.authurl = "https://login.microsoftonline.com/common/oauth2/v2.0/";
      this.scopes = "offline_access files.readwrite.all files.readwrite files.read Mail.ReadBasic Mail.Read Mail.ReadWrite Mail.Send";
      this.p = p;
      if (d) {
        this.prop = this.p.getProperties();
        this.client_id = this.prop.client_id;
        this.client_secret = this.prop.client_secret;
        this.redirect_uri = this.prop.redirect_uri;
        this.refresh_token = this.prop.refresh_token;
        this.access_token = getaccesstoken.call(this);
        this.baseurl = "https://graph.microsoft.com/v1.0";
        this.maxchunk = 10485760;
        this.sheeturl = "https://graph.microsoft.com/beta/me/drive/items/";
        if (this.refresh_token == null) {
          throw new Error("No refresh token. Please save refresh token by 'OnedriveApp.saveRefreshtoken(prop, refresh_token)'.");
        }
      }
    }

    OnedriveApp.prototype.setprop = function(client_id, client_secret, redirect_uri, refresh_token, scope) {
      this.p.setProperties({
        client_id: client_id,
        client_secret: client_secret,
        redirect_uri: redirect_uri,
        refresh_token: refresh_token,
        scope: scope || this.scopes
      });
      return JSON.stringify(this.p.getProperties());
    };

    OnedriveApp.prototype.saveRefreshtoken = function(refresh_token_) {
      if (refresh_token_ == null) {
        throw new Error("No refresh token.");
      }
      this.p.setProperties({
        refresh_token: refresh_token_
      });
      return JSON.stringify(this.p.getProperties());
    };

    OnedriveApp.prototype.doGet = function() {
      var appurl, ermsg, html, name, param, prop, qparams, url, value;
      prop = this.p.getProperties();
      if ((prop.client_id == null) || (prop.client_secret == null) || !prop.client_id || !prop.client_secret) {
        ermsg = "Error: Please set client_id and client_secret to ScriptProperties using 'OnedriveApp.setProp(client_id, client_secret)'.\n";
        return HtmlService.createHtmlOutput(ermsg);
      }
      url = this.authurl + "authorize";
      param = {
        response_type: "code",
        response_mode: "query",
        client_id: prop.client_id,
        redirect_uri: (function(p1) {
          var rurl;
          this.p = p1;
          rurl = ScriptApp.getService().getUrl();
          rurl = rurl.indexOf("/exec") >= 0 ? rurl = rurl.slice(0, -4) + 'usercallback' : rurl.slice(0, -3) + 'usercallback';
          this.p.setProperties({
            redirect_uri: rurl
          });
          return rurl;
        })(this.p),
        scope: prop.scope,
        state: ScriptApp.newStateToken().withArgument('client_id', prop.client_id).withArgument('client_secret', prop.client_secret).withArgument('redirect_uri', this.p.getProperties().redirect_uri).withMethod("OnedriveApp.getCode").withTimeout(300).createToken()
      };
      qparams = "?";
      for (name in param) {
        value = param[name];
        qparams += name + "=" + encodeURIComponent(value) + "&";
      }
      appurl = "https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationMenuBlade/~/Overview/appId/" + prop.client_id + "/isMSAApp~/true";
      html = "<p>Please push this button after set redirect_uri to '<b>" + param.redirect_uri + "</b>' at <a href=\"" + appurl + "\" target=\"_blank\">your application</a>.</p>";
      html += "<input type=\"button\" value=\"Get access token\" onclick=\"window.open('" + url + qparams + "', 'Authorization', 'width=500,height=600');\">";
      return HtmlService.createHtmlOutput(html);
    };

    OnedriveApp.prototype.callback = function(e) {
      var ermsg, method, payload, res, t, url;
      if (e.parameter.code == null) {
        ermsg = "Error: Please confirm client_id, client_secret and redirect_uri, again.\n";
        ermsg += "client_id, client_secret and redirect_uri you set are as follows.\n";
        ermsg += "client_id : " + e.parameter.client_id + "\n";
        ermsg += "client_secret : " + e.parameter.client_secret + "\n";
        ermsg += "redirect_uri : " + e.parameter.redirect_uri + "\n";
        return HtmlService.createHtmlOutput(ermsg);
      }
      url = this.authurl + "token";
      method = "post";
      payload = {
        client_id: e.parameter.client_id,
        client_secret: e.parameter.client_secret,
        redirect_uri: e.parameter.redirect_uri,
        code: e.parameter.code,
        grant_type: "authorization_code",
        scope: this.scopes
      };
      res = fetch.call(this, url, method, payload, null);
      t = HtmlService.createTemplateFromFile('doget');
      t.data = res;
      return t.evaluate();
    };

    OnedriveApp.prototype.getAccessToken = function() {
      return this.access_token;
    };

    OnedriveApp.prototype.getEmailList = function(obj) {
      var ar, headers, keys, method, number, res, url;
      url = this.baseurl + "/me/messages";
      number = -1;
      if (obj) {
        if (obj.hasOwnProperty("numberOfEmails")) {
          number = obj.numberOfEmails;
          delete obj.numberOfEmails;
        }
        if (obj.hasOwnProperty("select") && obj.select.indexOf("*") === -1) {
          obj.select = obj.select.join(",");
        } else if (obj.select.indexOf("*") > -1) {
          delete obj.select;
        }
        if (obj.hasOwnProperty("orderby")) {
          obj.orderby = obj.orderby + " " + (obj.hasOwnProperty("order") ? obj.order : "asc");
          if (obj.hasOwnProperty("order")) {
            if (obj.order !== "asc" && obj.order !== "desc") {
              throw new Error("Order is 'asc' or 'desc'.");
            }
            delete obj.order;
          }
        }
        if (obj.hasOwnProperty("folderId")) {
          url = this.baseurl + "/me/mailFolders/" + obj.folderId + "/messages";
          delete obj.folderId;
        }
        keys = Object.keys(obj);
        keys.forEach(function(k) {
          obj["$" + k] = obj[k];
          return delete obj[k];
        });
      } else {
        obj = {
          $select: "sender,subject,bodyPreview",
          $orderby: "createdDateTime asc"
        };
      }
      obj["$top"] = 1000;
      url += addQuery(obj);
      method = "get";
      headers = {
        "Authorization": "Bearer " + this.access_token
      };
      ar = [];
      while (true) {
        res = fetch.call(this, url, method, null, headers, null);
        if (!res.hasOwnProperty("value")) {
          throw new Error("Invalid parameter. Please check it again.");
        }
        if (res.value.length > 0) {
          ar = ar.concat(res.value);
        }
        url = res["@odata.nextLink"];
        if (!res.hasOwnProperty("@odata.nextLink")) {
          break;
        }
      }
      if (number !== -1) {
        ar.splice(number);
      }
      return ar;
    };

    OnedriveApp.prototype.getEmailMessages = function(obj) {
      var requests;
      if (!obj || !Array.isArray(obj)) {
        throw new Error('Please set message IDs as an array like ["messageId1", "messageId2",,,].');
      }
      requests = obj.map(function(id, i) {
        return {
          url: "/me/messages/" + id,
          method: "GET",
          id: i + 1
        };
      });
      return batchRequest.call(this, requests);
    };

    OnedriveApp.prototype.sendEmails = function(obj) {
      var requests;
      if (!obj || !Array.isArray(obj)) {
        throw new Error('Please set object for sending messages.');
      }
      if (obj.some(function(e) {
        return !e.subject || (!e.body && !e.htmlBody) || !e.to;
      })) {
        throw new Error('Please check the object again.');
      }
      requests = obj.map(function(e, i) {
        return {
          url: "/me/sendMail",
          method: "POST",
          id: i + 1,
          body: createEmailBody.call(this, e),
          headers: {
            "Content-Type": "application/json"
          }
        };
      });
      return batchRequest.call(this, requests);
    };

    OnedriveApp.prototype.replyEmails = function(obj) {
      var requests;
      if (!obj || !Array.isArray(obj)) {
        throw new Error('Please set object for sending messages.');
      }
      if (obj.some(function(e) {
        return !e.messageId || (!e.body && !e.htmlBody) || !e.to;
      })) {
        throw new Error('Please check the object again.');
      }
      requests = obj.map(function(e, i) {
        var body;
        body = createEmailBody.call(this, e);
        delete body.message.messageId;
        delete body.message.subject;
        return {
          url: "/me/messages/" + e.messageId + "/reply",
          method: "POST",
          id: i + 1,
          body: body,
          headers: {
            "Content-Type": "application/json"
          }
        };
      });
      return batchRequest.call(this, requests);
    };

    OnedriveApp.prototype.forwardEmails = function(obj) {
      var requests;
      if (!obj || !Array.isArray(obj)) {
        throw new Error('Please set object for sending messages.');
      }
      if (obj.some(function(e) {
        return !e.messageId || !e.to;
      })) {
        throw new Error('Please check the object again.');
      }
      requests = obj.map(function(e, i) {
        var body;
        body = createEmailBody.call(this, e);
        delete body.message.messageId;
        return {
          url: "/me/messages/" + e.messageId + "/forward",
          method: "POST",
          id: i + 1,
          body: body,
          headers: {
            "Content-Type": "application/json"
          }
        };
      });
      return batchRequest.call(this, requests);
    };

    OnedriveApp.prototype.getEmailFolders = function() {
      var headers, method, res, url;
      url = this.baseurl + "/me/mailFolders?$top=1000";
      method = "get";
      headers = {
        "Authorization": "Bearer " + this.access_token
      };
      res = fetch.call(this, url, method, null, headers, null);
      if (!res.hasOwnProperty("value")) {
        throw new Error("Invalid parameter. Please check it again.");
      }
      return res;
    };

    OnedriveApp.prototype.deleteEmails = function(obj) {
      var requests;
      if (!obj || !Array.isArray(obj)) {
        throw new Error('Please set message IDs as an array like ["messageId1", "messageId2",,,].');
      }
      requests = obj.map(function(id, i) {
        return {
          url: "/me/messages/" + id,
          method: "DELETE",
          id: i + 1
        };
      });
      return batchRequest.call(this, requests);
    };

    OnedriveApp.prototype.getESheet = function(id) {
      var headers, method, res, url;
      url = this.sheeturl + id + "/workbook/worksheets";
      method = "get";
      headers = {
        "Authorization": "Bearer " + this.access_token
      };
      res = fetch.call(this, url, method, null, headers, null);
      return res;
    };

    OnedriveApp.prototype.getAt = function() {
      return this.access_token;
    };

    OnedriveApp.prototype.createSession = function(path) {
      var headers, method, res, url;
      url = this.sheeturl + "root:/" + path + ":/workbook/";
      method = "get";
      headers = {
        "Authorization": "Bearer " + this.access_token
      };
      res = fetch.call(this, url, method, null, headers, null);
      return res;
    };

    OnedriveApp.prototype.getFilelist = function(folder) {
      var headers, method, res, url;
      url = this.baseurl + "/drive/items/root" + (folder ? ":/" + folder : "") + "?expand=children(select=id,name)";
      method = "get";
      headers = {
        "Authorization": "Bearer " + this.access_token
      };
      res = fetch.call(this, url, method, null, headers, null);
      return res.children;
    };

    OnedriveApp.prototype.downloadFile = function(path_, conv_, googlefolderId_) {
      var blob, dfile, fileId_, filename, folder, folderName, headers, method, query, url;
      if (path_ == null) {
        throw new Error("No path.");
      }
      if (conv_ == null) {
        conv_ = false;
      }
      if (googlefolderId_ == null) {
        googlefolderId_ = false;
      }
      fileId_ = "";
      folderName = "root";
      url = this.baseurl + "/me/drive/root:" + path_ + ":/content";
      method = "get";
      headers = {
        "Authorization": "Bearer " + this.access_token
      };
      blob = fetch.call(this, url, method, null, headers, null);
      filename = path_.substring(path_.lastIndexOf('/') + 1, path_.length);
      if (googlefolderId_) {
        folder = DriveApp.getFolderById(googlefolderId_);
        fileId_ = folder.createFile(blob).setName(filename).getId();
        folderName = folder.getName();
      } else {
        fileId_ = DriveApp.createFile(blob).setName(filename).getId();
      }
      if (conv_) {
        dfile = fileId_;
        fileId_ = convToGoogle.call(this, dfile);
        if (googlefolderId_) {
          query = "?addParents=" + googlefolderId_;
          query += "&removeParents=" + DriveApp.getFileById(fileId_).getParents().next().getId();
          url = "https://www.googleapis.com/drive/v3/files/" + fileId_ + query;
          headers = {
            "Authorization": "Bearer " + ScriptApp.getOAuthToken()
          };
          method = "patch";
          fetch.call(this, url, method, null, headers, null);
        }
        url = "https://www.googleapis.com/drive/v3/files/" + dfile;
        headers = {
          "Authorization": "Bearer " + ScriptApp.getOAuthToken()
        };
        method = "delete";
        fetch.call(this, url, method, null, headers, null);
      }
      return filename + " (fileId = " + fileId_ + ") was downloaded to " + folderName + " folder on your Google Drive.";
    };

    OnedriveApp.prototype.uploadFile = function(fileid_, path_) {
      var ar, byteAr, file, fileinf, filepath, filesize, headers, l, len, method, payload, ref, res, result, st, url;
      if (fileid_ == null) {
        throw new Error("No file ID.");
      }
      path_ = path_ == null ? "/" : path_;
      fileinf = (ref = convToMicrosoft.call(this, fileid_)) != null ? ref : [fileid_, DriveApp.getFileById(fileid_).getName()];
      file = DriveApp.getFileById(fileinf[0]);
      filesize = 0;
      byteAr = [];
      if (!~file.getMimeType().indexOf("google-apps.script")) {
        byteAr = file.getBlob().getBytes();
        filesize = byteAr.length;
      } else {
        throw new Error("Cannot upload '" + fileinf[0] + "'.");
      }
      ar = getparams.call(this, filesize);
      filepath = path_ + fileinf[1];
      url = this.baseurl + "/drive/root:" + filepath + ":/createUploadSession";
      headers = {
        "Authorization": "Bearer " + this.access_token,
        "Content-Type": "application/json"
      };
      method = "post";
      payload = '{"item": {"@microsoft.graph.conflictBehavior": "rename", "name": "' + fileinf[1] + '"}}';
      url = (fetch.call(this, url, method, payload, headers, null)).uploadUrl;
      method = "put";
      res = [];
      for (l = 0, len = ar.length; l < len; l++) {
        st = ar[l];
        headers = {
          "Content-Range": st.cr
        };
        payload = byteAr.slice(st.bstart, st.bend + 1);
        res.push(fetch.call(this, url, method, payload, headers, st.clen));
      }
      result = res[res.length - 1];
      if (fileinf[0] !== fileid_) {
        url = "https://www.googleapis.com/drive/v3/files/" + fileinf[0];
        headers = {
          "Authorization": "Bearer " + ScriptApp.getOAuthToken()
        };
        method = "delete";
        fetch.call(this, url, method, null, headers, null);
      }
      return {
        name: result.name,
        id: result.id,
        size: result.size,
        createdDateTime: result.createdDateTime
      };
    };

    OnedriveApp.prototype.creatFolder = function(foldername_, path_) {
      var headers, method, payload, result, url;
      if (foldername_ == null) {
        throw new Error("No folder name.");
      }
      url = path_ ? this.baseurl + "/drive/root:" + path_ + ":/children" : this.baseurl + "/drive/root/children";
      headers = {
        "Authorization": "Bearer " + this.access_token,
        "Content-Type": "application/json"
      };
      method = "post";
      payload = '{"name": "' + foldername_ + '", "folder": { }}';
      result = fetch.call(this, url, method, payload, headers, null);
      return {
        folderName: result.name,
        id: result.id,
        parentInf: result.parentReference.path
      };
    };

    OnedriveApp.prototype.deleteItemByName = function(path_) {
      var headers, method, result, url;
      if (path_ == null) {
        throw new Error("No path to item on OneDrive.");
      }
      if (path_.slice(-1) === "/") {
        path_ = path_.slice(0, -1);
      }
      url = this.baseurl + "/drive/root:" + path_;
      headers = {
        "Authorization": "Bearer " + this.access_token
      };
      method = "delete";
      result = fetch.call(this, url, method, null, headers, null);
      if (result.error != null) {
        return result;
      } else {
        return {
          message: "'/root" + path_ + "' was deleted."
        };
      }
    };

    OnedriveApp.prototype.deleteItemById = function(id_) {
      var headers, method, result, url;
      if (id_ == null) {
        throw new Error("No item id.");
      }
      url = this.baseurl + "/drive/items/" + id_;
      headers = {
        "Authorization": "Bearer " + this.access_token
      };
      method = "delete";
      result = fetch.call(this, url, method, null, headers, null);
      if (result.error != null) {
        return result;
      } else {
        return {
          message: "'" + id_ + "' was deleted."
        };
      }
    };

    OnedriveApp.prototype.convToMicrosoftDo = function(id_) {
      return convToMicrosoft.call(this, id_);
    };

    OnedriveApp.prototype.convToGoogleDo = function(id_, filename_) {
      return convToGoogle.call(this, id_);
    };

    convToGoogle = function(fileId_) {
      var ToMime, boundary, data, fields, file, filename, headers, metadata, method, mime, payload, res, url;
      if (fileId_ == null) {
        throw new Error("No file ID.");
      }
      file = DriveApp.getFileById(fileId_);
      filename = file.getName();
      mime = file.getMimeType();
      ToMime = "";
      switch (mime) {
        case "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
          ToMime = "application/vnd.google-apps.document";
          break;
        case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
          ToMime = "application/vnd.google-apps.spreadsheet";
          break;
        case "application/vnd.openxmlformats-officedocument.presentationml.presentation":
          ToMime = "application/vnd.google-apps.presentation";
          break;
        default:
          return null;
      }
      metadata = {
        name: filename,
        mimeType: ToMime
      };
      fields = "id,mimeType,name";
      url = "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=" + encodeURIComponent(fields);
      boundary = "xxxxxxxxxx";
      data = "--" + boundary + "\r\n";
      data += "Content-Disposition: form-data; name=\"metadata\";\r\n";
      data += "Content-Type: application/json; charset=UTF-8\r\n\r\n";
      data += JSON.stringify(metadata) + "\r\n";
      data += "--" + boundary + "\r\n";
      data += "Content-Disposition: form-data; name=\"file\"; filename=\"" + filename + "\"\r\n";
      data += "Content-Type: " + mime + "\r\n\r\n";
      payload = Utilities.newBlob(data).getBytes().concat(file.getBlob().getBytes()).concat(Utilities.newBlob("\r\n--" + boundary + "\r\n").getBytes());
      headers = {
        "Authorization": "Bearer " + ScriptApp.getOAuthToken(),
        "Content-Type": "multipart/related; boundary=" + boundary
      };
      method = "post";
      res = fetch.call(this, url, method, payload, headers, null);
      return res.id;
    };

    convToMicrosoft = function(fileId_) {
      var blob, deffilename, ext, file, fileid, filename, format, headers, method, mime, url;
      if (fileId_ == null) {
        throw new Error("No file ID.");
      }
      file = DriveApp.getFileById(fileId_);
      mime = file.getMimeType();
      format = "";
      ext = "";
      switch (mime) {
        case "application/vnd.google-apps.document":
          format = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
          ext = ".docx";
          break;
        case "application/vnd.google-apps.spreadsheet":
          format = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
          ext = ".xlsx";
          break;
        case "application/vnd.google-apps.presentation":
          format = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
          ext = ".pptx";
          break;
        default:
          return null;
      }
      url = "https://www.googleapis.com/drive/v3/files/" + fileId_ + "/export?mimeType=" + format;
      headers = {
        "Authorization": "Bearer " + ScriptApp.getOAuthToken()
      };
      method = "get";
      blob = fetch.call(this, url, method, null, headers, null);
      deffilename = file.getName();
      filename = ~deffilename.indexOf(ext) ? deffilename : deffilename + ext;
      fileid = DriveApp.createFile(blob).setName(filename).getId();
      return [fileid, filename];
    };

    getparams = function(filesize) {
      var allsize, ar, bend, bstart, clen, cr, i, l, ref, ref1, sep;
      allsize = filesize;
      sep = allsize < this.maxchunk ? allsize : this.maxchunk - 1;
      ar = [];
      for (i = l = 0, ref = allsize - 1, ref1 = sep; ref1 > 0 ? l <= ref : l >= ref; i = l += ref1) {
        bstart = i;
        bend = i + sep - 1 < allsize ? i + sep - 1 : allsize - 1;
        cr = 'bytes ' + bstart + '-' + bend + '/' + allsize;
        clen = bend !== allsize - 1 ? sep : allsize - i;
        ar.push({
          bstart: bstart,
          bend: bend,
          cr: cr,
          clen: clen
        });
      }
      return ar;
    };

    getaccesstoken = function() {
      var method, payload, res, url;
      url = this.authurl + "token";
      method = "post";
      payload = {
        client_id: this.client_id,
        client_secret: this.client_secret,
        redirect_uri: this.redirect_uri,
        refresh_token: this.refresh_token,
        grant_type: "refresh_token"
      };
      res = fetch.call(this, url, method, payload, null, null);
      if (res.refresh_token !== this.refresh_token) {
        this.p.setProperties({
          refresh_token: res.refresh_token
        });
      }
      this.access_token = res.access_token;
      if (this.access_token === null || this.access_token === "") {
        throw new Error("At first, please run setProp().");
      }
      return res.access_token;
    };

    fetch = function(url, method, payload, headers, contentLength) {
      var e, res;
      try {
        res = UrlFetchApp.fetch(url, {
          method: method,
          payload: payload,
          headers: headers,
          contentLength: contentLength,
          muteHttpExceptions: true
        });
      } catch (error) {
        e = error;
        throw new Error(e);
      }
      try {
        r = JSON.parse(res.getContentText());
      } catch (error) {
        e = error;
        r = res.getBlob();
      }
      return r;
    };

    batchRequest = function(requests) {
      var ar, headers, i, l, limit, method, payload, ref, res, split, url;
      url = this.baseurl + "/$batch";
      method = "POST";
      headers = {
        "Authorization": "Bearer " + this.access_token,
        "Content-Type": "application/json"
      };
      ar = [];
      limit = 20;
      split = Math.ceil(requests.length / limit);
      for (i = l = 0, ref = split; 0 <= ref ? l < ref : l > ref; i = 0 <= ref ? ++l : --l) {
        payload = {
          requests: requests.splice(0, limit)
        };
        res = fetch.call(this, url, method, JSON.stringify(payload), headers, null);
        ar = ar.concat(res.responses);
      }
      return ar;
    };

    addQuery = function(obj) {
      return Object.keys(obj).reduce(function(p, e, i) {
        return p + (i === 0 ? "?" : "&") + (Array.isArray(obj[e]) ? obj[e].reduce(function(str, f, j) {
          return str + e + "=" + encodeURIComponent(f) + (j !== obj[e].length - 1 ? "&" : "");
        }, "") : e + "=" + encodeURIComponent(obj[e]));
      }, "");
    };

    createEmailBody = function(e) {
      var remain, temp;
      e.to = e.to.filter(function(f) {
        return f && typeof f === "object";
      });
      temp = {
        message: {
          subject: e.subject,
          toRecipients: e.to.map(function(f) {
            return {
              emailAddress: {
                name: f.name,
                address: f.email
              }
            };
          })
        }
      };
      delete e.subject;
      delete e.to;
      if (e.body) {
        temp.message.body = {
          contentType: "Text",
          content: e.body
        };
        delete e.body;
      } else if (e.htmlBody) {
        temp.message.body = {
          contentType: "HTML",
          content: e.htmlBody
        };
        delete e.htmlBody;
      }
      if (e.cc && Array.isArray(e.cc)) {
        e.cc = e.cc.filter(function(f) {
          return f && typeof f === "object";
        });
        temp.message.ccRecipients = e.cc.map(function(f) {
          return {
            emailAddress: {
              name: f.name,
              address: f.email
            }
          };
        });
        delete e.cc;
      }
      if (e.bcc && Array.isArray(e.bcc)) {
        e.bcc = e.bcc.filter(function(f) {
          return f && typeof f === "object";
        });
        temp.message.bccRecipients = e.bcc.map(function(f) {
          return {
            emailAddress: {
              name: f.name,
              address: f.email
            }
          };
        });
        delete e.bcc;
      }
      if (e.attachments && Array.isArray(e.attachments)) {
        temp.message.attachments = e.attachments.map(function(f) {
          if (f.toString() !== "Blob") {
            throw new Error('Please set the attachment file as blob.');
          }
          return {
            "@odata.type": "#microsoft.graph.fileAttachment",
            name: f.getName() || "no name",
            contentType: f.getContentType(),
            contentBytes: Utilities.base64Encode(f.getBytes())
          };
        });
        delete e.attachments;
      }
      remain = Object.keys(e);
      if (remain.length > 0) {
        remain.forEach(function(f) {
          return temp.message[f] = e[f];
        });
      }
      return temp;
    };

    return OnedriveApp;

  })();
  return r.OnedriveApp = OnedriveApp;
})(this);

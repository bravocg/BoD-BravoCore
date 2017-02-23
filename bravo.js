"use strict";

/*
* Title: Bravo Core Library
* Source: https://github.com/bravocg/core
* Version: v1.7.6
* Author: Gunjan Datta
* Description: The Bravo core library translates the REST api as an object model.
* 
* Copyright © 2015 Bravo Consulting Group, LLC (Bravo). All Rights Reserved.
* Released under the MIT license.
*/

// **********************************************************************************
// Namespace
// **********************************************************************************
if (window.Type) { window.Type.registerNamespace("BRAVO"); }
else { window.BRAVO = window.BRAVO || {}; }

// **********************************************************************************
// Bravo Core Class
// This class converts the REST api as an object model.
// **********************************************************************************
BRAVO.Core = function () {
    // **********************************************************************************
    // Global Variables
    // **********************************************************************************
    var _dependenciesLoadedFl = null;

    // **********************************************************************************
    // REST Methods
    // **********************************************************************************
    var _restMethods = function () {
        var methods = {
            // **********************************************************************************
            // By Metadata Type
            // **********************************************************************************
            // Content Type
            "SP.ContentType": {
                post: ["deleteObject"],
                custom: [
                    { name: "addFieldLink", "function": function (data) { this.executePost("fieldlinks", null, data, true, "SP.FieldLink"); } },
                    { name: "getFieldByInternalName", "function": function (name) { name = encodeURIComponent(name); return this.executeGet("fields?$filter=InternalName eq '" + name + "'"); } },
                    { name: "getFieldByStaticName", "function": function (name) { name = encodeURIComponent(name); return this.executeGet("fields?$filter=StaticName eq '" + name + "'"); } },
                    { name: "getFieldByTitle", "function": function (title) { title = encodeURIComponent(title); return this.executeGet("fields?$filter=Title eq '" + title + "'"); } },
                    { name: "getFieldLinkByName", "function": function (name) { name = encodeURIComponent(name); return this.executeGet("fieldlinks?$filter=Name eq '" + name + "'"); } },
                    { name: "update", "function": function (data) { return this.executePost(null, null, data, true, "SP.ContentType", "MERGE"); } }
                ]
            },
            // Field
            "SP.Field": {
                post: ["deleteObject", "setShowInDisplayForm", "setShowInEditForm", "setShowInNewForm"],
                custom: [
                    { name: "update", "function": function (data) { return this.executePost(null, null, data, true, "SP.Field", "MERGE"); } }
                ]
            },
            // Field Link
            "SP.FieldLink": {
                // TO DO - Validate if this object exists
            },
            // File
            "SP.File": {
                get: ["getLimitedWebPartManager"],
                getBuffer: ["openBinaryStream"],
                post: ["approve", "cancelUpload", "checkIn", "checkOut", "copyTo", "deleteObject", "deny",
                    "moveTo", "publish", "recycle", "undoCheckOut", "unpublish"],
                postDataInBody: ["continueUpload", "finishUpload", "saveBinaryStream", "startUpload"],
                custom: [
                    { name: "update", "function": function (data) { return this.executePost(null, null, data, true, "SP.File", "MERGE"); } }
                ]
            },
            // File Version
            "SP.FileVersion": {
                post: ["deleteObject"]
            },
            // Folder
            "SP.Folder": {
                post: ["deleteObject", "recycle", "uniqueContentTypeOrder"],
                custom: [
                    { name: "addFile", "function": function (data, content) { return this.executePost("files/add", data, content, true); } },
                    { name: "addSubFolder", "function": function (name) { return this.executePost("folders/add", name); } },
                    { name: "getFile", "function": function (name) { return this.executeGet("files?$filter=Name eq '" + name + "'"); } },
                    { name: "getSubFolder", "function": function (name) { return this.executeGet("folders?$filter=Name eq '" + name + "'"); } },
                    { name: "update", "function": function (data) { return this.executePost(null, null, data, true, "SP.Folder", "MERGE"); } }
                ]
            },
            // Group
            "SP.Group": {
                custom: [
                    { name: "containsUser", "function": function (user) { return this.executeGet("users?$filter=Id eq " + user.Id).exists; } }
                ]
            },
            // List
            "SP.List": {
                get: ["getRelatedFields", "getView"],
                getDataAsParameter: ["getUserEffectivePermissions"],
                post: ["addItem", "breakRoleInheritance", "deleteObject", "recycle", "renderListData", "renderListFormData",
                    "reserveListItemId", "resetRoleInheritance"],
                postDataInBodyNoArgs: ["getChanges", "getItems", "getListItemChangesSinceToken"],
                custom: [
                    { name: "addContentType", "function": function (data) { return this.executePost("contenttypes", null, data, true, "SP.ContentType"); } },
                    { name: "addExistingContentType", "function": function (data) { return this.executePost("contenttypes/addAvailableContentType", data); } },
                    { name: "addField", "function": function (data) { return this.executePost("fields/add", null, data, true, "SP.Field"); } },
                    { name: "addFieldAsXml", "function": function (data) { return this.executePost("fields/createFieldAsXml", null, { parameters: { __metadata: { type: "SP.XmlSchemaFieldCreationInformation" }, Options: SP.AddFieldOptions.addFieldInternalNameHint, SchemaXml: data } }, true); } },
                    { name: "addItem", "function": function (data) { return this.executePost("items", null, data, true, "SP.ListItem"); } },
                    { name: "addSiteGroup", "function": function (data) { return this.executePost("roleassignments/addroleassignment", data); } },
                    { name: "addSubFolder", "function": function (name) { return this.get_RootFolder().addSubFolder(name); } },
                    { name: "addView", "function": function (data) { return this.executePost("views", null, data, true, "SP.View"); } },
                    { name: "getContentType", "function": function (title) { title = encodeURIComponent(title); return this.executeGet("contenttypes?$filter=Name eq '" + title + "'"); } },
                    { name: "getContentTypeById", "function": function (id) { return this.executeGet("contenttypes/getById", id); } },
                    { name: "getDefaultDisplayFormUrl", "function": function () { return this.executeGet("forms?$filter=FormType eq 4").ServerRelativeUrl; } },
                    { name: "getDefaultEditFormUrl", "function": function () { return this.executeGet("forms?$filter=FormType eq 6").ServerRelativeUrl; } },
                    { name: "getDefaultNewFormUrl", "function": function () { return this.executeGet("forms?$filter=FormType eq 8").ServerRelativeUrl; } },
                    { name: "getDefaultViewUrl", "function": function () { return this.get_DefaultView().ServerRelativeUrl; } },
                    { name: "getField", "function": function (title) { title = encodeURIComponent(title); return this.executeGet("fields?$filter=Title eq '" + title + "' or InternalName eq '" + title + "' or StaticName eq '" + title + "'"); } },
                    { name: "getFieldById", "function": function (id) { return this.executeGet("fields/getById", id); } },
                    { name: "getFieldByInternalName", "function": function (name) { name = encodeURIComponent(name); return this.executeGet("fields?$filter=InternalName eq '" + name + "'"); } },
                    { name: "getFieldByStaticName", "function": function (name) { name = encodeURIComponent(name); return this.executeGet("fields?$filter=StaticName eq '" + name + "'"); } },
                    { name: "getFieldByTitle", "function": function (title) { title = encodeURIComponent(title); return this.executeGet("fields?$filter=Title eq '" + title + "'"); } },
                    { name: "getItemById", "function": function (id) { return this.executeGet("items(" + id + ")"); } },
                    { name: "getItemByTitle", "function": function (title) { title = encodeURIComponent(title); return this.getItemsByFilter("Title eq '" + title + "'"); } },
                    { name: "getItemsByFilter", "function": function (filter) { return this.executeGet("items?$filter=" + filter); } },
                    { name: "getSchemaXml", "function": function () { return this.executeGet("schemaxml"); } },
                    { name: "getSubFolder", "function": function (name) { name = encodeURIComponent(name); return this.executeGet("rootfolder/folders?$filter=Name eq '" + name + "'"); } },
                    { name: "getViewById", "function": function (id) { return this.executeGet("views/getById", id); } },
                    { name: "getViewByTitle", "function": function (title) { title = encodeURIComponent(title); return this.executeGet("views?$filter=Title eq '" + title + "'"); } },
                    { name: "hasAccess", "function": function (userName, permissions) { return hasAccess(this, permissions, userName); } },
                    { name: "update", "function": function (data) { return this.executePost(null, null, data, true, "SP.List", "MERGE"); } }
                ]
            },
            // List Item
            "SP.ListItem": {
                getDataAsParameter: ["getUserEffectivePermissions"],
                post: ["breakRoleInheritance", "deleteObject", "recycle", "resetRoleInheritance"],
                postDataInBodyNoArgs: ["validateUpdateListItem"],
                custom: [
                    {
                        name: "addAttachment",
                        "function": function (fileName, content) {
                            return this.executePost("attachmentfiles/add", { FileName: fileName }, content, true);
                        }
                    },
                    {
                        name: "addAttachmentFile",
                        "function": function (file) {
                            var thisObj = this;
                            return new Promise(
                                function (resolve, reject) {
                                    getFileInfo(file).then(
                                        function (args) {
                                            var name = args.name,
                                                buffer = args.buffer;

                                            if (name && buffer) {
                                                thisObj.addAttachment(name, buffer).then(
                                                    function (file) {
                                                        resolve(file);
                                                    }
                                                );
                                            } else {
                                                resolve();
                                            }
                                        }
                                    );
                                }
                            );
                        }
                    },
                    {
                        name: "update",
                        "function": function (data) {
                            return this.executePost(null, null, data, true, this.__metadata.type, "MERGE");
                        }
                    }
                ]
            },
            // Role Assignment
            "SP.RoleAssignment": {
                post: ["deleteObject"]
            },
            // Role Definition
            "SP.RoleDefinition": {
                post: ["deleteObject"]
            },
            // Search Service
            "Microsoft.Office.Server.Search.REST.SearchService": {
                custom: [
                    { name: "query", "function": function (query) { if (typeof (query) === "string") { return this.executeGet("query?" + query); } query = { request: query }; query.request.__metadata = { type: "Microsoft.Office.Server.Search.REST.SearchRequest" }; return this.executePost("postquery", null, query, true); } },
                    { name: "querySuggestion", "function": function (query) { return this.executeGet("suggest?" + query); } },
                ]
            },
            // Site
            "SP.Site": {
                post: ["createPreviewSPSite", "extendUpgradeReminderDate", "getCatalog", "getCustomListTemplates", "getWebTemplates",
                    "invalidate", "needsUpgradeByType", "openWeb", "openWebById", "runHealthCheck", "runUpgradeSiteSession",
                    "updateClientObjectModelUseRemoteAPIsPermissionSetting"],
                postDataInBodyNoArgs: ["getChanges"],
                custom: [
                    { name: "addCustomAction", "function": function (data) { return this.executePost("usercustomactions", null, data, true, "SP.UserCustomAction"); } },
                    { name: "getCustomAction", "function": function (title) { title = encodeURIComponent(title); return this.executeGet("usercustomactions?$filter=Name eq '" + title + "' or Title eq '" + title + "'"); } },
                    { name: "getRootWeb", "function": function () { this._rootWeb = this._rootWeb || new BRAVO.Core.Web(this.ServerRelativeUrl, this.asyncFl); return this._rootWeb; } },
                    { name: "hasAccess", "function": function (permissions) { return hasAccess(this, permissions); } },
                    { name: "sendEmail", "function": function (data) { data = { properties: data }; data.properties.__metadata = { type: "SP.Utilities.EmailProperties" }; return this.executePost("_api/SP.Utilities.Utility.SendEmail", null, data, true); } },
                    { name: "update", "function": function (data) { return this.executePost(null, null, data, true, "SP.Site", "MERGE"); } }
                ]
            },
            // Social User
            "SP.Social.SocialRestActor": {
                custom: [
                    { name: "createPost", "function": function (data) { data = { restCreationData: { __metadata: { type: "SP.Social.SocialRestPostCreationData" }, ID: null, creationData: data } }; data.restCreationData.creationData.__metadata = { type: "SP.Social.SocialPostCreationData" }; return this.executePost("feed/post", null, data, true); } },
                    { name: "getFeed", "function": function () { return this.executeGet("feed"); } },
                ]
            },
            // Social Feed Manager
            "SP.Social.SocialRestFeedManager": {
                custom: [
                    { name: "createPost", "function": function (data) { data = { restCreationData: { __metadata: { type: "SP.Social.SocialRestPostCreationData" }, ID: null, creationData: data } }; data.restCreationData.creationData.__metadata = { type: "SP.Social.SocialPostCreationData" }; return this.executePost("my/feed/post", null, data, true); } },
                    { name: "deletePost", "function": function (id) { return this.executePost("post/delete", null, { ID: id }, true); } },
                    { name: "getMyFeed", "function": function () { return this.executeGet("my/feed"); } },
                    { name: "getMyInfo", "function": function () { return this.executeGet("my"); } },
                    { name: "getMyLikes", "function": function () { return this.executeGet("my/likes"); } },
                    { name: "getMyMentionFeed", "function": function () { return this.executeGet("my/mentionfeed"); } },
                    { name: "getMyNews", "function": function () { return this.executeGet("my/news"); } },
                    { name: "getMyTimeLineFeed", "function": function () { return this.executeGet("my/timelinefeed"); } },
                    { name: "getMyUnreadMentionCount", "function": function () { return this.executeGet("my/unreadmentioncount"); } },
                    { name: "getPost", "function": function (id) { return this.executePost("post", null, { ID: id }, true); } },
                    { name: "getPostLikers", "function": function (id) { return this.executePost("post/likers", null, { ID: id }, true); } },
                    { name: "getUser", "function": function (user) { return user.indexOf('i:0#.f|') == 0 ? this.executeGet("actor", null, user) : this.executeGet("actor", user); } },
                    { name: "likePost", "function": function (id) { return this.executePost("post/like", null, { ID: id }, true); } },
                    { name: "lockPost", "function": function (id) { return this.executePost("post/lock", null, { ID: id }, true); } },
                    { name: "replyToPost", "function": function (id, data) { data = { restCreationData: { __metadata: { type: "SP.Social.SocialRestPostCreationData" }, ID: id, creationData: data } }; data.restCreationData.creationData.__metadata = { type: "SP.Social.SocialPostCreationData" }; return this.executePost("post/reply", null, data, true); } },
                    { name: "unlikePost", "function": function (id) { return this.executePost("post/unlike", null, { ID: id }, true); } },
                    { name: "unlockPost", "function": function (id) { return this.executePost("post/unlock", null, { ID: id }, true); } },
                ]
            },
            // Social Thread
            "SP.Social.SocialRestThread": {
                custom: [
                    { name: "delete", "function": function (id) { return this.executePost("delete", null, { ID: id || this.ID }, true); } },
                    { name: "like", "function": function (id) { return this.executePost("like", null, { ID: id || this.ID }, true); } },
                    { name: "lock", "function": function (id) { return this.executePost("lock", null, { ID: id || this.ID }, true); } },
                    { name: "reply", "function": function (data, id) { data = { restCreationData: { __metadata: { type: "SP.Social.SocialRestPostCreationData" }, ID: id || this.ID, creationData: data } }; data.restCreationData.creationData.__metadata = { type: "SP.Social.SocialPostCreationData" }; return this.executePost("reply", null, data, true); } },
                    { name: "unlike", "function": function (id) { return this.executePost("unlike", null, { ID: id || this.ID }, true); } },
                    { name: "unlock", "function": function (id) { return this.executePost("unlock", null, { ID: id || this.ID }, true); } },
                ]
            },
            // User Custom Action
            "SP.UserCustomAction": {
                post: ["deleteObject"]
            },
            // People Manager
            "SP.UserProfiles.PeopleManager": {
                get: ["amlFollowedBy", "amlFollowing", "getFollowedTags", "getFollowersFor", "getMyFollowers", "getMyProperties", "getMySuggestions",
                    "getPeopleFollowedBy", "getPeopleFollowedByMe", "getPropertiesFor", "getUserProfilePropertyFor"],
                getAppendMethodToEndPoint: ["getTrendingTags", "isFollowing"],
                post: ["follow", "followTag", "hideSuggestion", "stopFollowing", "stopFollowingTag"],
                postDataInBody: ["setMyProfilePicture"]
            },
            // Profile Loader
            "SP.UserProfiles.ProfileLoader": {
                post: ["getOwnerUserProfile", "getUserProfile"],
                postDataInBody: ["createPersonalSiteEnqueueBulk"]
            },
            // User Profile
            "SP.UserProfiles.UserProfile": {
                post: ["createPersonalSiteEnque", "shareAllSocialData"],
                custom: [
                    { name: "getOneDriveUrl", "function": function () { return this.FollowPersonalSiteUrl + "_layouts/15/onedrive.aspx" } }
                ]
            },
            // View
            "SP.View": {
                get: ["renderAsHtml"],
                post: ["deleteObject"],
                custom: [
                    { name: "update", "function": function (data) { return this.executePost(null, null, data, true, "SP.View", "MERGE"); } }
                ]
            },
            // View Field Collection
            "SP.ViewFieldCollection": {
                post: ["addViewField", "moveViewFieldTo", "removeAllViewFields", "removeViewField"]
            },
            // Web
            "SP.Web": {
                get: ["doesPushNotificationSubscriberExist", "getAppInstanceById", "getAppInstancesByProductId", "getAvailableWebTemplates",
                    "getCatalog", "getContextWebInformation", "getCustomListTemplates", "getDocumentLibraries", "getFileByServerRelativeUrl",
                    "getFolderByServerRelativeUrl", "getList", "getPushNotificationSubscriber", "getPushNotificationSubscribersByArgs",
                    "getPushNotificationSubscribersByUser", "getSubwebsFilteredForCurrentUser", "getUserById", "getWebUrlFromPageUrl", "mapsToIcon"],
                getDataAsParameter: ["doesUserHavePermissions", "getUserEffectivePermissions"],
                post: ["applyTheme", "applyWebTemplate", "breakRoleInheritance", "deleteObject", "getAppBdcCatalog", "getAppBdcCatalogForAppInstance", "getEntity",
                    "registerPushNotificationSubscriber", "resetRoleInheritance", "unregisterPushNotificationSubscriber"],
                postDataAsParameter: ["addCustomAction", "ensureUser"],
                postDataInBodyNoArgs: ["executeRemoteLOB", "getChanges", "loadAndInstallApp", "loadAndInstallAppInSpecifiedLocale", "loadApp", "processExternalNotification"],
                custom: [
                    { name: "addContentType", "function": function (data) { return this.executePost("contenttypes", null, data, true, "SP.ContentType"); } },
                    { name: "addCustomAction", "function": function (data) { return this.executePost("usercustomactions", null, data, true, "SP.UserCustomAction"); } },
                    { name: "addExistingContentType", "function": function (data) { return this.executePost("contenttypes/addAvailableContentType", data); } },
                    { name: "addField", "function": function (data) { return this.executePost("fields/add", null, data, true, "SP.Field"); } },
                    { name: "addFieldAsXml", "function": function (data) { return this.executePost("fields/createFieldAsXml", null, { parameters: { __metadata: { type: "SP.XmlSchemaFieldCreationInformation" }, Options: SP.AddFieldOptions.addFieldInternalNameHint, SchemaXml: data } }, true); } },
                    { name: "addFile", "function": function (data, content) { return this.executePost("rootfolder/files/add", data, content, true); } },
                    { name: "addList", "function": function (data) { return this.executePost("lists", null, data, true, "SP.List"); } },
                    { name: "addPermission", "function": function (data) { data.__metadata = { type: "SP.RoleDefinition" }; return this.executePost("roledefinitions", null, data, true, "SP.RoleDefinition"); } },
                    { name: "addSiteGroup", "function": function (name) { return this.executePost("sitegroups", null, { Title: name }, true, "SP.Group"); } },
                    { name: "addSubFolder", "function": function (name) { return this.executePost("rootfolder/folders/add", name); } },
                    { name: "addWeb", "function": function (data) { return this.get_Webs().add(data); } },
                    { name: "getContentType", "function": function (title) { title = encodeURIComponent(title); return this.executeGet("contenttypes?$filter=Name eq '" + title + "'"); } },
                    { name: "getContentTypeById", "function": function (id) { return this.executeGet("contenttypes/getById", id); } },
                    { name: "getCustomAction", "function": function (title) { title = encodeURIComponent(title); return this.executeGet("usercustomactions?$filter=Name eq '" + title + "' or Title eq '" + title + "'"); } },
                    { name: "getField", "function": function (title) { title = encodeURIComponent(title); return this.executeGet("fields?$filter=Title eq '" + title + "' or InternalName eq '" + title + "' or StaticName eq '" + title + "'"); } },
                    { name: "getFieldById", "function": function (id) { return this.executeGet("fields/getById", id); } },
                    { name: "getFieldByInternalName", "function": function (name) { name = encodeURIComponent(name); return this.executeGet("fields?$filter=InternalName eq '" + name + "'"); } },
                    { name: "getFieldByStaticName", "function": function (name) { name = encodeURIComponent(name); return this.executeGet("fields?$filter=StaticName eq '" + name + "'"); } },
                    { name: "getFieldByTitle", "function": function (title) { title = encodeURIComponent(title); return this.executeGet("fields?$filter=Title eq '" + title + "'"); } },
                    { name: "getFile", "function": function (name) { name = encodeURIComponent(name); return this.executeGet("rootfolder/files?$filter=Name eq '" + name + "'"); } },
                    { name: "getListById", "function": function (id) { return this.executeGet("lists/getById", id); } },
                    { name: "getListByTitle", "function": function (title) { return this.executeGet("lists/getByTitle", title); } },
                    { name: "getSiteGroupById", "function": function (id) { return this.executeGet("sitegroups/getById", id); } },
                    { name: "getSiteGroupByName", "function": function (name) { return this.executeGet("sitegroups/getByName", name); } },
                    { name: "getSubFolder", "function": function (name) { name = encodeURIComponent(name); return this.executeGet("rootfolder/folders?$filter=Name eq '" + name + "'"); } },
                    { name: "getUserById", "function": function (id) { return this.executeGet("siteusers?$filter=Id eq " + id); } },
                    { name: "getUserByLogin", "function": function (login) { return this.executeGet("siteusers/getByLoginName", null, login); } },
                    { name: "hasAccess", "function": function (permissions) { return hasAccess(this, permissions); } },
                    { name: "sendEmail", "function": function (data) { data = { properties: data }; data.properties.__metadata = { type: "SP.Utilities.EmailProperties" }; return this.executePost("_api/SP.Utilities.Utility.SendEmail", null, data, true); } },
                    { name: "update", "function": function (data) { return this.executePost(null, null, data, true, "SP.Web", "MERGE"); } }
                ]
            },
            // Web Part Manager
            "SP.WebParts.LimitedWebPartManager": {
                custom: [
                    {
                        name: "get_WebParts", "function": function () {
                            return new BRAVO.Core.Object(this.TargetUrl, this.TargetEndPoint + "/WebParts?$expand=WebPart", this.asyncFl);
                        }
                    }
                ]
            },
            // **********************************************************************************
            // By End Point
            // **********************************************************************************
            // Attachment Files
            "attachmentfiles": {
                custom: [
                    {
                        name: "add",
                        "function": function (fileName, content) {
                            return this.executePost("add", { FileName: fileName }, content, true);
                        }
                    },
                    {
                        name: "addFile",
                        "function": function (file) {
                            var thisObj = this;
                            return new Promise(function(resolve, reject){
                                getFileInfo(file).then(function (args) {
                                    var name = args.name, 
                                    buffer = args.buffer;
                                    if (name && buffer) {
                                        thisObj.add(name, buffer).then(function (file) {
                                            resolve(file);
                                        });
                                    } else {
                                        resolve();
                                    }
                                });
                            });
                        }
                    },
                ]
            },
            // Content Type Collection
            "contenttypes": {
                get: ["getById"],
                post: ["addAvailableContentType"],
                postDataInBodyNoArgs: ["add"]
            },
            // Field Collection
            "fields": {
                get: ["getById", "getByInternalNameOrTitle"],
                post: ["addDependentLookup"],
                postDataInBodyNoArgs: ["addField", "createFieldAsXml"]
            },
            // Field Link Collection
            "fieldlinks": {
                get: ["getById"],
                postDataInBodyNoArgs: ["add"]
            },
            // File Collection
            "files": {
                get: ["getByTitle", "getByUrl"],
                postDataInBody: ["add", "addTemplateFile"]
            },
            // File Version Collection
            "versions": {
                get: ["getById"],
                post: ["deleteAll", "deleteById", "deleteByLabel", "restoreByLabel"]
            },
            // Folder Collection
            "folders": {
                get: ["getByUrl"],
                post: ["add"]
            },
            // Group Collection
            "sitegroups": {
                get: ["getById", "getByName"],
                post: ["removeById", "removeByLoginName"]
            },
            // Item Collection
            "items": {
                get: ["getById"]
            },
            // List Collection
            "lists": {
                get: ["getById"],
                post: ["ensureSiteAssetsLibrary", "ensureSitePagesLibrary"]
            },
            // Role Assignment Collection
            "roleassignments": {
                get: ["getByPrincipalId"],
                post: ["addRoleAssignment", "removeRoleAssignment"]
            },
            // Role Definition Collection
            "roledefinitions": {
                get: ["getById", "getByName", "getByType"]
            },
            // User Collection
            "users": {
                get: ["getByEmail", "getById"],
                getDataAsParameter: ["getByLoginName"],
                post: ["removeById"],
                postDataAsParameter: ["removeByLoginName"]
            },
            // User Custom Action Collection
            "usercustomactions": {
                get: ["getById"],
                post: ["clear"]
            },
            // View Collection
            "views": {
                get: ["getById"]
            },
            // Web Collection
            "webs": {
                custom: [
                    { name: "add", "function": function (data) { data = { parameters: data }; data.parameters.__metadata = { type: "SP.WebCreationInformation" }; return this.executePost("add", null, data, true); } },
                ]
            }
        };

        // The User Collection also has the endpoint 'setusers'
        methods["siteusers"] = methods["users"];

        // Return the methods
        return methods;
    };

    // **********************************************************************************
    // Private Methods
    // **********************************************************************************

    // Add Methods
    // This method will add functions based on the type of object.
    var addMethods = function (obj) {
        var metadataType = null;
        var methods = null;
        var restMethods = _restMethods();

        // Set the type
        obj.type = obj.__metadata ? obj.__metadata.type : null;

        // Generic execute get method
        obj["executeGet"] = executeGet;

        // Generic execute post method
        obj["executePost"] = executePost;

        // Generic refresh method
        obj["refresh"] = refresh;

        // Generic set property method
        obj["setProperty"] = setProperty;

        // See if the metadata exists
        if (obj.__metadata) {
            metadataType = obj.__metadata.type;

            // See if this is a field
            if (/^SP\.Field/.test(metadataType)) {
                metadataType = "SP.Field";
            }

            // See if this is a list item
            if (/^SP\.Data\./.test(metadataType) && /Item$/.test(metadataType)) {
                metadataType = "SP.ListItem";
            }

            // Set the methods
            methods = restMethods[metadataType];
        } else {
            // Get the end point
            var endPoint = obj.TargetEndPoint.split('/');
            endPoint = endPoint[endPoint.length - 1].toLowerCase();

            // Set the methods
            methods = restMethods[endPoint];
        }

        // Parse the parameters
        for (var methodType in methods) {
            // Parse the methods
            for (var i = 0; i < methods[methodType].length; i++) {
                var method = methods[methodType][i];

                // Add the method
                switch (methodType) {
                    case "custom":
                        // Add the method
                        obj[method.name] = method.function;
                        break;
                    case "get":
                        // Add the method
                        obj[method] = new Function("args",
                            "return this.executeGet('" + method + "', args);");
                        break;
                    case "getAppendMethodToEndPoint":
                        // Add the method
                        obj[method] = new Function("args",
                            "return this.executeGet(null, args, null, null, null, null, '" + method + "');");
                        break;
                    case "getBuffer":
                        // Add the method
                        obj[method] = new Function("args",
                            "return this.executeGet('" + method + "', args, null, false, true);");
                        break;
                    case "getDataAsParameter":
                        // Add the method
                        obj[method] = new Function("data",
                            "return this.executeGet('" + method + "', null, data);");
                        break;
                    case "getDataInBody":
                        // Add the method
                        obj[method] = new Function("args", "data",
                            "return this.executeGet('" + method + "', args, data, true);");
                        break;
                    case "getDataInBodyNoArgs":
                        // Add the method
                        obj[method] = new Function("data",
                            "return this.executeGet('" + method + "', null, data, true);");
                        break;
                    case "post":
                        // Add the method
                        obj[method] = new Function("args", "data", "type",
                            "return this.executePost('" + method + "', args, data, false, type);");
                        break;
                    case "postAppendMethodToEndPoint":
                        // Add the method
                        obj[method] = new Function("args", "data", "type",
                            "return this.executePost(null, args, data, false, type, null, '" + method + "');");
                        break;
                    case "postDataAsParameter":
                        // Add the method
                        obj[method] = new Function("args", "data", "type",
                            "return this.executePost('" + method + "', args, data, false, type);");
                        break;
                    case "postDataInBody":
                        // Add the method
                        obj[method] = new Function("args", "data", "type",
                            "return this.executePost('" + method + "', args, data, true, type);");
                        break;
                    case "postDataInBodyNoArgs":
                        // Add the method
                        obj[method] = new Function("data", "type",
                            "return this.executePost('" + method + "', null, data, true, type);");
                        break;
                };
            }
        }

        // See if the object contains results
        if (obj.results) {
            // Generic get by id method
            obj["getById"] = function (id) { return getResult(this, ["Id"], id); }

            // Generic get by title method
            obj["getByTitle"] = function (title) { return getResult(this, ["Title", "Name", "InternalName", "StaticName"], title); }

            // Parse the results
            for (var i = 0; i < obj.results.length; i++) {
                // Update the properties of the object
                obj.results[i] = updateProperties({}, obj.results[i], obj);
            }
        }

        // Update the target end point, based on the metadata type
        switch (metadataType) {
            case "SP.Field":
                obj.TargetEndPoint = obj.TargetEndPoint.replace("AvailableFields", "Fields");
                break;
        };
    };

    // Execute Get
    // This is a generic method stub for executing a get request.
    // funcName - The function name.
    // args - The function parameters passed in the function call.
    // data - The function parameters passed as a data object.
    // sendDataInBodyFl - Flag to send the data in the request body.
    // bufferFl - Flag to indicate the response is an array buffer.
    // metadataType - The metadata type.
    // endPoint - The sub-method of the api's endpoint.
    var executeGet = function (funcName, args, data, sendDataInBodyFl, bufferFl, metadataType, endPoint) {
        // See if we are sending the data as the request body
        if (sendDataInBodyFl) { return executeMethod(this, "GET", metadataType, funcName, data, null, bufferFl); }

        // See if we are returning an array buffer
        if (bufferFl) { return executeMethod(this, "GET", metadataType, funcName, data, null, bufferFl); }

        // Update the function name
        funcName = updateFunctionName(this, funcName, args, data);

        // Update the url
        var url = this.TargetEndPoint + (endPoint ? "." + endPoint : "") + (funcName ? "/" + funcName : "");

        // Return a new request
        return new BRAVO.Core.Object(this.TargetUrl, url, this.asyncFl);
    };

    // Execute Method
    // This is a generic method stub for executing requests.
    // obj - This object.
    // method - The method type.
    // metadataType - The metadata type.
    // funcName - The function name.
    // data - The data sent with the xml http request.
    // headers - The request headers.
    // bufferFl - Flag to indicate the response is an array buffer.
    var executeMethod = function (obj, method, metadataType, funcName, data, headers, bufferFl) {
        var endPoint = "";

        // Set the metadata type
        if (metadataType) {
            data["__metadata"] = { "type": metadataType };
        }

        // Update the end point
        if (funcName && funcName.indexOf("_api/") == 0) {
            endPoint = funcName.substr(5);
        }
        else {
            // See if the end point contains a query string
            var idx = obj.TargetEndPoint.indexOf('?');
            var qs = idx > 0 ? obj.TargetEndPoint.substr(idx) : "";
            var url = idx > 0 ? obj.TargetEndPoint.substr(0, idx) : obj.TargetEndPoint;

            // Set the end point
            endPoint = url + (funcName ? "/" + funcName : "") + qs;
        }

        // Set the url
        var url = obj.TargetTemplate.replace("{{EndPoint}}", endPoint).replace("{{Url}}", obj.TargetUrl);

        // Execute the request
        return executeRequest(obj, url, method, data, headers, bufferFl);
    };

    // Execute Post
    // This is a generic method stub for executing a post request.
    // funcName - The function name
    // args - The function parameters passed in the function call.
    // data - The function parameters passed as a data object.
    // sendDataInBodyFl - Flag to send the data in the request body.
    // metadataType - The metadata type.
    // methodType - The method type.
    // endPoint - The sub-method of the api's endpoint.
    var executePost = function (funcName, args, data, sendDataInBodyFl, metadataType, methodType, endPoint) {
        var headers = null;
        var response = null;

        // Update the headers, based on the method
        methodType = methodType ? methodType : "POST";
        if (methodType == "DELETE" || methodType == "MERGE") {
            headers = { "IF-MATCH": "*" };
        }

        // Update the function name
        funcName = updateFunctionName(this, funcName, args, sendDataInBodyFl ? null : data);

        // Update the end point
        if (endPoint) { this.TargetEndPoint += "." + endPoint; }

        // Method to process the request
        var processRequest = function (obj, response) {
            // Parse the response
            response = response ? JSON.parse(response) : response;
            if (response.d) {
                // See if the metadata property exists
                if (response.d.__metadata && response.d.__metadata.uri) {
                    // Set the global variables
                    var uriInfo = getUriInfo(obj, response.d.__metadata.uri);

                    // Return the object
                    return new BRAVO.Core.Object(uriInfo.url, uriInfo.endPoint, obj.asyncFl, response);
                }

                // Return the object
                return new BRAVO.Core.Object(obj.TargetUrl, obj.TargetEndPoint, obj.asyncFl, response);
            }

            // Return the response
            return response;
        };

        // See if we are making an asynchronous request
        if (this.asyncFl) {
            return new Promise(
                function (resolve, reject) {
                    // Execute the request
                    executeMethod(this, methodType, metadataType, funcName, data, headers).then(
                        function (args) {
                            var obj = args.obj,
                                response = args.response;
                            // Process the response and resolve the promise
                            resolve(processRequest(obj, response));
                        }
                    );
                }
            );
        }

        // Execute the request
        response = executeMethod(this, methodType, metadataType, funcName, data, headers);

        // Process the request and return it
        return processRequest(this, response);
    };

    // Execute Request
    // This method will execute a REST url request and return the json result.
    // obj - This object.
    // url - The url for the xml http request.
    // method - "DELETE", "GET", "MERGE", "POST", "PUT"
    // data - The data sent with the xml http request.
    // headers - Additional headers to add to the xml http request.
    // bufferFl - Flag to indicate the response is an array buffer.
    var executeRequest = function (obj, url, method, data, headers, bufferFl) {
        // Ensure the url exists
        if (url) {
            var xhr = getXmlHttpRequest();

            // Ensure the method is set
            method = method ? method : "GET";

            // Open the request
            xhr.open(method == "GET" ? "GET" : "POST", url, obj.asyncFl || bufferFl ? true : false);

            // See if we are returning an array buffer
            if (bufferFl) {
                return new Promise(
                    function (resolve, reject) {
                        // Set the response type
                        xhr.responseType = "arraybuffer";

                        // Set the state change event
                        xhr.onreadystatechange = function () {
                            // See if the request has finished
                            if (xhr.readyState == 4) {
                                // Resolve the promise
                                resolve({
                                    obj: obj,
                                    response: xhr.response
                                });
                            }
                        }

                        // Execute the request
                        xhr.send();
                    }
                );
            }

            // Set the default headers
            xhr.setRequestHeader("Accept", "application/json;odata=verbose");
            xhr.setRequestHeader("Content-Type", "application/json;odata=verbose");
            xhr.setRequestHeader("X-HTTP-Method", method);
            xhr.setRequestHeader("X-RequestDigest", document.querySelector("#__REQUESTDIGEST").value);

            // See if the headers exist
            if (headers) {
                // Parse the headers
                for (var header in headers) {
                    // Add the header
                    xhr.setRequestHeader(header, headers[header]);
                }
            }

            // Stringify the data for the response
            data = data ? (data.byteLength ? data : JSON.stringify(data)) : null;

            // See if we are making an aysnchronous request
            if (obj.asyncFl) {
                return new Promise(
                    function (resolve, reject) {
                        // Set the state change event
                        xhr.onreadystatechange = function () {
                            // See if the request has finished
                            if (xhr.readyState == 4) {
                                // Resolve the promise
                                resolve({
                                    obj: obj,
                                    response: xhr.response
                                });
                            }
                        }

                        // Execute the request
                        xhr.send(data);
                    }
                );
            }

            // Execute the request
            xhr.send(data);

            // Return the response
            return xhr.response;
        }
    };

    // Get Domain Url
    // Method to get the domain url.
    var getDomainUrl = function () {
        var url = document.location.href;

        // See if this is an app web
        if (_spPageContextInfo && _spPageContextInfo.isAppWeb) {
            // Set the url to the host url
            url = BRAVO.Core.getQueryStringValue("SPHostUrl") + "";
        }

        // Split the url and validate it
        url = url.split('/');
        if (url && url.length >= 2) {
            // Set the url
            url = url[0] + "//" + url[2];
        }

        // Return the url
        return url;
    };

    // Method to get the file information.
    // file - An input element w/ its type set to 'file'.
    var getFileInfo = function (file) {
        return new Promise(
            function (resolve, reject) {
                // Set the file
                file = file.files && file.files.length > 0 ? file.files[0] : file;

                // Ensure the file exists
                if (file && file.name) {
                    // Create a file reader
                    var reader = FileReader ? new FileReader() : null;
                    if (reader) {
                        // Event triggered after the file is read successfully
                        reader.onloadend = function (e) {
                            // Resolve the promise
                            resolve({
                                name: file.name,
                                buffer: e.target.result
                            });
                        }

                        // Event triggered on error
                        reader.onerror = function (e) {
                            // Resolve the promise
                            resolve(e.target.error);
                        }

                        // Read the file as an array buffer
                        reader.readAsArrayBuffer(file);
                    }
                    else {
                        // Resolve the promise
                        resolve();
                    }
                }
                else {
                    // Resolve the promise
                    resolve();
                }
            }
        );
    }

    // Get Query String Value
    // This method will get the query string value by the specified key.
    // key - The query string key.
    var getQueryStringValue = function (key) {
        // Get the query string
        var queryString = document.location.href.split('?');
        queryString = queryString.length > 1 ? queryString[1] : queryString[0];

        // Parse the values
        var values = queryString.split('&');
        for (var i = 0; i < values.length; i++) {
            var keyValue = values[i].split('=');
            if (keyValue.length == 1) { continue; }

            // See if this is the key we are looking for
            if (decodeURIComponent(keyValue[0]) == key) {
                return decodeURIComponent(keyValue[1]);
            }
        }

        // Key was not found
        return null;
    };

    // Get Result
    // This method will search the results and find the it by a specifiec property name and value
    var getResult = function (obj, propertyNames, value) {
        // Ensure results and the properties exist
        if (obj.results == null || obj.results.length == 0) { return null; }

        // Parse the property names
        for (var i = 0; i < propertyNames.length; i++) {
            var propertyName = propertyNames[i];

            // Ensure the result contains this property
            if (obj.results[0][propertyName] == null) { continue; }

            // Parse the results
            for (var j = 0; j < obj.results.length; j++) {
                var result = obj.results[j];
                var propertyValue = result[propertyName];

                // See if this is a webpart
                if (result.__metadata && result.__metadata.type == "SP.WebParts.WebPartDefinition") {
                    // Update the property name
                    propertyValue = result.WebPart[propertyName];
                }

                // See if this is the result we are looking for
                if (propertyValue != null && propertyValue.toLowerCase() == value.toLowerCase()) { return result; }
            }
        }

        // Result not found
        return null;
    };

    // Get URI Info
    // This method will extract the url and end point from the uri.
    // obj - The object to get the uri from.
    // uri - The absolute url of the request.
    var getUriInfo = function (obj, uri) {
        var endPoint = "";

        // Get the end point for the target web
        var uriInfo = uri.split("/_api/SP.AppContextSite(@target)/");

        // See if the uri is for the current web
        if (uriInfo.length == 1) { uriInfo = uri.split("/_api/"); }

        // Set the end point
        if (uriInfo.length == 1) {
            // Set the end point
            endPoint = uriInfo[0].split('/');
            endPoint = (obj.TargetEndPoint ? obj.TargetEndPoint + "/" : "") + endPoint[endPoint.length - 1];
        } else {
            // Set the end point
            endPoint = uriInfo[1];
        }

        // Return the uri info
        return { endPoint: endPoint, url: obj.TargetUrl };
    };

    // Get XML Http Request
    // This method will create the xml http request object
    var getXmlHttpRequest = function () {
        if (typeof XMLHttpRequest === "undefined") {
            // Try to create the request
            try { return new ActiveXObject("Msxml2.XMLHTTP.6.0"); }
            catch (e) { }
            try { return new ActiveXObject("Msxml2.XMLHTTP.3.0"); }
            catch (e) { }
            try { return new ActiveXObject("Microsoft.XMLHTTP"); }
            catch (e) { }

            // Throw an error
            throw new Error("This browser does not support XMLHttpRequest.");
        }

        // Create an instance of the xml http request object
        return new XMLHttpRequest();
    };

    // Has Access
    // This method will determine if the user has the specified permissions.
    // obj - The web or site object.
    // permissions - The permissions to check.
    // userName - The user name of the check access for.
    var hasAccess = function (obj, permissions, userName) {
        var hasPermissionFl = false;

        // Store the current setting, and disable the asynchronous flag
        var isAsync = obj.asyncFl;
        obj.asyncFl = false;

        // See if this is a site
        if (obj.__metadata.type == "SP.Site") {
            var rootWebUrl = null;

            // See if this is an app web
            if (_spPageContextInfo.isAppWeb) {
                // Set the root web url
                rootWebUrl = getDomainUrl() + obj.get_RootWeb().ServerRelativeUrl;
            } else {
                // Set the root web url
                rootWebUrl = _spPageContextInfo.siteAbsoluteUrl;
            }

            // Check the permissions against the root web
            hasPermissionFl = (new BRAVO.Core.Web(rootWebUrl)).hasAccess(permissions);
        }
        else {
            // See if this is a web
            if (obj.__metadata.type == "SP.Web") {
                // Set the user name
                userName = userName || obj.get_CurrentUser().LoginName;
            }

            // Get the permissions of the current user
            var userPermissions = obj.getUserEffectivePermissions(userName);
            if (userPermissions.exists) {
                // Default the permission flag
                hasPermissionFl = true;

                // Set the permissions
                var basePermissions = new SP.BasePermissions();
                basePermissions.initPropertiesFromJson(userPermissions.GetUserEffectivePermissions);

                // Parse the permissions
                for (var i = 0; i < permissions.length; i++) {
                    // Determine if the user has the specified permission
                    hasPermissionFl &= basePermissions.has(permissions[i]);
                }
            }
        }

        // Reset the asynchronous flag
        obj.asyncFl = isAsync;

        // Return the flag
        return hasPermissionFl ? true : false;
    };

    // Method to load the dependencies
    // callback - The method to execute, after the dependencies are loaded.
    var loadDependencies = function (callback) {
        // See if the page context exists
        if (window._spPageContextInfo) {
            // Execute the callback function and return
            if (callback && typeof (callback) === "function") { callback(); }
            return;
        }

        // Wait for the window to be loaded
        window.addEventListener("load", function () {
            // See if the page context exists
            if (window._spPageContextInfo) {
                // Execute the callback function and return
                if (callback && typeof (callback) === "function") { callback(); }
                return;
            }

            // Ensure the dependencies have been loaded
            if (!_dependenciesLoadedFl) {
                // Set the flag
                _dependenciesLoadedFl = true;

                // Parse the scripts to load
                ["MicrosoftAjax.js", "init.js", "sp.runtime.js", "sp.js", "sp.core.js", "core.js"].every(function (fileName) {
                    // Create the script element
                    var el = document.createElement("script");

                    // Set the properties
                    el.setAttribute("src", "/_layouts/15/" + fileName);
                    el.setAttribute("type", "text/javascript");

                    // Add it to the head
                    document.head.appendChild(el);
                });
            }

            // Keep looping until the page context exists
            var counter = 0;
            var maxLoops = 400;
            var intervalId = window.setInterval(function () {
                // Wait for the page context to exist
                if (window._spPageContextInfo) {
                    // Stop the loop
                    window.clearInterval(intervalId);

                    // Execute the callback function
                    if (callback && typeof (callback) === "function") { callback(); }
                }

                // See if we have hit the maximum # of tries
                if (++counter >= maxLoops) {
                    // Stop the loop
                    window.clearInterval(intervalId);
                }
            }, 25);
        });
    };

    // Refresh
    // This method will execute the same request
    var refresh = function () {
        // Copy the target information
        var obj = {
            asyncFl: this.asyncFl,
            RequestUrl: this.RequestUrl,
            TargetEndPoint: this.TargetEndPoint,
            TargetTemplate: this.TargetTemplate,
            TargetUrl: this.TargetUrl
        };

        // See if we are making an asynchronous request
        if (obj.asyncFl) {
            return new Promise(
                function (resolve, reject) {
                    // Execute the request
                    executeRequest(obj, obj.RequestUrl).then(function (args) {
                        var obj = args.obj,
                            response = args.response;
                        // Set the response
                        obj.Response = JSON.parse(response);

                        // Update the properties
                        updateProperties(obj, obj.Response.d);

                        // Resolve the promise
                        resolve(obj);
                    });
                }
            );
        }

        // Execute the request
        obj.Response = JSON.parse(executeRequest(obj, obj.RequestUrl));

        // Update the properties
        updateProperties(obj, obj.Response.d);

        // Return the object
        return obj;
    };

    // Set Property
    // This method will update the property of the current object.
    // name - The name of the property.
    // value - The value to set the property to.
    var setProperty = function (name, value) {
        var promise;

        // Ensure the value exists, and determine if an update is required
        if (value === undefined && this[name] != null && this[name] == value) {
            // Resolve the promise
            promise = new Promise(function (resolve) {
                resolve(true);
            });
        }
        else {
            // Create the data to update the property
            var data = '{ "' + name + '": ' + (typeof (value) == "string" ? '"' + value + '"' : value) + ' }';

            // Method called after the request is made
            var postRequest = function (obj, response) {
                // See if the update was successful
                if (response == "") {
                    // Update the property in the current object
                    obj[name] = value;

                    // Return the promise, if we are using asynchronous requests
                    return true;
                }

                // Request was not successful
                return false;
            };

            // Execute the request
            var response = executeMethod(this, "MERGE", this.__metadata.type, null, JSON.parse(data), { "IF-MATCH": "*" });

            // See if this is an asynchronous request
            if (response.then) {
                // Wait for the response to complete
                promise = new Promise(
                    function (resolve) {
                        response.then(function (args) {
                            var obj = args.obj,
                                response = args.response;
                            // Resolve the promise
                            resolve(postRequest(obj, response));
                        });
                    }
                );
            }
            else {
                return postRequest(this, response);
            }
        }

        // Return the promise if we are using asynchronous requests
        return this.asyncFl ? promise : false;
    };

    // Update Function Name
    // This method will add the arguments or data object to the function name
    // obj - The current object.
    // functionName - The function name.
    // args - The function parameters passed in the function call.
    // data - The function parameters passed as a data object.
    var updateFunctionName = function (obj, functionName, args, data) {
        // Ensure the function name exists
        if (functionName) {
            var encodeUri = functionName.indexOf("?$filter") > 0;

            // See if the args exist
            if (args) {
                var argType = typeof (args);

                // Append the arguments to the function
                functionName += "(";

                // Set the function name, based on the argument type
                switch (argType) {
                    case "boolean":
                        functionName += args ? "true" : "false";
                        break;
                    case "string":
                        functionName += "'" + (encodeUri ? encodeURIComponent(args) : args) + "'";
                        break;
                    case "object":
                        // Parse the arguments
                        for (var key in args) {
                            var value = args[key];
                            value = typeof (value) === "string" ? "'" + value + "'" : value;

                            // Add the argument to the function
                            functionName += key + "=" + (encodeUri ? encodeURIComponent(value) : value) + ", ";
                        }
                        break;
                    default:
                        functionName += args;
                        break;
                }

                // Add the closing ')' and return the value
                return (functionName.endsWith(", ") ? functionName.substr(0, functionName.length - 2) : functionName) + ")";
            }

            // See if data exists
            if (data) {
                // Encode data if it's a string
                data = typeof (data) === "string" ? "'" + encodeURIComponent(data) + "'" : JSON.stringify(data);

                // Append the data to the function
                functionName += "(" + (obj.__metadata.type.indexOf("SP.Social.") == 0 ? "item=" : "") + "@v)?@v=" + data;
            }
        }

        // Return the function name
        return functionName;
    };

    // Update Properties
    // This method will read a json object and apply the properties/methods to the this object.
    // obj - This object, to update.
    // data - The JSON data to parse and copy to the object.
    // parent - The parent object, used for analyzing a collection.
    var updateProperties = function (obj, data, parent) {
        // Set the exists flag
        obj.exists = data && data.error == null;

        // Validate the input parameters
        if (data) {
            var addResultArrayFl = false;

            // See if the parent exists
            if (parent) {
                // Set the target information
                obj.asyncFl = parent.asyncFl;
                obj.TargetEndPoint = data.__metadata && data.__metadata.uri ? data.__metadata.uri.split("/_api/")[1] : parent.targetEndPoint;
                obj.TargetTemplate = parent.TargetTemplate;
                obj.TargetUrl = parent.TargetUrl;
            }

            // See if this is a collection
            if (data.results) {
                // See if this is a query
                if (obj.TargetEndPoint && obj.TargetEndPoint.indexOf("?$filter")) {
                    // Update exists flag
                    obj.exists = data.results.length > 0;

                    // See if only 1 result exists
                    if (data.results.length == 1) {
                        // Set the flag to add the result array
                        addResultArrayFl = true;

                        // Set the data to the result
                        data = data.results[0];

                        // Update the target end point and template
                        var uriInfo = data.__metadata.uri.split("/_api/");
                        obj.TargetEndPoint = uriInfo[1];
                        obj.TargetTemplate = obj.TargetTemplate.replace(/&@target/, "?@target");
                    }
                }
            }

            // Parse the properties of the object
            for (var propName in data) {
                // Ensure the property has a value
                if (data[propName] == null) { obj[propName] = data[propName]; continue; }

                // See if this is a method
                if (data[propName].__deferred && data[propName].__deferred.uri) {
                    var uriInfo = getUriInfo(obj, data[propName].__deferred.uri);

                    // Generate the method
                    obj["get_" + propName] = new Function("return new BRAVO.Core.Object(\"" + uriInfo.url + "\", \"" + uriInfo.endPoint + "\", this.asyncFl);");
                }
                // Else, add the property
                else { obj[propName] = data[propName]; }
            }

            // Add the methods
            addMethods(obj);

            // See if we are adding the results array
            if (addResultArrayFl) {
                // Retain the 'results' property, so the developer can access it
                obj.results = obj.results ? obj.results : [obj];
            }
        }

        // Return the object
        return obj;
    };

    // **********************************************************************************
    // Public Interface
    // **********************************************************************************
    return {
        // Get query string value
        getQueryStringValue: getQueryStringValue,

        // Load the dependencies
        loadDependencies: loadDependencies,

        // Core Object
        // Takes the following input parameters:
        // string, string - The host url and end point.
        // string, string, boolean - The host url, end point, and asynchronous request flag
        // string, string, boolean, object - The host url, end point, asynchronous request flag and the json results object.
        // string, string, boolean, string - The host url, end point, asynchronous request flag and method type of the api.
        Object: function () {
            var obj = new function () { };

            // Determine how to create the object, based on the input parameters
            if (arguments.length > 1) {
                // Set the global variables
                obj.TargetEndPoint = arguments[1];
                obj.TargetUrl = arguments[0];
                obj.TargetTemplate =
                    (_spPageContextInfo.isAppWeb ? _spPageContextInfo.webAbsoluteUrl : obj.TargetUrl) + "/_api/" +
                    (_spPageContextInfo.isAppWeb ? "SP.AppContextSite(@target)/" : "") +
                    "{{EndPoint}}" +
                    (_spPageContextInfo.isAppWeb ? (obj.TargetEndPoint.indexOf('?') > 0 ? "&" : "?") + "@target='{{Url}}'" : "");
                obj.RequestUrl = obj.TargetTemplate
                    .replace("{{EndPoint}}", obj.TargetEndPoint)
                    .replace("{{Url}}", obj.TargetUrl);
            }

            // See if the request url is set
            if (obj.RequestUrl) {
                // Set the asynchronous flag
                obj.asyncFl = arguments[2] && typeof (arguments[2]) === "boolean" ? arguments[2] : false;

                // See if a result has been passed
                if (arguments.length == 4 && typeof (arguments[3]) === "object") {
                    // Set the response
                    obj.Response = arguments[3];

                    // Update the properties
                    updateProperties(obj, obj.Response.d);
                }
                else {
                    // Determine the method type
                    var methodType = arguments[3] && typeof (arguments[3] === "string") ? arguments[3] : "GET";

                    // See if we are making an asynchronous request
                    if (obj.asyncFl) {
                        return new Promise(function (resolve, reject) {
                            // Execute the request
                            executeRequest(obj, obj.RequestUrl, methodType).then(function (args) {
                                var obj = args.obj,
                                response = args.response;
                                // Set the response
                                obj.Response = JSON.parse(response);

                                // Update the properties
                                updateProperties(obj, obj.Response.d);

                                // Resolve the promise
                                resolve(obj);
                            });
                        });
                    }

                    // Execute the request
                    obj.Response = JSON.parse(executeRequest(obj, obj.RequestUrl, methodType));

                    // Update the properties
                    updateProperties(obj, obj.Response.d);
                }
            }

            // Return this object
            return obj;
        },

        // Promise
        // Promise: function () {
        //     return {
        //         _arguments: null,
        //         _callback: null,
        //         _resolveFl: false,
        //         done: function (callback) { this._callback = callback; if (this._callback && this._resolveFl) { this._callback.apply(this, this._arguments); } },
        //         resolve: function () { this._arguments = arguments; this._resolveFl = true; if (this._callback) { this._callback.apply(this, this._arguments); } }
        //     };
        // },

        // **********************************************************************************
        // SP Objects
        // **********************************************************************************

        // List
        // listName - The name of the list.
        // hostUrl - The url to the web.
        // asyncFl - Flag to determine if the requests are asynchronous.
        List: function (listName, hostUrl, asyncFl) {
            // Encode the list name
            listName = encodeURIComponent(listName);

            // Create the site
            return hostUrl ?
                new BRAVO.Core.Object(hostUrl.indexOf("http") == 0 ? hostUrl : getDomainUrl() + hostUrl, "web/lists?$filter=Title eq '" + listName + "'", asyncFl) :
                window._spPageContextInfo ? new BRAVO.Core.Object(window._spPageContextInfo.webAbsoluteUrl, "web/lists?$filter=Title eq '" + listName + "'", asyncFl) : null;
        },

        // Asynchronous List
        // listName - The name of the list.
        // hostUrl - The url to the web.
        ListAsync: function (listName, hostUrl) { return new BRAVO.Core.List(listName, hostUrl, true); },

        // People Manager
        // hostUrl - The url to the web.
        // asyncFl - Flag to determine if the requests are asynchronous.
        PeopleManager: function (hostUrl, asyncFl) {
            // Create the site
            return hostUrl ?
                new BRAVO.Core.Object(hostUrl.indexOf("http") == 0 ? hostUrl : getDomainUrl() + hostUrl, "sp.userprofiles.peoplemanager", asyncFl) :
                window._spPageContextInfo ? new BRAVO.Core.Object(window._spPageContextInfo.siteAbsoluteUrl, "sp.userprofiles.peoplemanager", asyncFl) : null;
        },

        // Asynchronous People Manager
        // hostUrl - The url to the web.
        PeopleManagerAsync: function (hostUrl) { return new BRAVO.Core.PeopleManager(hostUrl, true); },

        // Profile Loader
        // hostUrl - The url to the web.
        // asyncFl - Flag to determine if the requests are asynchronous.
        ProfileLoader: function (hostUrl, asyncFl) {
            // Create the site
            return hostUrl ?
                new BRAVO.Core.Object(hostUrl.indexOf("http") == 0 ? hostUrl : getDomainUrl() + hostUrl, "sp.userprofiles.profileloader.getprofileloader", asyncFl, "POST") :
                window._spPageContextInfo ? new BRAVO.Core.Object(window._spPageContextInfo.siteAbsoluteUrl, "sp.userprofiles.profileloader.getprofileloader", asyncFl, "POST") : null;
        },

        // Asynchronous Profile Loader
        // hostUrl - The url to the web.
        ProfileLoaderAsync: function (hostUrl) { return new BRAVO.Core.ProfileLoader(hostUrl, true); },

        // Search
        // hostUrl - The url to the web.
        // asyncFl - Flag to determine if the requests are asynchronous.
        Search: function (hostUrl, asyncFl) {
            // Create the site
            return hostUrl ?
                new BRAVO.Core.Object(hostUrl.indexOf("http") == 0 ? hostUrl : getDomainUrl() + hostUrl, "search", asyncFl) :
                window._spPageContextInfo ? new BRAVO.Core.Object(window._spPageContextInfo.siteAbsoluteUrl, "search", asyncFl) : null;
        },

        // Asynchronous Search
        // hostUrl - The url to the web.
        SearchAsync: function (hostUrl) { return new BRAVO.Core.Search(hostUrl, true); },

        // Social Manager
        // hostUrl - The url to the web.
        // asyncFl - Flag to determine if the requests are asynchronous.
        SocialManager: function (hostUrl, asyncFl) {
            // Create the site
            return hostUrl ?
                new BRAVO.Core.Object(hostUrl.indexOf("http") == 0 ? hostUrl : getDomainUrl() + hostUrl, "social.feed", asyncFl) :
                window._spPageContextInfo ? new BRAVO.Core.Object(window._spPageContextInfo.siteAbsoluteUrl, "social.feed", asyncFl) : null;
        },

        // Asynchronous Social Manager
        // hostUrl - The url to the web.
        SocialManagerAsync: function (hostUrl) { return new BRAVO.Core.SocialManager(hostUrl, true); },

        // Site
        // hostUrl - The url to the web.
        // asyncFl - Flag to determine if the requests are asynchronous.
        Site: function (hostUrl, asyncFl) {
            // Create the site
            return hostUrl ?
                new BRAVO.Core.Object(hostUrl.indexOf("http") == 0 ? hostUrl : getDomainUrl() + hostUrl, "site", asyncFl) :
                window._spPageContextInfo ? new BRAVO.Core.Object(window._spPageContextInfo.siteAbsoluteUrl, "site", asyncFl) : null;
        },

        // Asynchronous Site
        // hostUrl - The url to the web.
        SiteAsync: function (hostUrl) { return new BRAVO.Core.Site(hostUrl, true); },

        Utility: {
            createEmailBodyForInvitation: function () { "POST" },
            getCurrentUserEmailAddresses: function () { return new BRAVO.Core.Object(window._spPageContextInfo.siteAbsoluteUrl, "sp.utilities.utility.getCurrentUserEmailAddresses") }
        },

        // Web
        // hostUrl - The url to the web.
        // asyncFl - Flag to determine if the requests are asynchronous.
        Web: function (hostUrl, asyncFl) {
            // Create the web
            return hostUrl ?
                new BRAVO.Core.Object(hostUrl.indexOf("http") == 0 ? hostUrl : getDomainUrl() + hostUrl, "web", asyncFl) :
                window._spPageContextInfo ? new BRAVO.Core.Object(window._spPageContextInfo.webAbsoluteUrl, "web", asyncFl) : null;
        },

        // Asynchronous Web
        // hostUrl - The url to the web.
        WebAsync: function (hostUrl) { return new BRAVO.Core.Web(hostUrl, true); }
    };
}();

// The help interface
// objectType - [Required] The object type to get help for.
// methodName - [Optional] The method name to get help for.
BRAVO.Help = function (objectType, methodName) {

    // **********************************************************************************
    // Private Variables
    // **********************************************************************************

    var _line = "************************************************************************";
    var _lineDashes = "------------------------------------------------------------------------";
    var _methodType = { Get: 0, Post: 1 };
    var _sampleBoolean = "true";
    var _sampleContentTypeId = "0x0101000728167cd9c94899925ba69c4a601720";
    var _sampleId = "{7B926655-E840-484C-91F5-6017201E9DD6}";
    var _sampleNumber = "0";
    var _sampleText = "'Title'";

    // **********************************************************************************
    // Object Help Information
    // **********************************************************************************
    var _objInfo = {

        // **********************************************************************************
        // By Metadata Type
        // **********************************************************************************

        // Content Type
        "contenttype": {
            properties: {
                "description": { description: "Gets or sets a description of the content type.", name: "description", readOnly: true },
                "displayformtemplatename": { description: "Gets or sets a value that specifies the name of a custom display form template to use for list items that have been assigned the content type.", name: "displayFormTemplateName", readOnly: true },
                "displayformurl": { description: "Gets or sets a value that specifies the URL of a custom display form to use for list items that have been assigned the content type.", name: "displayFormUrl", readOnly: true },
                "documenttemplate": { description: "Gets or sets a value that specifies the file path to the document template used for a new list item that has been assigned the content type.", name: "documentTemplate", readOnly: true },
                "documenttemplateurl": { description: "Gets a value that specifies the URL of the document template assigned to the content type.", name: "documentTemplateUrl", readOnly: false },
                "editformtemplatename": { description: "Gets or sets a value that specifies the name of a custom edit form template to use for list items that have been assigned the content type.", name: "editFormTemplateName", readOnly: true },
                "editformurl": { description: "Gets or sets a value that specifies the URL of a custom edit form to use for list items that have been assigned the content type.", name: "editFormUrl", readOnly: true },
                "fieldlinks": { description: "Gets the column (also known as field) references in the content type.", name: "fieldLinks", methodName: "get_FieldLinks" },
                "fields": { description: "Gets a value that specifies the collection of fields for the content type.", name: "fields", methodName: "get_Fields" },
                "group": { description: "Gets or sets a value that specifies the content type group for the content type.", name: "group", readOnly: true },
                "hidden": { description: "Gets or sets a value that specifies whether the content type is unavailable for creation or usage directly from a user interface.", name: "hidden", readOnly: true },
                "id": { description: "Gets a value that specifies an identifier for the content type.", name: "id", readOnly: false },
                "jslink": { description: "Gets or sets the JSLink for the content type custom form template.", name: "jsLink", readOnly: true },
                "name": { description: "Gets or sets a value that specifies the name of the content type.", name: "name", readOnly: true },
                "newformtemplatename": { description: "Gets or sets a value that specifies the name of the content type.", name: "newFormTemplateName", readOnly: true },
                "newformurl": { description: "Gets or sets a value that specifies the name of the content type.", name: "newFormUrl", readOnly: true },
                "parent": { description: "Gets the parent content type of the content type.", name: "parent", readOnly: false },
                "readonly": { description: "Gets or sets a value that specifies whether changes to the content type properties are denied.", name: "readOnly", readOnly: true },
                "schemaxml": { description: "Gets a value that specifies the XML Schema representing the content type.", name: "schemaXml", readOnly: false },
                "schemaxmlwithresourcetokens": { description: "Gets a non-localized version of the XML schema that defines the content type.", name: "schemaXmlWithResourceTokens", readOnly: false },
                "scope": { description: "Gets a value that specifies a server-relative path to the content type scope of the content type.", name: "scope", readOnly: false },
                "sealed": { description: "Gets or sets whether the content type can be modified.", name: "sealed", readOnly: true },
                "stringid": { description: "A string representation of the value of the Id.", name: "stringId", readOnly: false },
                "workflowassociations": { description: "Gets a value that specifies the collection of workflow associations for the content type.", name: "workflowAssociations", readOnly: false },
            },
            methods: {
                addfieldlink: {
                    name: "addFieldLink",
                    description: "This method will add a field link to the collection.",
                    parameters: {
                        body: [
                            {
                                name: "FieldInternalName",
                                description: "The internal field name property.",
                                sampleValue: _sampleText
                            },
                            {
                                name: "Hidden",
                                description: "Specifies whether the field is displayed in forms.",
                                sampleValue: _sampleBoolean
                            },
                            {
                                name: "Required",
                                description: "Specifies whether the field requires a value.",
                                sampleValue: _sampleBoolean
                            }
                        ]
                    },
                    type: _methodType.Post
                },
                deleteobject: {
                    name: "deleteObject",
                    description: "The method will delete the content type.",
                    type: _methodType.Post
                },
                getfieldbyinternalname: {
                    name: "getFieldByInternalName",
                    description: "This method will return a SP.Field object, by the internal name property.",
                    parameters: {
                        name: "internalName",
                        description: "A string value, representing the internal name of the field.",
                        sampleValue: _sampleText
                    },
                    type: _methodType.Get
                },
                getfieldbystaticname: {
                    name: "getFieldByStaticName",
                    description: "This method will return a SP.Field object, by static name property.",
                    parameters: {
                        name: "staticName",
                        description: "A string value, representing the internal name of the field.",
                        sampleValue: _sampleText
                    },
                    type: _methodType.Get
                },
                getfieldbytitle: {
                    name: "getFieldByTitle",
                    description: "This method will return a SP.Field object, by title property.",
                    parameters: {
                        name: "title",
                        description: "A string value, representing the internal name of the field.",
                        sampleValue: _sampleText
                    },
                    type: _methodType.Get
                },
                getfieldlinkbyname: {
                    name: "getFieldLinkByName",
                    description: "This method will return a SP.FieldLink link object, by the name property.",
                    parameters: {
                        name: "name",
                        description: "A string value, representing the internal name of the field.",
                        sampleValue: _sampleText
                    },
                    type: _methodType.Get
                },
                update: {
                    name: "update",
                    description: "This method will update the content type properties.",
                    type: _methodType.Post
                }
            }
        },
        // Content Types
        "contenttypes": {
            methods: {
                add: {
                    name: "add",
                    description: "This method will add a new content type to the collection.",
                    parameters: [
                        {
                            name: "Description",
                            description: "A string value, representing the description of the content type.",
                            sampleValue: _sampleText
                        },
                        {
                            name: "Group",
                            description: "A string value, representing the group to associate the content type to.",
                            sampleValue: _sampleText
                        },
                        {
                            name: "Name",
                            description: "A string value, representing the name of the content type.",
                            sampleValue: _sampleText
                        },
                        {
                            name: "ParentContentType",
                            description: "A string value, representing the parent of the content type.",
                            sampleValue: _sampleContentTypeId
                        }
                    ],
                    type: _methodType.Post
                },
                addAvailableContentType: {
                    name: "addAvailableContentType",
                    description: "",
                    parameters: [
                        {
                            name: "contentTypeId",
                            description: "A string value, representing the content type id.",
                            sampleValue: _sampleContentTypeId
                        }
                    ],
                    type: _methodType.Post
                },
                getById: {
                    name: "getById",
                    description: "A string value, representing the content type id.",
                    parameters: [
                        {
                            name: "contentTypeId",
                            description: "A string value, representing the content type id.",
                            sampleValue: "0x0101000728167cd9c94899925ba69c4a601720"
                        }
                    ],
                    type: _methodType.Get
                }
            }
        },
        // Field
        "field": {
            properties: {
                "CanBeDeleted": { description: "Gets a value that specifies whether the field can be deleted.", name: "CanBeDeleted", readOnly: false },
                "DefaultValue": { description: "Gets or sets a value that specifies the default value for the field.", name: "DefaultValue", readOnly: true },
                "Description": { description: "Gets or sets a value that specifies the description of the field.", name: "Description", readOnly: true },
                "Direction": { description: "Gets or sets a value that specifies the reading order of the field.", name: "Direction", readOnly: true },
                "EnforceUniqueValues": { description: "Gets or sets a value that specifies whether to require unique field values in a list or library column.", name: "EnforceUniqueValues", readOnly: true },
                "EntityPropertyName": { description: "Gets the name of the entity property for the list item entity that uses this field.", name: "EntityPropertyName", readOnly: false },
                "FieldTypeKind": { description: "Gets or sets a value that specifies the type of the field. Represents a FieldType value. See FieldType in the .NET client object model reference for a list of field type values.", name: "FieldTypeKind", readOnly: true },
                "Filterable": { description: "Gets a value that specifies whether list items in the list can be filtered by the field value.", name: "Filterable", readOnly: false },
                "FromBaseType": { description: "Gets a Boolean value that indicates whether the field derives from a base field type.", name: "FromBaseType", readOnly: false },
                "Group": { description: "Gets or sets a value that specifies the field group.", name: "Group", readOnly: true },
                "Hidden": { description: "Gets or sets a value that specifies whether the field is hidden in list views and list forms.", name: "Hidden", readOnly: true },
                "Id": { description: "Gets a value that specifies the field identifier.", name: "Id", readOnly: false },
                "Indexed": { description: "Gets or sets a Boolean value that specifies whether the field is indexed.", name: "Indexed", readOnly: true },
                "InternalName": { description: "Gets a value that specifies the field internal name.", name: "InternalName", readOnly: false },
                "JSLink": { description: "Gets or sets the name of an external JS file containing any client rendering logic for fields of this type.", name: "JSLink", readOnly: true },
                "ReadOnlyField": { description: "Gets or sets a value that specifies whether the value of the field is read-only.", name: "ReadOnlyField", readOnly: true },
                "Required": { description: "Gets or sets a value that specifies whether the field requires a value.", name: "Required", readOnly: true },
                "SchemaXml": { description: "Gets or sets a value that specifies the XML schema that defines the field.", name: "SchemaXml", readOnly: true },
                "SchemaXmlWithResourceTokens": { description: "Gets the schema that defines the field and includes resource tokens.", name: "SchemaXmlWithResourceTokens", readOnly: false },
                "Scope": { description: "Gets a value that specifies the server-relative URL of the list or the site to which the field belongs.", name: "Scope", readOnly: true },
                "Sealed": { description: "Gets a value that specifies whether properties on the field cannot be changed and whether the field cannot be deleted.", name: "Sealed", readOnly: true },
                "Sortable": { description: "Gets a value that specifies whether list items in the list can be sorted by the field value.", name: "Sortable", readOnly: true },
                "StaticName": { description: "Gets or sets a value that specifies a customizable identifier of the field.", name: "StaticName", readOnly: true },
                "Title": { description: "Gets or sets value that specifies the display name of the field.", name: "Title", readOnly: true },
                "TypeAsString": { description: "Gets or sets a value that specifies the type of the field.", name: "TypeAsString", readOnly: true },
                "TypeDisplayName": { description: "Gets a value that specifies the display name for the type of the field.", name: "TypeDisplayName", readOnly: false },
                "TypeShortDescription": { description: "Gets a value that specifies the description for the type of the field.", name: "TypeShortDescription", readOnly: false },
                "ValidationFormula": { description: "Gets or sets a value that specifies the data validation criteria for the value of the field.", name: "ValidationFormula", readOnly: true },
                "ValidationMessage": { description: "Gets or sets a value that specifies the error message returned when data validation fails for the field.", name: "ValidationMessage", readOnly: true },
            },
            methods: {
                "deleteObject": {
                    name: "deleteObject",
                    description: "This method will delete the field.",
                    parameters: {},
                    type: _methodType.Post
                },
                "setShowInDisplayForm": {
                    name: "setShowInDisplayForm",
                    description: "This method will update the ShowInDisplayForm property of the field.",
                    parameters: {},
                    type: _methodType.Post
                },
                "setShowInEditForm": {
                    name: "setShowInEditForm",
                    description: "This method will update the ShowInEditForm property of the field.",
                    parameters: {},
                    type: _methodType.Post
                },
                "setShowInNewForm": {
                    name: "setShowInNewForm",
                    description: "This method will update the ShowInNewForm property of the field.",
                    parameters: {},
                    type: _methodType.Post
                },
                "update": {
                    name: "update",
                    description: "This method will update the field properties.",
                    parameters: {},
                    type: _methodType.Post
                }
            }
        },
        // Fields
        "fields": {
            properties: {
                "schemaxml": { description: "Specifies the XML schema of the collection of fields.", name: "SchemaXml", readOnly: true },
            },
            methods: {}
        },
        // Field Link
        "fieldlink": {
            properties: {
                "hidden": { description: "Gets or sets a value that specifies whether the field is displayed in forms that can be edited.", name: "hidden", readOnly: true },
                "id": { description: "Gets a value that specifies the GUID of the FieldLink.", name: "id", readOnly: false },
                "name": { description: "Gets a value that specifies the name of the FieldLink.", name: "name", readOnly: false },
                "required": { description: "Gets or sets a value that specifies whether the field (2) requires a value.", name: "required", readOnly: true },
            },
            methods: {}
        },
        // Field Links
        "fieldlinks": {
            methods: {}
        },
        // File
        "file": {
            properties: {
                "author": { description: "Gets a value that specifies the user who added the file.", name: "Author", readOnly: true },
                "checkedoutbyuser": { description: "Gets a value that returns the user who has checked out the file.", name: "CheckedOutByUser", readOnly: true },
                "checkincomment": { description: "Gets a value that returns the comment used when a document is checked in to a document library.", name: "CheckInComment", readOnly: true },
                "checkouttype": { description: "Gets a value that indicates how the file is checked out of a document library. Represents an SP.CheckOutType value: Online = 0; Offline = 1; None = 2.", name: "CheckOutType", readOnly: true },
                "contenttag": { description: "Returns internal version of content, used to validate document equality for read purposes.", name: "ContentTag", readOnly: true },
                "customizedpagestatus": { description: "Gets a value that specifies the customization status of the file. Represents an SP.CustomizedPageStatus value: None = 0; Uncustomized = 1; Customized = 2.", name: "CustomizedPageStatus", readOnly: true },
                "etag": { description: "Gets a value that specifies the ETag value.", name: "ETag", readOnly: true },
                "exists": { description: "Gets a value that specifies whether the file exists.", name: "Exists", readOnly: true },
                "length": { description: "Gets the size of the file in bytes, excluding the size of any Web Parts that are used in the file.", name: "Length", readOnly: true },
                "level": { description: "Gets a value that specifies the publishing level of the file. Represents an SP.FileLevel value: Published = 1; Draft = 2; Checkout = 255.", name: "Level", readOnly: true },
                "listitemallfields": { description: "Gets a value that specifies the list item field values for the list item corresponding to the file.", name: "ListItemAllFields", readOnly: true },
                "lockedbyuser": { description: "Gets a value that returns the user that owns the current lock on the file.", name: "LockedByUser", readOnly: true },
                "majorversion": { description: "Gets a value that specifies the major version of the file.", name: "MajorVersion", readOnly: true },
                "minorversion": { description: "Gets a value that specifies the minor version of the file.", name: "MinorVersion", readOnly: true },
                "modifiedby": { description: "Gets a value that returns the user who last modified the file.", name: "ModifiedBy", readOnly: true },
                "name": { description: "Gets the name of the file including the extension.", name: "Name", readOnly: true },
                "serverrelativeurl": { description: "Gets the relative URL of the file based on the URL for the server.", name: "ServerRelativeUrl", readOnly: true },
                "timecreated": { description: "Gets a value that specifies when the file was created.", name: "TimeCreated", readOnly: true },
                "timelastmodified": { description: "Gets a value that specifies when the file was last modified.", name: "TimeLastModified", readOnly: true },
                "title": { description: "Gets a value that specifies the display name of the file.", name: "Title", readOnly: true },
                "uiversion": { description: "Gets a value that specifies the implementation-specific version identifier of the file.", name: "UiVersion", readOnly: true },
                "uiversionlabel": { description: "Gets a value that specifies the implementation-specific version identifier of the file.", name: "UiVersionLabel", readOnly: true },
                "versions": { description: "Gets a value that returns a collection of file version objects that represent the versions of the file.", name: "Versions", methodName: "get_Versions" },
            },
            methods: {}
        },
        // Files
        "files": {
            methods: {}
        },
        // File Version
        "fileversion": {
            properties: {
                "checkincomment": { description: "Gets a value that specifies the check-in comment.", name: "CheckInComment", readOnly: true },
                "created": { description: "Gets a value that specifies the creation date and time for the file version.", name: "Created", readOnly: true },
                "createdby": { description: "Gets a value that specifies the user that represents the creator of the file version.", name: "CreatedBy", readOnly: true },
                "id": { description: "Gets the internal identifier for the file version.", name: "ID", readOnly: true },
                "iscurrentversion": { description: "Gets a value that specifies whether the file version is the current version.", name: "IsCurrentVersion", readOnly: true },
                "size": { description: "", name: "Size", readOnly: true },
                "url": { description: "Gets a value that specifies the relative URL of the file version based on the URL for the site or subsite.", name: "Url", readOnly: true },
                "versionlabel": { description: "Gets a value that specifies the implementation specific identifier of the file. Uses the majorVersionNumber.minorVersionNumber format, for example: 1.2.", name: "VersionLabel", readOnly: true },
            },
            methods: {}
        },
        // File Versions
        "fileversions": {
            methods: {}
        },
        // Folder
        "folder": {
            properties: {
                "contenttypeorder": { description: "Specifies the sequence in which content types are displayed.", name: "ContentTypeOrder", readOnly: true },
                "files": { description: "Gets the collection of all files contained in the list folder. You can add a file to a folder by using the Add method on the folder’s FileCollection resource.", name: "Files", readOnly: true },
                "folders": { description: "Gets the collection of list folders contained in the list folder.", name: "Folders", readOnly: true },
                "itemcount": { description: "Gets a value that specifies the count of items in the list folder.", name: "ItemCount", readOnly: true },
                "listitemallfields": { description: "Specifies the list item field (2) values for the list item corresponding to the file.", name: "ListItemAllFields", methodName: "get_ListItemAllFields" },
                "name": { description: "Gets the name of the folder.", name: "Name", readOnly: true },
                "parentfolder": { description: "Gets the parent list folder of the folder.", name: "ParentFolder", readOnly: true },
                "properties": { description: "Gets the collection of all files contained in the folder.", name: "Properties", methodName: "get_Properties" },
                "serverrelativeurl": { description: "Gets the server-relative URL of the list folder.", name: "ServerRelativeUrl", readOnly: true },
                "uniquecontenttypeorder": { description: "Gets or sets a value that specifies the content type order.", name: "UniqueContentTypeOrder", readOnly: false },
                "welcomepage": { description: "Gets or sets a value that specifies folder-relative URL for the list folder welcome page.", name: "WelcomePage", readOnly: false },
            },
            methods: {}
        },
        // Folders
        "folders": {
            methods: {}
        },
        // Group
        "group": {
            properties: {
                "allowmemberseditmembership": { description: "Gets or sets a value that indicates whether the group members can edit membership in the group.", name: "AllowMembersEditMembership", readOnly: false },
                "allowrequesttojoinleave": { description: "Gets or sets a value that indicates whether to allow users to request membership in the group and request to leave the group.", name: "AllowRequestToJoinLeave", readOnly: false },
                "autoacceptrequesttojoinleave": { description: "Gets or sets a value that indicates whether the request to join or leave the group can be accepted automatically.", name: "AutoAcceptRequestToJoinLeave", readOnly: false },
                "cancurrentusereditmembership": { description: "Gets a value that indicates whether the current user can edit the membership of the group.", name: "CanCurrentUserEditMembership", readOnly: true },
                "cancurrentusermanagegroup": { description: "Gets a value that indicates whether the current user can manage the group.", name: "CanCurrentUserManageGroup", readOnly: true },
                "cancurrentuserviewmembership": { description: "Gets a value that indicates whether the current user can view the membership of the group.", name: "CanCurrentUserViewMembership", readOnly: true },
                "description": { description: "Gets or sets the description of the group.", name: "Description", readOnly: false },
                "id": { description: "Gets a value that specifies the member identifier for the user or group.", name: "Id", readOnly: true },
                "ishiddeninui": { description: "Gets a value that indicates whether this member should be hidden in the UI.", name: "IsHiddenInUI", readOnly: true },
                "loginname": { description: "Gets the name of the group.", name: "LoginName", readOnly: true },
                "onlyallowmembersviewmembership": { description: "Gets or sets a value that indicates whether only group members are allowed to view the membership of the group.", name: "OnlyAllowMembersViewMembership", readOnly: false },
                "owner": { description: "Gets or sets the owner of the group which can be a user or another group assigned permissions to control security.", name: "Owner", readOnly: false },
                "ownertitle": { description: "Gets the name for the owner of this group.", name: "OwnerTitle", readOnly: true },
                "requesttojoinleaveemailsetting": { description: "Gets or sets the email address to which the requests of the membership are sent.", name: "RequestToJoinLeaveEmailSetting", readOnly: false },
                "principaltype": { description: "Gets a value containing the type of the principal. Represents a bitwise SP.PrincipalType value: None = 0; User = 1; DistributionList = 2; SecurityGroup = 4; SharePointGroup = 8; All = 15.", name: "PrincipalType", readOnly: true },
                "title": { description: "Gets or sets a value that specifies the name of the principal.", name: "Title", readOnly: false },
                "users": { description: "Gets a collection of user objects that represents all of the users in the group.", name: "Users", methodName: "get_Users" },
            },
            methods: {}
        },
        // Groups
        "groups": {
            methods: {}
        },
        // Limited Web Part Manager
        "limitedwebpartmanager": {
            properties: {
                "haspersonalizedparts": { description: "Gets a value that indicates whether the page contains one or more personalized Web Parts.", name: "hasPersonalizedParts", readOnly: false },
                "scope": { description: "Gets a value that specifies the current personalization scope of the Web Part Page.", name: "scope", readOnly: false },
                "webparts": { description: "Gets a value that specifies collection of the Web Parts on the Web Part Page available to the current user based on the current user's permissions.", name: "webParts", methodName: "get_WebParts" },
            },
            methods: {}
        },
        // List
        "list": {
            properties: {
                "allowcontenttypes": { description: "Gets a value that specifies whether the list supports content types.", name: "AllowContentTypes", readOnly: true },
                "basetemplate": { description: "Gets the list definition type on which the list is based. Represents a ListTemplateType value. See ListTemplateType in the .NET client object model reference for template type values.", name: "BaseTemplate", readOnly: true },
                "basetype": { description: "Gets the base type for the list. Represents an SP.BaseType value: Generic List = 0; Document Library = 1; Discussion Board = 3; Survey = 4; Issue = 5.", name: "BaseType", readOnly: true },
                "browserfilehandling": { description: "Gets a value that specifies the override of the web application's BrowserFileHandling property at the list level. Represents an SP.BrowserFileHandling value: Permissive = 0; Strict = 1.", name: "BrowserFileHandling", readOnly: true },
                "contenttypes": { description: "Gets the content types that are associated with the list.", name: "ContentTypes", methodName: "get_ContentTypes" },
                "contenttypesenabled": { description: "Gets or sets a value that specifies whether content types are enabled for the list.", name: "ContentTypesEnabled", readOnly: false },
                "created": { description: "Gets a value that specifies when the list was created.", name: "Created", readOnly: true },
                "datasource": { description: "Gets the data source associated with the list, or null if the list is not a virtual list. Returns null if the HasExternalDataSource property is false.", name: "DataSource", readOnly: true },
                "defaultcontentapprovalworkflowid": { description: "Gets or sets a value that specifies the default workflow identifier for content approval on the list. Returns an empty GUID if there is no default content approval workflow.", name: "DefaultContentApprovalWorkflowId", readOnly: false },
                "defaultdisplayformurl": { description: "Gets or sets a value that specifies the location of the default display form for the list. Clients specify a server-relative URL, and the server returns a site-relative URL", name: "DefaultDisplayFormUrl", readOnly: false },
                "defaulteditformurl": { description: "Gets or sets a value that specifies the URL of the edit form to use for list items in the list. Clients specify a server-relative URL, and the server returns a site-relative URL.", name: "DefaultEditFormUrl", readOnly: false },
                "defaultnewformurl": { description: "Gets or sets a value that specifies the location of the default new form for the list. Clients specify a server-relative URL, and the server returns a site-relative URL.", name: "DefaultNewFormUrl", readOnly: false },
                "defaultview": { description: "", name: "DefaultView", readOnly: true },
                "defaultviewurl": { description: "Gets the URL of the default view for the list.", name: "DefaultViewUrl", readOnly: true },
                "description": { description: "Gets or sets a value that specifies the description of the list.", name: "Description", readOnly: false },
                "direction": { description: "Gets or sets a value that specifies the reading order of the list. Returns 'NONE', 'LTR', or 'RTL'.", name: "Direction", readOnly: false },
                "documenttemplateurl": { description: "Gets or sets a value that specifies the server-relative URL of the document template for the list. Returns a server-relative URL if the base type is DocumentLibrary, otherwise returns null.", name: "DocumentTemplateUrl", readOnly: false },
                "draftversionvisibility": { description: "Gets or sets a value that specifies the minimum permission required to view minor versions and drafts within the list. Represents an SP.DraftVisibilityType value: Reader = 0; Author = 1; Approver = 2.", name: "DraftVersionVisibility", readOnly: false },
                "effectivebasepermissions": { description: "Gets a value that specifies the effective permissions on the list that are assigned to the current user.", name: "EffectiveBasePermissions", methodName: "get_EffectiveBasePermissions" },
                "effectivebasepermissionsforui": { description: "", name: "EffectiveBasePermissionsForUI", readOnly: true },
                "enableattachments": { description: "Gets or sets a value that specifies whether list item attachments are enabled for the list.", name: "EnableAttachments", readOnly: false },
                "enablefoldercreation": { description: "Gets or sets a value that specifies whether new list folders can be added to the list.", name: "EnableFolderCreation", readOnly: false },
                "enableminorversions": { description: "Gets or sets a value that specifies whether minor versions are enabled for the list.", name: "EnableMinorVersions", readOnly: false },
                "enablemoderation": { description: "Gets or sets a value that specifies whether content approval is enabled for the list.", name: "EnableModeration", readOnly: false },
                "enableversioning": { description: "Gets or sets a value that specifies whether historical versions of list items and documents can be created in the list.", name: "EnableVersioning", readOnly: false },
                "entitytypename": { description: "", name: "EntityTypeName", readOnly: true },
                "eventreceivers": { description: "", name: "EventReceivers", methodName: "get_EventReceivers" },
                "fields": { description: "Gets a value that specifies the collection of all fields in the list.", name: "Fields", methodName: "get_Fields" },
                "firstuniqueancestorsecurableobject": { description: "Gets the object where role assignments for this object are defined. If role assignments are defined directly on the current object, the current object is returned.", name: "FirstUniqueAncestorSecurableObject", readOnly: true },
                "forcecheckout": { description: "Gets or sets a value that indicates whether forced checkout is enabled for the document library.", name: "ForceCheckout", readOnly: false },
                "forms": { description: "Gets a value that specifies the collection of all list forms in the list.", name: "Forms", methodName: "get_Forms" },
                "hasexternaldatasource": { description: "Gets a value that specifies whether the list is an external list.", name: "HasExternalDataSource", readOnly: true },
                "hasuniqueroleassignments": { description: "Gets a value that specifies whether the role assignments are uniquely defined for this securable object or inherited from a parent securable object.", name: "HasUniqueRoleAssignments", readOnly: true },
                "hidden": { description: "Gets or sets a Boolean value that specifies whether the list is hidden. If true, the server sets the OnQuickLaunch property to false.", name: "Hidden", readOnly: false },
                "id": { description: "Gets the GUID that identifies the list in the database.", name: "Id", readOnly: true },
                "imageurl": { description: "Gets a value that specifies the URI for the icon of the list.", name: "ImageUrl", readOnly: true },
                "informationrightsmanagementsettings": { description: "", name: "InformationRightsManagementSettings", readOnly: true },
                "irmenabled": { description: "", name: "IrmEnabled", readOnly: false },
                "irmexpire": { description: "", name: "IrmExpire", readOnly: false },
                "irmreject": { description: "", name: "IrmReject", readOnly: false },
                "isapplicationlist": { description: "Gets or sets a value that specifies a flag that a client application can use to determine whether to display the list.", name: "IsApplicationList", readOnly: false },
                "iscatalog": { description: "Gets a value that specifies whether the list is a gallery.", name: "IsCatalog", readOnly: true },
                "isprivate": { description: "", name: "IsPrivate", readOnly: true },
                "issiteassetslibrary": { description: "Gets a value that indicates whether the list is designated as a default asset location for images or other files which the users upload to their wiki pages.", name: "IsSiteAssetsLibrary", readOnly: true },
                "itemcount": { description: "Gets a value that specifies the number of list items in the list.", name: "ItemCount", readOnly: true },
                "items": { description: "Gets all the items in the list.", name: "Items", methodName: "get_Items" },
                "lastitemdeleteddate": { description: "Gets a value that specifies the last time a list item was deleted from the list.", name: "LastItemDeletedDate", readOnly: true },
                "lastitemmodifieddate": { description: "Gets a value that specifies the last time a list item, field, or property of the list was modified.", name: "LastItemModifiedDate", readOnly: false },
                "listitementitytypefullname": { description: "", name: "ListItemEntityTypeFullName", readOnly: true },
                "multipledatalist": { description: "Gets or sets a value that indicates whether the list in a Meeting Workspace site contains data for multiple meeting instances within the site.", name: "MultipleDataList", readOnly: false },
                "nocrawl": { description: "Gets or sets a value that specifies that the crawler must not crawl the list.", name: "NoCrawl", readOnly: false },
                "onquicklaunch": { description: "Gets or sets a value that specifies whether the list appears on the Quick Launch of the site. If true, the server sets the Hidden property to false.", name: "OnQuickLaunch", readOnly: false },
                "parentweb": { description: "Gets a value that specifies the site that contains the list.", name: "ParentWeb", readOnly: true },
                "parentweburl": { description: "Gets a value that specifies the server-relative URL of the site that contains the list.", name: "ParentWebUrl", readOnly: true },
                "roleassignments": { description: "Gets the role assignments for the securable object.", name: "RoleAssignments", methodName: "get_RoleAssignments" },
                "rootfolder": { description: "Gets the root folder that contains the files in the list and any related files.", name: "RootFolder", readOnly: true },
                "schemaxml": { description: "Gets a value that specifies the list schema of the list.", name: "SchemaXml", readOnly: true },
                "servertemplatecancreatefolders": { description: "Gets a value that indicates whether folders can be created within the list.", name: "ServerTemplateCanCreateFolders", readOnly: true },
                "templatefeatureid": { description: "Gets a value that specifies the feature identifier of the feature that contains the list schema for the list. Returns an empty GUID if the list schema is not contained within a feature.", name: "TemplateFeatureId", readOnly: true },
                "title": { description: "Gets or sets the displayed title for the list. Its length must be <= 255 characters.", name: "Title", readOnly: false },
                "usercustomactions": { description: "Gets a value that specifies the collection of all user custom actions for the list.", name: "UserCustomActions", readOnly: true },
                "validationformula": { description: "Gets or sets a value that specifies the data validation criteria for a list item. Its length must be <= 1023.", name: "ValidationFormula", readOnly: false },
                "validationmessage": { description: "Gets or sets a value that specifies the error message returned when data validation fails for a list item. Its length must be <= 1023.", name: "ValidationMessage", readOnly: false },
                "views": { description: "Gets a value that specifies the collection of all public views on the list and personal views of the current user on the list.", name: "Views", methodName: "get_Views" },
                "workflowassociations": { description: "Gets a value that specifies the collection of all workflow associations for the list.", name: "WorkflowAssociations", methodName: "get_WorkflowAssociations" },
            },
            methods: {}
        },
        // Lists
        "lists": {
            methods: {}
        },
        // List Item
        "listitem": {
            properties: {
                "attachmentfiles": { description: "Specifies the collection of attachments that are associated with the list item.", name: "AttachmentFiles", methodName: "get_AttachmentFiles" },
                "contenttype": { description: "Gets a value that specifies the content type of the list item.", name: "ContentType", readOnly: true },
                "displayname": { description: "Gets a value that specifies the display name of the list item.", name: "DisplayName", readOnly: true },
                "effectivebasepermissions": { description: "Gets a value that specifies the effective permissions on the list item that are assigned to the current user.", name: "EffectiveBasePermissions", readOnly: true },
                "effectivebasepermissionsforui": { description: "Gets the effective base permissions for the current user, as they should be displayed in UI.", name: "EffectiveBasePermissionsForUI", readOnly: true },
                "fieldvaluesashtml": { description: "Gets the values for the list item as HTML.", name: "FieldValuesAsHtml", methodName: "get_FieldValuesAsHtml" },
                "fieldvaluesastext": { description: "Gets the list item's field values as a collection of string values.", name: "FieldValuesAsText", methodName: "get_FieldValuesAsText" },
                "fieldvaluesforedit": { description: "Gets the formatted values to be displayed in an edit form.", name: "FieldValuesForEdit", methodName: "get_FieldValuesForEdit" },
                "file": { description: "Gets the file that is represented by the item from a document library.", name: "File", readOnly: true },
                "filesystemobjecttype": { description: "Gets a value that specifies whether the list item is a file or a list folder. Represents an SP.FileSystemObjectType value: Invalid = -1; File = 0; Folder = 1; Web = 2.", name: "FileSystemObjectType", readOnly: true },
                "firstuniqueancestorsecurableobject": { description: "Gets the object where role assignments for this object are defined. If role assignments are defined directly on the current object, the current object is returned.", name: "FirstUniqueAncestorSecurableObject", readOnly: true },
                "folder": { description: "Gets a folder object that is associated with a folder item.", name: "Folder", readOnly: true },
                "hasuniqueroleassignments": { description: "Gets a value that specifies whether the role assignments are uniquely defined for this securable object or inherited from a parent securable object.", name: "HasUniqueRoleAssignments", readOnly: true },
                "id": { description: "Gets a value that specifies the list item identifier.", name: "Id", readOnly: true },
                "parentlist": { description: "Gets the parent list that contains the list item.", name: "ParentList", readOnly: true },
                "roleassignments": { description: "Gets the role assignments for the securable object.", name: "RoleAssignments", methodName: "get_RoleAssignments" },
            },
            methods: {}
        },
        // List Items
        "listitems": {
            methods: {}
        },
        // Role Assignment
        "roleassignment": {
            properties: {
                "member": { description: "Gets the user or group that corresponds to the Role Assignment.", name: "Member", readOnly: true },
                "principalid": { description: "The unique identifier of the role assignment.", name: "PrincipalId", readOnly: true },
                "roledefinitionbindings": { description: "Gets the collection of role definition bindings for the role assignment.", name: "RoleDefinitionBindings", methodName: "get_RoleDefinitionBindings" },
            },
            methods: {}
        },
        // Role Assignments
        "roleassignments": {
            methods: {}
        },
        // Role Definition
        "roledefinition": {
            properties: {
                "basepermissions": { description: "Gets or sets a value that specifies the base permissions for the role definition.", name: "BasePermissions", readOnly: false },
                "description": { description: "Gets or sets a value that specifies the description of the role definition.", name: "Description", readOnly: false },
                "hidden": { description: "Gets a value that specifies whether the role definition is displayed.", name: "Hidden", readOnly: true },
                "id": { description: "Gets a value that specifies the Id of the role definition.", name: "Id", readOnly: true },
                "name": { description: "Gets or sets a value that specifies the role definition name.", name: "Name", readOnly: false },
                "order": { description: "Gets or sets a value that specifies the order position of the object in the site collection Permission Levels page.", name: "Order", readOnly: false },
                "roletypekind": { description: "Gets a value that specifies the type of the role definition. Represents an SP.RoleType value. See RoleType in the .NET client object model reference for a list of role type", name: "RoleTypeKind", readOnly: true },
            },
            methods: {}
        },
        // Role Definitions
        "roledefinitions": {
            methods: {}
        },
        // Search Service
        "searchservice": {
            properties: {},
            methods: {}
        },
        // Site
        "site": {
            properties: {
                "allowdesigner": { description: "Gets or sets a value that specifies whether a designer can be used on this site collection.", name: "allowDesigner", readOnly: true },
                "allowmasterpageediting": { description: "Gets or sets a value that specifies whether master page editing is allowed on this site collection.", name: "allowMasterPageEditing", readOnly: true },
                "allowrevertfromtemplate": { description: "Gets or sets a value that specifies whether this site collection can be reverted to its base template.", name: "allowRevertFromTemplate", readOnly: true },
                "allowselfserviceupgrade": { description: "Whether version to version upgrade is allowed on this site.", name: "allowSelfServiceUpgrade", readOnly: false },
                "allowselfserviceupgradeevaluation": { description: "Whether upgrade evaluation site collection is allowed.", name: "allowSelfServiceUpgradeEvaluation", readOnly: false },
                "canupgrade": { description: "Property indicating whether or not this object can be upgraded.", name: "canUpgrade", readOnly: false },
                "compatibilitylevel": { description: "Gets the major version of this site collection for purposes of major version-level compatibility checks.", name: "compatibilityLevel", readOnly: false },
                "eventreceivers": { description: "Provides event receivers for events that occur at the scope of the site collection.", name: "eventReceivers", methodName: "get_EventReceivers" },
                "features": { description: "Gets a value that specifies the collection of the site collection features for the site collection that contains the site.", name: "features", readOnly: false },
                "id": { description: "Gets the GUID that identifies the site collection.", name: "id", readOnly: false },
                "lockissue": { description: "Gets or sets the comment that is used in locking a site collection.", name: "lockIssue", readOnly: true },
                "maxitemsperthrottledoperation": { description: "Gets a value that specifies the maximum number of list items allowed per operation before throttling will occur.", name: "maxItemsPerThrottledOperation", readOnly: false },
                "owner": { description: "Gets or sets the owner of the site collection. (Read-only in sandboxed solutions.)", name: "owner", readOnly: true },
                "primaryuri": { description: "Specifies the primary URI of this site collection, including the host name, port number, and path.", name: "primaryUri", readOnly: false },
                "readonly": { description: "Gets or sets a Boolean value that specifies whether the site collection is read-only, locked, and unavailable for write access.", name: "readOnly", readOnly: true },
                "recyclebin": { description: "Gets a value that specifies the collection of recycle bin items for the site collection.", name: "recycleBin", readOnly: false },
                "rootweb": { description: "Gets a value that returns the top-level site of the site collection.", name: "rootWeb", readOnly: false },
                "serverrelativeurl": { description: "Gets the server-relative URL of the root Web site in the site collection.", name: "serverRelativeUrl", readOnly: false },
                "sharebylinkenabled": { description: "Property that indicates whether users will be able to share links to documents that can be accessed without logging in.", name: "shareByLinkEnabled", readOnly: false },
                "showurlstructure": { description: "Gets or sets a value that specifies whether the URL structure of this site collection is viewable.", name: "showUrlStructure", readOnly: true },
                "uiversionconfigurationenabled": { description: "Gets or sets a value that specifies whether the Visual Upgrade UI of this site collection is displayed.", name: "uiVersionConfigurationEnabled", readOnly: true },
                "upgradeinfo": { description: "Specifies the upgrade information of this site collection.", name: "upgradeInfo", readOnly: false },
                "upgradereminderdate": { description: "Specifies a date, after which site collection administrators will be reminded to upgrade the site collection.", name: "upgradeReminderDate", readOnly: false },
                "upgrading": { description: "Specifies whether the site is currently upgrading.", name: "upgrading", readOnly: false },
                "url": { description: "Gets the full URL to the root Web site of the site collection, including host name, port number, and path.", name: "url", readOnly: false },
                "usage": { description: "Gets a value that specifies usage information about the site, including bandwidth, storage, and the number of visits to the site collection.", name: "usage", readOnly: false },
                "usercustomactions": { description: "Gets a value that specifies the collection of user custom actions for the site collection.", name: "userCustomActions", methodName: "get_UserCustomActions" },
            },
            methods: {}
        },
        // Social User
        "socialrestactor": {
            properties: {
                "accountname": { description: "Gets the actor's account name. Applies to users.", name: "accountName", readOnly: false },
                "actortype": { description: "Gets the type of actor (user, document, site, or tag).", name: "actorType", readOnly: false },
                "canfollow": { description: "Gets a value that indicates whether the actor can be followed.", name: "canFollow", readOnly: false },
                "contenturi": { description: "Gets the actor's content URI. Applies to documents and sites.", name: "contentUri", readOnly: false },
                "emailaddress": { description: "Gets the actor's email address. Applies to users.", name: "emailAddress", readOnly: false },
                "followedcontenturi": { description: "Gets the URI of the actor's list of followed content. Applies to users.", name: "followedContentUri", readOnly: false },
                "id": { description: "Gets the actor's unique identifier.", name: "id", readOnly: false },
                "imageuri": { description: "Gets the actor's image URI. Applies to users, documents, and sites.", name: "imageUri", readOnly: false },
                "isfollowed": { description: "Gets a value that indicates whether the current user is being followed.", name: "isFollowed", readOnly: false },
                "libraryuri": { description: "Gets the actor's library URI. Applies to documents.", name: "libraryUri", readOnly: false },
                "name": { description: "Gets the actor's display name.", name: "name", readOnly: false },
                "personalsiteuri": { description: "Gets the URI of the actor's personal site. Applies to users.", name: "personalSiteUri", readOnly: false },
                "status": { description: "Gets a code that indicates recoverable errors that occurred during the actor's retrieval.", name: "status", readOnly: false },
                "statustext": { description: "Gets the text of the actor's most recent post. Applies to users.", name: "statusText", readOnly: false },
                "tagguid": { description: "Gets the actor's tag GUID. Applies to tags.", name: "tagGuid", readOnly: false },
                "title": { description: "Gets the actor's title. Applies to users.", name: "title", readOnly: false },
                "typeid": { description: "This member is reserved for internal use and is not intended to be used directly from your code.", name: "typeId", readOnly: false },
                "uri": { description: "Gets the actor's canonical URI.", name: "uri", readOnly: false },
            },
            methods: {}
        },
        // Social Feed Manager
        "socialrestfeedmanager": {
            properties: {
                "owner": { description: "Gets a SocialActor object that represents the current user.", name: "owner", readOnly: false },
                "personalsiteportaluri": { description: "Gets the URI of the default personal site portal for the current user.", name: "personalSitePortalUri", readOnly: false },
            },
            methods: {}
        },
        // Social Thread
        "socialrestthread": {
            properties: {
                "actors": { description: "The merged array of participating actors.", name: "Actors", methodName: "get_Actors" },
                "attributes": { description: "The bitwise value that represents the set of attributes for the thread.", name: "Attributes", methodName: "get_Attributes" },
                "id": { description: "The unique identifier of the thread.", name: "Id", readOnly: false },
                "ownerindex": { description: "The index of the thread's owner within the thread's actors.", name: "OwnerIndex", readOnly: false },
                "permalink": { description: "The string representation of the stable URI for navigating directly to the thread, if one is available.", name: "Permalink", readOnly: false },
                "postreference": { description: "The referenced post.", name: "PostReference", readOnly: false },
                "replies": { description: "The replies to the thread.", name: "Replies", methodName: "get_Replies" },
                "rootpost": { description: "The root post of the thread.", name: "RootPost", readOnly: false },
                "status": { description: "The code that identifies recoverable errors that occurred during thread retrieval. See SP.Social.SocialStatusCode.", name: "Status", readOnly: false },
                "threadtype": { description: "The thread type.", name: "ThreadType", readOnly: false },
                "totalreplycount": { description: "The count of the total number of replies for the thread.", name: "TotalReplyCount", readOnly: false },
            },
            methods: {}
        },
        // People Manager
        "peoplemanager": {
            properties: {
                "editprofilelink": { description: "The URL of the edit profile page for the current user.", name: "EditProfileLink", readOnly: true },
                "ismypeoplelistpublic": { description: "A Boolean value that indicates whether the current user's People I'm Following list is public.", name: "IsMyPeopleListPublic", readOnly: true },
            },
            methods: {}
        },
        // Profile Loader
        "profileloader": {
            methods: {}
        },
        // User Custom Action
        "usercustomaction": {
            properties: {
                "commanduiextension": { description: "Gets or sets a value that specifies an implementation specific XML fragment that determines user interface properties of the custom action.", name: "CommandUIExtension", readOnly: false },
                "description": { description: "Gets or sets the description of the custom action.", name: "Description", readOnly: false },
                "group": { description: "Gets or sets a value that specifies an implementation-specific value that determines the position of the custom action in the page.", name: "Group", readOnly: false },
                "id": { description: "Gets a value that specifies the identifier of the custom action.", name: "Id", readOnly: true },
                "imageurl": { description: "Gets or sets the URL of the image associated with the custom action.", name: "ImageUrl", readOnly: false },
                "location": { description: "Gets or sets the location of the custom action.", name: "Location", readOnly: false },
                "name": { description: "Gets or sets the name of the custom action.", name: "Name", readOnly: false },
                "registrationid": { description: "Gets or sets the value that specifies the identifier of the object associated with the custom action.", name: "RegistrationId", readOnly: false },
                "registrationtype": { description: "Gets or sets the value that specifies the type of object associated with the custom action. Represents an SP.UserCustomActionRegistrationType value: None = 0; List = 1; ContentType = 2; ProgId = 3; FileType = 4.", name: "RegistrationType", readOnly: false },
                "rights": { description: "Gets or sets the value that specifies the permissions needed for the custom action.", name: "Rights", readOnly: false },
                "scope": { description: "Gets a value that specifies the scope of the custom action.", name: "Scope", readOnly: true },
                "scriptblock": { description: "Gets or sets the value that specifies the ECMAScript to be executed when the custom action is performed.", name: "ScriptBlock", readOnly: false },
                "scriptsrc": { description: "Gets or sets a value that specifies the URI of a file which contains the ECMAScript to execute on the page.", name: "ScriptSrc", readOnly: false },
                "sequence": { description: "Gets or sets the value that specifies an implementation-specific value that determines the order of the custom action that appears on the page.", name: "Sequence", readOnly: false },
                "title": { description: "Gets or sets the display title of the custom action.", name: "Title", readOnly: false },
                "url": { description: "Gets or sets the URL, URI, or ECMAScript (JScript, JavaScript) function associated with the action.", name: "Url", readOnly: false },
                "versionofusercustomaction": { description: "Gets a value that specifies an implementation specific version identifier.", name: "VersionOfUserCustomAction", readOnly: true },
            },
            methods: {}
        },
        // User Custom Actions
        "usercustomactions": {
            methods: {}
        },
        // User Profile
        "userprofile": {
            properties: {
                "followedcontent": { description: "An object containing the user's FollowedDocumentsUrl and FollowedSitesUrl.", name: "FollowedContent", readOnly: true },
                "accountname": { description: "The account name of the user. (SharePoint Online only)", name: "AccountName", readOnly: true },
                "displayname": { description: "The display name of the user. (SharePoint Online only)", name: "DisplayName", readOnly: true },
                "o15firstrunexperience": { description: "The FirstRun flag of the user. (SharePoint Online only)", name: "O15FirstRunExperience", readOnly: true },
                "personalsite": { description: "The personal site of the user.", name: "PersonalSite", readOnly: true },
                "personalsitecapabilities": { description: "The capabilities of the user's personal site. Represents a bitwise PersonalSiteCapabilities value: None = 0; Profile Value = 1; Social Value = 2; Storage Value = 4; MyTasksDashboard Value = 8; Education Value = 16; Guest Value = 32.", name: "PersonalSiteCapabilities", readOnly: true },
                "personalsitefirstcreationerror": { description: "The error thrown when the user's personal site was first created, if any. (SharePoint Online only)", name: "PersonalSiteFirstCreationError", readOnly: true },
                "personalsitefirstcreationtime": { description: "The date and time when the user's personal site was first created. (SharePoint Online only)", name: "PersonalSiteFirstCreationTime", readOnly: true },
                "personalsiteinstantiationstate": { description: "The status for the state of the personal site instantiation. See PersonalSiteInstantiationState in the .NET client object model reference for a list of instantiation state values.", name: "PersonalSiteInstantiationState", readOnly: true },
                "personalsitelastcreationtime": { description: "The date and time when the user's personal site was last created. (SharePoint Online only)", name: "PersonalSiteLastCreationTime", readOnly: true },
                "personalsitenumberofretries": { description: "The number of attempts made to create the user's personal site. (SharePoint Online only)", name: "PersonalSiteNumberOfRetries", readOnly: true },
                "pictureimportenabled": { description: "A Boolean value that indicates whether the user's picture is imported from Exchange.", name: "PictureImportEnabled", readOnly: true },
                "publicurl": { description: "The public URL of the personal site of the current user. (SharePoint Online only)", name: "PublicUrl", readOnly: true },
                "urltocreatepersonalsite": { description: "The URL used to create the user's personal site.", name: "UrlToCreatePersonalSite", readOnly: true },
            },
            methods: {}
        },
        // Users
        "users": {
            methods: {}
        },
        // View
        "view": {
            properties: {
                "aggregations": { description: "Gets or sets a value that specifies fields and functions that define totals shown in a list view. If not null, the XML must conform to FieldRefDefinitionAggregation, as specified in [MS-WSSCAML].", name: "Aggregations", readOnly: false },
                "aggregationsstatus": { description: "Gets or sets a value that specifies whether totals are shown in the list view.", name: "AggregationsStatus", readOnly: false },
                "baseviewid": { description: "Gets a value that specifies the base view identifier of the list view.", name: "BaseViewId", readOnly: true },
                "contenttypeid": { description: "Gets or sets the identifier of the content type with which the view is associated so that the view is available only on folders of this content type.", name: "ContentTypeId", readOnly: false },
                "defaultview": { description: "Gets or sets a value that specifies whether the list view is the default list view.", name: "DefaultView", readOnly: false },
                "defaultviewforcontenttype": { description: "Gets or sets a value that specifies whether the list view is the default list view for the content type specified by contentTypeId.", name: "DefaultViewForContentType", readOnly: false },
                "editormodified": { description: "Gets or sets a value that specifies whether the list view was modified in an editor.", name: "EditorModified", readOnly: false },
                "formats": { description: "Gets or sets a value that specifies the column and row formatting for the list view. If not null, the XML must conform to ViewFormatDefinitions, as specified in [MS-WSSCAML].", name: "Formats", readOnly: false },
                "hidden": { description: "Gets or sets a value that specifies whether the list view is hidden.", name: "Hidden", readOnly: false },
                "htmlschemaxml": { description: "Gets a value that specifies the XML document that represents the list view.", name: "HtmlSchemaXml", readOnly: true },
                "id": { description: "Gets a value that specifies the view identifier of the list view.", name: "Id", readOnly: true },
                "imageurl": { description: "Gets a value that specifies the URI (Uniform Resource Identifier) of the image for the list view.", name: "ImageUrl", readOnly: true },
                "includerootfolder": { description: "Gets or sets a value that specifies whether the current folder is displayed in the list view.", name: "IncludeRootFolder", readOnly: false },
                "jslink": { description: "Gets or sets the name of the JavaScript file used for the view.", name: "JsLink", readOnly: false },
                "listviewxml": { description: "Gets or sets a string that represents the view XML.", name: "ListViewXml", readOnly: false },
                "method": { description: "Gets or sets a value that specifies the view method for the list view. If not null, the XML must conform to Method, as specified in [MS-WSSCAP].", name: "Method", readOnly: false },
                "mobiledefaultview": { description: "Gets or sets a value that specifies whether the list view is the default mobile list view.", name: "MobileDefaultView", readOnly: false },
                "mobileview": { description: "Gets or sets a value that specifies whether the list view is a mobile list view.", name: "MobileView", readOnly: false },
                "moderationtype": { description: "Gets a value that specifies the content approval type for the list view.", name: "ModerationType", readOnly: true },
                "orderedview": { description: "Gets a value that specifies whether list items can be reordered in the list view.", name: "OrderedView", readOnly: true },
                "paged": { description: "Gets or sets a value that specifies whether the list view is a paged view.", name: "Paged", readOnly: false },
                "personalview": { description: "Gets a value that specifies whether the list view is a personal view.", name: "PersonalView", readOnly: true },
                "readonlyview": { description: "Gets a value that specifies whether the list view is read-only.", name: "ReadOnlyView", readOnly: true },
                "requiresclientintegration": { description: "Gets a value that specifies whether the list view requires client integration rights.", name: "RequiresClientIntegration", readOnly: true },
                "rowlimit": { description: "Gets or sets a value that specifies the maximum number of list items to display in a visual page of the list view.", name: "RowLimit", readOnly: false },
                "scope": { description: "Gets or sets a value that specifies the scope for the list view. Represents a ViewScope value. DefaultValue = 0, Recursive = 1, RecursiveAll = 2, FilesOnly = 3.", name: "Scope", readOnly: false },
                "serverrelativeurl": { description: "Gets a value that specifies the server-relative URL of the list view page.", name: "ServerRelativeUrl", readOnly: true },
                "styleid": { description: "Gets a value that specifies the identifier of the view style for the list view.", name: "StyleId", readOnly: true },
                "threaded": { description: "Gets a value that specifies whether the list view is a threaded view.", name: "Threaded", readOnly: true },
                "title": { description: "Gets or sets a value that specifies the display name of the list view.", name: "Title", readOnly: false },
                "toolbar": { description: "Gets or sets a value that specifies the toolbar for the list view.", name: "Toolbar", readOnly: false },
                "toolbartemplatename": { description: "Gets a value that specifies the name of the template for the toolbar that is used in the list view.", name: "ToolbarTemplateName", readOnly: true },
                "viewdata": { description: "Gets or sets a value that specifies the view data for the list view. If not null, the XML must conform to FieldRefDefinitionViewData, as specified in [MS-WSSCAML].", name: "ViewData", readOnly: false },
                "viewfields": { description: "Gets a value that specifies the collection of fields in the list view.", name: "ViewFields", methodName: "get_ViewFields" },
                "viewjoins": { description: "Gets or sets a value that specifies the joins that are used in the list view. If not null, the XML must conform to ListJoinsDefinition, as specified in [MS-WSSCAML].", name: "ViewJoins", readOnly: false },
                "viewprojectedfields": { description: "Gets or sets a value that specifies the projected fields that will be used by the list view. If not null, the XML must conform to ProjectedFieldsDefinitionType, as specified in [MS-WSSCAML].", name: "ViewProjectedFields", readOnly: false },
                "viewquery": { description: "Gets or sets a value that specifies the query that is used by the list view. If not null, the XML must conform to CamlQueryRoot, as specified in [MS-WSSCAML].", name: "ViewQuery", readOnly: false },
                "viewtype": { description: "Gets a value that specifies the type of the list view. Can be HTML, GRID, CALENDAR, RECURRENCE, CHART, or GANTT.", name: "ViewType", readOnly: true },
            },
            methods: {}
        },
        // Views
        "views": {
            methods: {}
        },
        // View Fields
        "viewfields": {
            methods: {}
        },
        // Web
        "web": {
            properties: {
                "allowcreatedeclarativeworkflowforcurrentuser": { description: "Specifies whether the current user can create declarative workflows. If not disabled on the Web application, the value is the same as the AllowCreateDeclarativeWorkflow property of the site collection. Default value: true.", name: "AllowCreateDeclarativeWorkflowForCurrentUser", readOnly: true },
                "allowdesignerforcurrentuser": { description: "Gets a value that specifies whether the current user is allowed to use a designer application to customize this site.", name: "AllowDesignerForCurrentUser", readOnly: true },
                "allowmasterpageeditingforcurrentuser": { description: "Gets a value that specifies whether the current user is allowed to edit the master page.", name: "AllowMasterPageEditingForCurrentUser", readOnly: true },
                "allowrevertfromtemplateforcurrentuser": { description: "Gets a value that specifies whether the current user is allowed to revert the site to a default site template.", name: "AllowRevertFromTemplateForCurrentUser", readOnly: true },
                "allowrssfeeds": { description: "Gets a value that specifies whether the site allows RSS feeds.", name: "AllowRssFeeds", readOnly: true },
                "allowsavedeclarativeworkflowastemplateforcurrentuser": { description: "Specifies whether the current user can save declarative workflows as a template. If not disabled on the Web application, the value is the same as the AllowSaveDeclarativeWorkflowAsTemplate property of the site collection. Default value: true.", name: "AllowSaveDeclarativeWorkflowAsTemplateForCurrentUser", readOnly: true },
                "allowsavepublishdeclarativeworkflowforcurrentuser": { description: "Specifies whether the current user can save or publish declarative workflows. If not disabled on the Web application, the value is the same as the AllowSavePublishDeclarativeWorkflowAsTemplate property of the site collection. When enabled, can only be set by a site collection administrator. Default value: true.", name: "AllowSavePublishDeclarativeWorkflowForCurrentUser", readOnly: false },
                "allproperties": { description: "Gets a collection of metadata for the Web site.", name: "AllProperties", readOnly: true },
                "appinstanceid": { description: "The instance Id of the App Instance that this web represents.", name: "AppInstanceId", readOnly: true },
                "associatedmembergroup": { description: "Gets or sets the group of users who have been given contribute permissions to the Web site.", name: "AssociatedMemberGroup", readOnly: false },
                "associatedownergroup": { description: "Gets or sets the associated owner group of the Web site.", name: "AssociatedOwnerGroup", readOnly: false },
                "associatedvisitorgroup": { description: "Gets or sets the associated visitor group of the Web site.", name: "AssociatedVisitorGroup", readOnly: false },
                "availablecontenttypes": { description: "Gets the collection of all content types that apply to the current scope, including those of the current Web site, as well as any parent Web sites.", name: "AvailableContentTypes", methodName: "get_AvailableContentTypes" },
                "availablefields": { description: "Gets a value that specifies the collection of all fields available for the current scope, including those of the current site, as well as any parent sites.", name: "AvailableFields", methodName: "get_AvailableFields" },
                "configuration": { description: "Gets either the identifier (ID) of the site definition configuration that was used to create the site, or the ID of the site definition configuration from which the site template used to create the site was derived.", name: "Configuration", readOnly: true },
                "contenttypes": { description: "Gets the collection of content types for the Web site.", name: "ContentTypes", readOnly: true },
                "created": { description: "Gets a value that specifies when the site was created.", name: "Created", readOnly: true },
                "currentuser": { description: "Gets the current user of the site.", name: "CurrentUser", readOnly: true },
                "custommasterurl": { description: "Gets or sets the URL for a custom master page file to apply to the website.", name: "CustomMasterUrl", readOnly: false },
                "description": { description: "Gets or sets the description for the site.", name: "Description", readOnly: false },
                "designerdownloadurlforcurrentuser": { description: "Gets the URL where the current user can download SharePoint Designer.", name: "DesignerDownloadUrlForCurrentUser", readOnly: true },
                "documentlibrarycalloutofficewebapppreviewersdisabled": { description: "Determines if the Document Library Callout's WAC previewers are enabled or not.", name: "DocumentLibraryCalloutOfficeWebAppPreviewersDisabled", readOnly: true },
                "effectivebasepermissions": { description: "Represents the intersection of permissions of the app principal and the user principal. In the app-only case, this property returns only the permissions of the app principal.", name: "EffectiveBasePermissions", readOnly: true },
                "enableminimaldownload": { description: "Gets or sets a Boolean value that specifies whether the Web site should use Minimal Download Strategy.", name: "EnableMinimalDownload", readOnly: false },
                "eventreceivers": { description: "Gets the collection of event receiver definitions that are currently available on the website.", name: "EventReceivers", methodName: "get_EventReceivers" },
                "features": { description: "Gets a value that specifies the collection of features that are currently activated in the site.", name: "Features", methodName: "get_Features" },
                "fields": { description: "Gets the collection of field objects that represents all the fields in the Web site.", name: "Fields", methodName: "get_Fields" },
                "folders": { description: "Gets the collection of all first-level folders in the Web site.", name: "Folders", methodName: "get_Folders" },
                "id": { description: "Gets a value that specifies the site identifier for the site.", name: "Id", readOnly: true },
                "language": { description: "Gets a value that specifies the LCID for the language that is used on the site.", name: "Language", readOnly: true },
                "lastitemmodifieddate": { description: "Gets a value that specifies when an item was last modified in the site.", name: "LastItemModifiedDate", readOnly: true },
                "lists": { description: "Gets the collection of all lists that are contained in the Web site available to the current user based on the permissions of the current user.", name: "Lists", methodName: "get_Lists" },
                "listtemplates": { description: "Gets a value that specifies the collection of list definitions and list templates available for creating lists on the site.", name: "ListTemplates", methodName: "get_ListTemplates" },
                "masterurl": { description: "Gets or sets the URL of the master page that is used for the website.", name: "MasterUrl", readOnly: false },
                "navigation": { description: "Gets a value that specifies the navigation structure on the site, including the Quick Launch area and the top navigation bar.", name: "Navigation", readOnly: true },
                "parentweb": { description: "Gets the parent website of the specified website.", name: "ParentWeb", readOnly: true },
                "pushnotificationsubscribers": { description: "Gets the collection of push notification subscribers over the site.", name: "PushNotificationSubscribers", methodName: "get_PushNotificationSubscribers" },
                "quicklaunchenabled": { description: "Gets or sets a value that specifies whether the Quick Launch area is enabled on the site.", name: "QuickLaunchEnabled", readOnly: false },
                "recyclebin": { description: "Specifies the collection of recycle bin items of the recycle bin of the site.", name: "RecycleBin", readOnly: true },
                "recyclebinenabled": { description: "Gets or sets a value that determines whether the recycle bin is enabled for the website.", name: "RecycleBinEnabled", readOnly: true },
                "regionalsettings": { description: "Gets the regional settings that are currently implemented on the website.", name: "RegionalSettings", readOnly: true },
                "roledefinitions": { description: "Gets the collection of role definitions for the Web site.", name: "RoleDefinitions", methodName: "get_RoleDefinitions" },
                "rootfolder": { description: "Gets the root folder for the Web site.", name: "RootFolder", readOnly: true },
                "savesiteastemplateenabled": { description: "Gets or sets a Boolean value that specifies whether the Web site can be saved as a site template.", name: "SaveSiteAsTemplateEnabled", readOnly: false },
                "serverrelativeurl": { description: "Gets or sets the server-relative URL for the Web site.", name: "ServerRelativeUrl", readOnly: false },
                "showurlstructureforcurrentuser": { description: "Gets a value that specifies whether the current user is able to view the file system structure of this site.", name: "ShowUrlStructureForCurrentUser", readOnly: true },
                "sitegroups": { description: "Gets the collection of groups for the site collection.", name: "SiteGroups", methodName: "get_SiteGroups" },
                "siteuserinfolist": { description: "Gets the UserInfo list of the site collection that contains the Web site.", name: "SiteUserInfoList", readOnly: true },
                "siteusers": { description: "Gets the collection of all users that belong to the site collection.", name: "SiteUsers", methodName: "get_SiteUsers" },
                "supporteduilanguageids": { description: "Specifies the language code identifiers (LCIDs) of the languages that are enabled for the site.", name: "SupportedUILanguageIds", readOnly: true },
                "syndicationenabled": { description: "Gets or sets a value that specifies whether the RSS feeds are enabled on the site.", name: "SyndicationEnabled", readOnly: false },
                "themeinfo": { description: "The theming information for this site. This includes information like colors, fonts, border radii sizes etc.", name: "ThemeInfo", readOnly: true },
                "title": { description: "Gets or sets the title for the Web site.", name: "Title", readOnly: false },
                "treeviewenabled": { description: "Gets or sets value that specifies whether the tree view is enabled on the site.", name: "TreeViewEnabled", readOnly: false },
                "uiversion": { description: "Gets or sets the user interface (UI) version of the Web site.", name: "UIVersion", readOnly: false },
                "uiversionconfigurationenabled": { description: "Gets or sets a value that specifies whether the settings UI for visual upgrade is shown or hidden.", name: "UIVersionConfigurationEnabled", readOnly: false },
                "url": { description: "Gets the absolute URL for the website.", name: "Url", readOnly: true },
                "usercustomactions": { description: "Gets a value that specifies the collection of user custom actions for the site.", name: "UserCustomActions", methodName: "get_UserCustomActions" },
                "webinfos": { description: "Represents key properties of the subsites of a site.", name: "WebInfos", readOnly: true },
                "webs": { description: "Gets a Web site collection object that represents all Web sites immediately beneath the Web site, excluding children of those Web sites.", name: "Webs", methodName: "get_Webs" },
                "webtemplate": { description: "Gets the name of the site definition or site template that was used to create the site.", name: "WebTemplate", readOnly: true },
                "workflowassociations": { description: "Gets a value that specifies the collection of all workflow associations for the site.", name: "WorkflowAssociations", methodName: "get_WorkflowAssociations" },
                "workflowtemplates": { description: "Gets a value that specifies the collection of workflow templates associated with the site.", name: "WorkflowTemplates", readOnly: true },
            },
            methods: {}
        },
        // Webs
        "webs": {
            methods: {}
        }
    };

    // **********************************************************************************
    // Methods
    // **********************************************************************************

    // Method to generate the header.
    // header - The header.
    var generateHeaderInfo = function (header) {
        var messages = [];

        // Generate the header
        messages.push(_line);
        messages.push(header)
        messages.push(_line);

        // Return the messages;
        return messages;
    };

    // Method to generate method parameters
    // parameters - The method parameters.
    var generateMethodParameters = function (parameters) {
        // See if the parameters exist
        if (parameters && parameters.length > 0) {
            // Set the starting tag
            var code = "{";

            // Parse the parameters
            for (var i = 0; i < parameters.length; i++) {
                // Append the parameter
                code += (i == 0 ? "" : ",") + " " + parameters[i].name + ": " + parameters[i].sampleValue;
            }

            // Add the closing tag
            code += " }";

            // Return the code
            return code;
        }

        // Return nothing
        return null;
    };

    // Method to generate the js for the method.
    // methodInfo - The method information.
    var getMethodExample = function (methodInfo) {
        var messages = [];

        // Header
        messages.push("Code Example:");

        // Generate the code example
        var code = methodInfo.name + "(";

        // See if parameters are specified
        if (methodInfo.parameters) {
            // Generate the parameters
            var bodyParams = generateMethodParameters(methodInfo.parameters.body);
            var qsParams = generateMethodParameters(methodInfo.parameters.qs);

            // Generate the method
            code += (qsParams ? qsParams + ", " : "") + (bodyParams ? bodyParams : "") + (methodInfo.parameters.sampleValue ? methodInfo.parameters.sampleValue : "");
        }

        // Close the function
        code += ");";

        // Add the code example
        messages.push(code);

        // Return the messages
        return messages;
    };

    // Method to get the method information
    // objInfo - The object information.
    var getMethodInfo = function (objInfo) {
        // Header
        var messages = generateHeaderInfo("Method Information");

        // Parse the properties
        for (var key in objInfo.methods) {
            var methodInfo = objInfo.methods[key];

            // Add a messages
            messages.push(_lineDashes);
            messages.push(methodInfo.name + " [" + (methodInfo.type == _methodType.Get ? "Get" : "Post") + " Request]");
            messages.push(methodInfo.description);
            messages = messages.concat(getMethodExample(methodInfo));
            messages.push(_lineDashes);
            messages.push("");
        }

        // Return the messages
        return messages;
    };

    // Method to get the property information
    // objInfo - The object information.
    var getPropertyInfo = function (objInfo) {
        // Header
        var messages = generateHeaderInfo("Property Information");

        // Parse the properties
        for (var key in objInfo.properties) {
            var propInfo = objInfo.properties[key];

            // Add a messages
            messages.push(_lineDashes);
            messages.push(propInfo.name + " [" + (propInfo.readOnly ? "Read Only" : "Read/Write") + "]");
            if (propInfo.methodName) { messages.push("Method: " + propInfo.methodName); }
            messages.push(propInfo.description);
            messages.push(_lineDashes);
            messages.push("");
        }

        // Return the messages
        return messages;
    };

    // Method to log the help menu.
    var logHelpMenu = function () {
        var messages = [];

        // Log the default help to the console
        messages.push(_line);
        messages.push("Welcome to the help interface for the Bravo Core Library");
        messages.push("Author: Gunjan Datta");
        messages.push(_line);
        messages.push("You can interact with help by sending me the following information:");
        messages.push("");
        messages.push("'Object Type' - [Required] This is the 'type' property of the object.");
        messages.push("'Method' - [Optional] The method name of the object.");
        messages.push("");
        messages.push("Examples:");
        messages.push("BRAVO.Help('Web');");
        messages.push("BRAVO.Help('List', 'addItem');");
        messages.push("");
        messages.push("Use BRAVO.Help('All'); to output all available object types.");
        messages.push("Use BRAVO.Help('[Object Type]', 'Methods'); to output all available methods for a specified object type.");
        messages.push("Use BRAVO.Help('[Object Type]', 'Properties'); to output all available properties for a specified object type.");
        messages.push(_line);

        // Write to the log
        writeToLog(messages);
    };

    // Method to log the method information.
    // methodInfo - The method information.
    var logMethodInfo = function (methodInfo) {
        // Header
        var messages = generateHeaderInfo("Method Information");

        // Get method information
        // Properties: name, description, properties.qs, properties.body, type
        messages.push("Name: " + methodInfo.name);
        messages.push("Description: " + methodInfo.description);
        messages.push(_line);
        messages = messages.concat(getMethodExample(methodInfo));
        messages.push("");
        messages = messages.concat(getMethodExample(methodInfo));

        // Write to the log
        writeToLog(messages);
    };

    // Method to log the object information.
    // objInfo - The object information.
    var logObjectInfo = function (objInfo) {
        var messages = [];

        // Get object information
        messages = messages.concat(getPropertyInfo(objInfo));
        messages.push("");
        messages = messages.concat(getMethodInfo(objInfo));

        // Write to the log
        writeToLog(messages);
    };

    // Method to log the object types.
    var logObjectTypes = function () {
        var messages = [];

        // Header
        messages.push(_line);
        messages.push("Object Types");
        messages.push(_line);

        // Parse the object information
        for (var key in _objInfo) {
            // Output the key
            messages.push(key);
        }

        // Write to the log
        writeToLog(messages);
    };

    // Method to log the messages.
    // messages - The array of messages to log to the console.
    var writeToLog = function (messages) {
        var message = "";

        // Clear the console
        console.clear();

        // Parse the messages
        for (var i = 0; i < messages.length; i++) {
            message += messages[i] + "\r\n";
        }

        // Log the message
        console.log(message);
    }

    // **********************************************************************************
    // Main
    // **********************************************************************************

    // Ensure the input parameters are lower case
    objectType = objectType ? objectType.toLowerCase() : "";
    methodName = methodName ? methodName.toLowerCase() : "";

    // See if the object type was specified
    if (objectType) {
        // See if 'All' object types are being requested
        if (objectType == "all") {
            // Log the object types
            logObjectTypes();
            return;
        }

        // Get the object information for the object
        var objInfo = _objInfo[objectType];
        if (objInfo) {
            // See if the method name was specified
            if (methodName) {
                // See if only the methods are being requested
                if (methodName == "methods") {
                    // Log the method types
                    writeToLog(getMethodInfo(objInfo));
                    return;
                }

                // See if only the properties are being requested
                if (methodName == "properties") {
                    // Log the object types
                    writeToLog(getPropertyInfo(objInfo));
                    return;
                }

                // Get the method
                var methodInfo = _objInfo.methods[methodName];
                if (methodInfo) {
                    // Log the method information
                    logMethodInfo(methodInfo);
                    return;
                }
            }

            // Log the object information
            logObjectInfo(objInfo);
            return;
        }
    }

    // Log the help menu
    logHelpMenu();
};

// The JS Link class
BRAVO.JSLink = function () {
    // **********************************************************************************
    // Templates
    // **********************************************************************************

    // Button Classes
    var _btnClasses = {
        Button: "bravo-form-button",
        Row: "bravo-row bravo-button-set"
    };

    // Button Template
    var _btnCancelTemplate = '<input class="ms-ButtonHeightWidth" type="button" value="{{Text}}" target="_self" onclick="BRAVO.ModalDialog.close();" />';
    var _btnSaveTemplate = '<input class="ms-ButtonHeightWidth" type="button" value="Save" target="_self" onclick="SPClientForms.ClientFormManager.SubmitClientForm(\'{{FormID}}\');" />';
    var _btnFormTemplate = '<div class="{{CSSForm}}"><div class="{{CSSRow}}"><div class="{{CSSButton}}">{{Save}}</div><div class="{{CSSButton}}">{{Cancel}}</div></div></div>';
    var _btnTableTemplate = '<table width="100%" class="ms-formtoolbar" role="presentation" border="0" cellspacing="0" cellpadding="2"><tbody><tr>' +
        '<td width="99%" class="ms-toolbar" nowrap="nowrap"><img width="1" height="18" alt="" src="/_layouts/15/images/blank.gif?rev=43" data-accessibility-nocheck="true"></td>' +
        '<td class="ms-toolbar" nowrap="nowrap">{{Save}}</td>' +
        '<td class="ms-separator">&nbsp;</td>' +
        '<td class="ms-toolbar" nowrap="nowrap">{{Cancel}}</td>' +
        '</tr></tbody></table>';

    // Form Classes
    var _formClasses = {
        Field: "bravo-field",
        FieldDescription: "bravo-field-desc",
        FieldLabel: "bravo-field-label",
        FieldRequired: "bravo-field-required",
        Form: "bravo-form",
        Row: "bravo-row"
    };

    // Form Template
    var _formTemplate = '<div class="{{CSSForm}}">{{Rows}}</div>';
    var _formRowTemplate = '<div class="{{CSSRow}}" data-field-name="{{FieldName}}">' +
        '<div data-field-name="{{FieldName}}" class="{{CSSFieldLabel}}{{CSSFieldRequired}}">{{Label}}</div>' +
        '<div data-field-name="{{FieldName}}" class="{{CSSField}}">{{Field}}</div>' +
        '<div data-field-name="{{FieldName}}" class="{{CSSFieldDescription}}">{{Description}}</div>' +
        '</div>';

    // Table Template
    var _tblTemplate = '<table width="100%" class="ms-formtable" style="margin-top: 8px;" border="0" cellspacing="0" cellpadding="0"><tbody>{{TBODY}}</tbody></table>';
    var _tblButtonTemplate = '<tr><td colspan="99"><table width="100%" class="ms-formtoolbar" role="presentation" border="0" cellspacing="0" cellpadding="2"><tr><td>{{Save}}</td><td class="ms-separator">&nbsp;</td><td>{{Cancel}}</td></tr></table></td></tr>';
    var _tblRowTemplate = '<tr data-field-name="{{FieldName}}"><td width="113" class="ms-formlabel" data-field-name="{{FieldName}}" nowrap="true" valign="top">{{Label}}</td><td width="350" class="ms-formbody" data-field-name="{{FieldName}}" valign="top">{{Field}}</td></tr>';

    // **********************************************************************************
    // Global Variables
    // **********************************************************************************

    // Current Field
    var _currentField = null;
    var getCurrentField = function (ctx) {
        // Return if it already exists
        if (_currentField != null) { return _currentField; }

        // Ensure the field schema exists, and return it
        return ctx && ctx.CurrentFieldSchema ? ctx.CurrentFieldSchema : null;
    };

    // Current Item
    var _currentItem = null;
    var getCurrentItem = function (ctx) {
        // Return if it already exists
        if (_currentItem != null) { return _currentItem; }

        // Get the item id
        var itemId = ctx.FormContext ? ctx.FormContext.itemAttributes.Id : BRAVO.Core.getQueryStringValue("ID");
        if (itemId) {
            // Get the current list
            var list = getCurrentList(ctx);
            if (list && list.exists) {
                // Get the current list item
                _currentItem = list.getItemById(itemId);
            }
        }

        // Return the current list item
        return _currentItem;
    };

    // Current List
    var _currentList = null;
    var getCurrentList = function (ctx) {
        // Return if it already exists
        if (_currentList != null) { return _currentList; }

        // Get the current web
        var web = getCurrentWeb();
        if (web.exists) {
            // Get the current list id or title
            var listId = ctx && ctx.FormContext && ctx.FormContext.listAttributes ? ctx.FormContext.listAttributes.Id : null;
            var listTitle = ctx ? ctx.ListTitle || ctx.ListName : null;

            // Get the list
            _currentList = listId ? web.getListById(listId) : (listTitle ? web.getListByTitle(listTitle) : null);
        }

        // Return the current list
        return _currentList;
    };

    // Current List Fields
    var _currentListFields = null;
    var getCurrentListFields = function (ctx) {
        // Return if it already exists
        if (_currentListFields) { return _currentListFields; }

        // Get the current list
        var list = getCurrentList(ctx);
        if (list && list.exists) {
            // Get the fields
            _currentListFields = list.get_Fields();
        }

        // Return the current list fields
        return _currentListFields;
    };

    // Current User
    var _currentUser = null;
    function getCurrentUser() {
        // Return if it already exists
        if (_currentUser != null) { return _currentUser; }

        // Get the current web
        var web = getCurrentWeb();
        if (web.exists) {
            // Get the current user
            _currentUser = web.get_CurrentUser();
        }

        // Return the current user
        return _currentUser;
    };

    // Current Web
    var _currentWeb = null;
    var getCurrentWeb = function () {
        // Return if it already exists
        if (_currentWeb != null) { return _currentWeb; }

        // Set the current web
        _currentWeb = new BRAVO.Core.Web();

        // Return the current web
        return _currentWeb;
    };

    // Field to Method Mapper
    var _fieldToMethodMapper = {
        'Attachments': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPFieldAttachments_Default,
            'EditForm': window.SPFieldAttachments_Default,
            'NewForm': window.SPFieldAttachments_Default
        },
        'Boolean': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPField_FormDisplay_DefaultNoEncode,
            'EditForm': window.SPFieldBoolean_Edit,
            'NewForm': window.SPFieldBoolean_Edit
        },
        'Currency': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPField_FormDisplay_Default,
            'EditForm': window.SPFieldNumber_Edit,
            'NewForm': window.SPFieldNumber_Edit
        },
        'Calculated': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPField_FormDisplay_Default,
            'EditForm': window.SPField_FormDisplay_Empty,
            'NewForm': window.SPField_FormDisplay_Empty
        },
        'Choice': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPField_FormDisplay_Default,
            'EditForm': window.SPFieldChoice_Edit,
            'NewForm': window.SPFieldChoice_Edit
        },
        'Computed': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPField_FormDisplay_Default,
            'EditForm': window.SPField_FormDisplay_Default,
            'NewForm': window.SPField_FormDisplay_Default
        },
        'DateTime': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPFieldDateTime_Display,
            'EditForm': window.SPFieldDateTime_Edit,
            'NewForm': window.SPFieldDateTime_Edit
        },
        'File': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPFieldFile_Display,
            'EditForm': window.SPFieldFile_Edit,
            'NewForm': window.SPFieldFile_Edit
        },
        'Integer': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPField_FormDisplay_Default,
            'EditForm': window.SPFieldNumber_Edit,
            'NewForm': window.SPFieldNumber_Edit
        },
        'Lookup': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPFieldLookup_Display,
            'EditForm': window.SPFieldLookup_Edit,
            'NewForm': window.SPFieldLookup_Edit
        },
        'LookupMulti': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPFieldLookup_Display,
            'EditForm': window.SPFieldLookup_Edit,
            'NewForm': window.SPFieldLookup_Edit
        },
        'MultiChoice': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPField_FormDisplay_Default,
            'EditForm': window.SPFieldMultiChoice_Edit,
            'NewForm': window.SPFieldMultiChoice_Edit
        },
        'Note': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPFieldNote_Display,
            'EditForm': window.SPFieldNote_Edit,
            'NewForm': window.SPFieldNote_Edit
        },
        'Number': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPField_FormDisplay_Default,
            'EditForm': window.SPFieldNumber_Edit,
            'NewForm': window.SPFieldNumber_Edit
        },
        'Text': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPField_FormDisplay_Default,
            'EditForm': window.SPFieldText_Edit,
            'NewForm': window.SPFieldText_Edit
        },
        'URL': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPFieldUrl_Display,
            'EditForm': window.SPFieldUrl_Edit,
            'NewForm': window.SPFieldUrl_Edit
        },
        'User': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPFieldUser_Display,
            'EditForm': window.SPClientPeoplePickerCSRTemplate,
            'NewForm': window.SPClientPeoplePickerCSRTemplate
        },
        'UserMulti': {
            'View': window.RenderFieldValueDefault,
            'DisplayForm': window.SPFieldUserMulti_Display,
            'EditForm': window.SPClientPeoplePickerCSRTemplate,
            'NewForm': window.SPClientPeoplePickerCSRTemplate
        }
    };

    // **********************************************************************************
    // Methods
    // **********************************************************************************

    // Method to add the field name as a custom data attribute
    var addFieldNameAttribute = function (ctx) {
        var elementTypes = ["input", "select"];

        // Get the default html for this element
        var fld = document.createElement("div");
        fld.innerHTML = getFieldDefaultHtml(ctx);

        // Parse the elements
        for (var i = 0; i < elementTypes.length; i++) {
            var element = fld.querySelector(elementTypes[i]);
            if (element) {
                // Set the custom attribute
                element.setAttribute("data-field-name", ctx.CurrentFieldSchema.Name);
                break;
            }
        }

        // Return the default html
        return fld.innerHTML;
    };

    // Method to add a script reference to the page.
    var addScript = function (url) {
        // Create the element
        var e = document.createElement("script");

        // Set the properties
        e.setAttribute("src", url);
        e.setAttribute("type", "text/javascript");

        // Add the element to the head
        document.head.appendChild(e);
    };

    // Method to add a style sheet reference to the page.
    var addStyle = function (url) {
        // Create the element
        var e = document.createElement("link");

        // Set the properties
        e.setAttribute("href", url);
        e.setAttribute("rel", "stylesheet");
        e.setAttribute("type", "text/css");

        // Add the element to the head
        document.head.appendChild(e);
    };

    // Method to add custom validation to the current field
    var addValidation = function (ctx, validationFunc) {
        // Validate the input
        if (ctx == null || validationFunc == null) { return; }

        // Get the form context
        var formContext = SPClientTemplates.Utility.GetFormContextForCurrentField(ctx);

        // Create a validator set
        var validators = new SPClientForms.ClientValidation.ValidatorSet();

        // See if we are adding multiple validation functions
        if (typeof (validationFunc) == "object" && validationFunc.length) {
            // Parse the validation functions
            for (var i = 0; i < validationFunc.length; i++) {
                // Register the custom validation function
                validators.RegisterValidator(new fieldValidator(validationFunc[i]));
            }
        } else {
            // Create the validation function
            var f = function () { }
            f.prototype.Validate = validationFunc;

            // Register the custom validation function
            validators.RegisterValidator(new f());
        }

        // Register the custom validation
        formContext.registerClientValidator(ctx.CurrentFieldSchema.Name, validators);
    };

    // Method to default a user field's value to the current user
    var defaultToCurrentUser = function (ctx) {
        // Get the current user
        var user = getCurrentUser();
        if (user.exists) {
            // Set the default user
            return getFieldUserHtml(ctx, user);
        }

        // Return the default html
        return getFieldDefaultHtml(ctx);
    };

    // Method to disable edit
    var disableEdit = function (ctx, field) {
        // Ensure a value exists
        if (ctx.CurrentFieldValue) {
            var fieldValue = "";

            // Update the context, based on the field type
            switch (ctx.CurrentFieldSchema.Type) {
                case "MultiChoice":
                    var regExp = new RegExp(SPClientTemplates.Utility.UserLookupDelimitString, "g");

                    // Update the field value
                    fieldValue = ctx.CurrentFieldValue
                        // TO DO - Add comment to explain regular expression
                        .replace(regExp, "; ")
                        // TO DO - Add comment to explain regular expression
                        .replace(/^; /g, "")
                        // TO DO - Add comment to explain regular expression
                        .replace(/; $/g, "");
                    break;
                case "Note":
                    // Replace the return characters
                    fieldValue = "<div>" + ctx.CurrentFieldValue.replace(/\n/g, "<br />") + "</div>";
                    break;
                case "User":
                case "UserMulti":
                    // Parse the user values
                    for (var i = 0; i < ctx.CurrentFieldValue.length; i++) {
                        // Add the user value
                        fieldValue +=
                            // User Lookup ID
                            ctx.CurrentFieldValue[i].EntityData.SPUserID +
                            // Delimiter
                            SPClientTemplates.Utility.UserLookupDelimitString +
                            // User Lookup Value
                            ctx.CurrentFieldValue[i].DisplayText +
                            // Optional Delimiter
                            ((i == ctx.CurrentFieldValue.length - 1 ? "" : SPClientTemplates.Utility.UserLookupDelimitString));
                    }
                    break;
            };

            // Update the current field value
            ctx.CurrentFieldValue = fieldValue;
        }

        // Return the display value of the field
        return getFieldDefaultHtml(ctx, field, "DisplayForm");
    };

    // Method to disable quick edit
    var disableQuickEdit = function (ctx, field) {
        // Ensure we are in grid edit mode
        if (ctx.inGridMode) {
            // Disable editing for this field
            field.AllowGridEditing = false;
            return "";
        }

        // Return the default field value html
        return getFieldDefaultHtml(ctx, field);
    };

    // The field validator helper method
    var fieldValidator = function (validationFunc) { fieldValidator.prototype.Validate = validationFunc; };

    // Method to get the default form fields
    // ctx - The form context.
    var getDefaultFormFields = function (ctx) {
        var formFields = [];

        // Get the list and fields
        var list = getCurrentList(ctx);
        var fields = list && list.exists ? list.get_Fields() : null;
        if (fields) {
            // Get the default content type field links
            var fieldLinks = list.get_ContentTypes().results[0].get_FieldLinks().results;

            // Default the form fields
            for (var i = 0; i < fieldLinks.length; i++) {
                // Get the field
                var field = fields.getByTitle(fieldLinks[i].Name);
                if (field) {
                    // See if this field should be displayed
                    if (field.Hidden || field.Group == "_Hidden" || (!field.CanBeDeleted && field.InternalName != "Title")) { continue; }

                    // Add the field
                    formFields.push(field.InternalName);
                }
            }
        }

        // Return the form fields
        return formFields;
    };

    // Method to get the default field html
    var getFieldDefaultHtml = function (ctx, field, formType) {
        // Determine the field type
        var fieldType = field ? field.Type : (ctx.CurrentFieldSchema ? ctx.CurrentFieldSchema.Type : null);

        // Ensure the form type is set
        formType = formType ? formType : getFormType(ctx);

        // Ensure a field to method mapper exists
        if (_fieldToMethodMapper[fieldType] && _fieldToMethodMapper[fieldType][formType]) {
            // Return the default html for this field
            var defaultHtml = _fieldToMethodMapper[fieldType][formType](ctx);
            if (defaultHtml) { return defaultHtml; }
        }

        // Set the field renderer based on the field type
        var field = ctx.CurrentFieldSchema;
        var fieldRenderer = null;
        switch (field.Type) {
            case "AllDayEvent": fieldRenderer = new AllDayEventFieldRenderer(field.Name); break;
            case "Attachments": fieldRenderer = new AttachmentFieldRenderer(field.Name); break;
            case "BusinessData": fieldRenderer = new BusinessDataFieldRenderer(field.Name); break;
            case "Computed": fieldRenderer = new ComputedFieldRenderer(field.Name); break;
            case "CrossProjectLink": fieldRenderer = new ProjectLinkFieldRenderer(field.Name); break;
            case "Currency": fieldRenderer = new NumberFieldRenderer(field.Name); break;
            case "DateTime": fieldRenderer = new DateTimeFieldRenderer(field.Name); break;
            case "Lookup": fieldRenderer = new LookupFieldRenderer(field.Name); break;
            case "LookupMulti": fieldRenderer = new LookupFieldRenderer(field.Name); break;
            case "Note": fieldRenderer = new NoteFieldRenderer(field.Name); break;
            case "Number": fieldRenderer = new NumberFieldRenderer(field.Name); break;
            case "Recurrence": fieldRenderer = new RecurrenceFieldRenderer(field.Name); break;
            case "Text": fieldRenderer = new TextFieldRenderer(field.Name); break;
            case "URL": fieldRenderer = new UrlFieldRenderer(field.Name); break;
            case "User": fieldRenderer = new UserFieldRenderer(field.Name); break;
            case "UserMulti": fieldRenderer = new UserFieldRenderer(field.Name); break;
            case "WorkflowStatus": fieldRenderer = new RawFieldRenderer(field.Name); break;
        };

        // Get the current item
        var currentItem = ctx.CurrentItem || ctx.ListData.Items[0];

        // Return the item's field value html
        return fieldRenderer ? fieldRenderer.RenderField(ctx, field, currentItem, ctx.ListSchema) : currentItem[field.Name];
    };

    // Method to get a multi-user field html
    var getFieldMultiUserHtml = function (ctx, users) {
        // Clear the field value
        ctx.CurrentFieldValue = [];

        // Parse the users
        for (var i = 0; i < users.length; i++) {
            // Add the user field value
            ctx.CurrentFieldValue.push({
                Description: users[i].LoginName,
                DisplayText: users[i].Title,
                EntityGroupName: "",
                EntityType: "",
                HierarchyIdentifier: null,
                IsResolved: true,
                Key: users[i].LoginName,
                MultipleMatches: [],
                ProviderDisplayName: "",
                ProviderName: ""
            });
        }

        // Return the html
        return getFieldDefaultHtml(ctx);
    };

    // Method to get a user field value
    var getFieldUserHtml = function (ctx, user) {
        // Get the value
        return getFieldMultiUserHtml(ctx, [user]);
    };

    // Method to determine the form type
    // ctx - The form context.
    var getFormType = function (ctx) {
        var formType = null;

        // Determine the form type
        switch (ctx.ControlMode) {
            case SPClientTemplates.ClientControlMode.DisplayForm:
                formType = "DisplayForm";
                break;
            case SPClientTemplates.ClientControlMode.EditForm:
                formType = "EditForm";
                break;
            case SPClientTemplates.ClientControlMode.NewForm:
                formType = "NewForm";
                break;
            case SPClientTemplates.ClientControlMode.View:
                formType = "View";
                break;
        }

        // Return the form type
        return formType;
    };

    // Method to get the list view element
    var getListView = function (ctx) { var wp = getWebPart(ctx); return wp ? wp.querySelector(".ms-formtable") : null; };

    // Method to get the view's list items
    var getViewListItems = function (ctx) { return ctx.ListData ? ctx.ListData.Row : []; };

    // Method to get the view's selected items
    var getViewSelectedItems = function (ctx) { return SP.ListOperation.Selection.getSelectedItems(); };

    // Method to get the list view element
    var getWebPart = function (ctx) { return document.querySelector("#WebPart" + (ctx.FormUniqueId || ctx.wpq)); };

    // Method to hide the row containing the field
    // ctx - The form context.
    var hideField = function (ctx) {
        // See if we have already set the event to hide fields
        if (BRAVO.JSLink._hideFieldEventFl == null) {
            // Add an onload event to it
            window.addEventListener("load", function () {
                debugger;
                // Query for the elements we need to hide
                var fieldElements = document.querySelectorAll(".hide-row");
                for (var i = 0; i < fieldElements.length; i++) {
                    var fieldElement = fieldElements[i];

                    // Ensure the parent element exists
                    if (fieldElement.parentNode && fieldElement.parentNode.parentNode) {
                        // See if the parent element is this row element
                        if (fieldElement.parentNode.getAttribute("data-field-name") == fieldElement.parentNode.parentNode.getAttribute("data-field-name")) {
                            // Set the element to the parent
                            fieldElement = fieldElement.parentNode.parentNode;
                        }
                        else {
                            // Find the parent row
                            while (fieldElement && fieldElement.nodeName.toLowerCase() != "tr") { fieldElement = fieldElement.parentNode; }
                        }
                    }


                    // Hide the row
                    if (fieldElement) { fieldElement.style.display = "none"; }
                }
            });

            // Set the flag
            BRAVO.JSLink._hideFieldEventFl = true;
        }

        // Create an empty element
        return "<div class='hide-row'>" + getFieldDefaultHtml(ctx) + "</div>";
    };

    // Method to determine if the page/form is currently being edited
    var inEditMode = function () {
        // Ensure the page form name exists
        var wppForm = document.forms[MSOWebPartPageFormName];
        if (wppForm) {
            // Detect web part page
            if (wppForm.MSOLayout_InDesignMode && wppForm.MSOLayout_InDesignMode.value == "1") { return true; }

            // Detect wiki page
            if (wppForm._wikiPageMode && wppForm._wikiPageMode.value == "Edit") { return true; }
        }

        // Page is not detected in edit mode
        return false;
    };

    // Method to remove the row containing the field
    var removeField = function (ctx) {
        // Hide this field
        hideField(ctx);

        // Create an empty element
        return "<div class='hide-row' />";
    };

    // Method to render the field and return the html for it.
    // ctx - The form context.
    // fieldName - The internal field name.
    // showDescFl - Flag to display the description.
    var renderFieldHtml = function (ctx, fieldName, showDescFl) {
        // Default the flag to display the description
        showDescFl = showDescFl == null ? true : showDescFl;

        // Get the field
        var field = ctx.ListSchema.Field.filter(function (field) { return field.Name == fieldName; });
        if (field && field.length > 0) {
            // Set the field as the current one
            ctx.CurrentFieldSchema = field[0];
            ctx.CurrentFieldValue = ctx.ListData.Items[0][fieldName];

            // Update the description
            ctx.CurrentFieldSchema.Description = showDescFl ? ctx.CurrentFieldSchema.Description : "";

            // Note - There is a bug w/ the user field containing a trailing </div> tag that is causing the html to shift in the form.
            // The code below will ensure the html is normalized.

            // Create a dummy element to store the html for this field
            var fieldHtml = document.createElement("div");
            fieldHtml.innerHTML = ctx.Templates.Fields[field[0].Name](ctx);

            // Return the field's html
            return fieldHtml.innerHTML;
        }

        // Invalid field, return nothing
        return "";
    };

    // Method to render a form.
    // ctx - The form context.
    // formFields - The fields to render in the form.
    var renderForm = function (ctx, formFields, css) {
        var formRows = "";

        // Default the css
        css = css ? css : _formClasses;

        // Get the current list fields
        var fields = getCurrentListFields(ctx);

        // Default the form fields
        formFields = formFields ? formFields : (fields ? getDefaultFormFields(ctx) : null);

        // Ensure the form and list fields exist
        if (formFields && fields) {
            // Parse the fields to add
            for (var i = 0; i < formFields.length; i++) {
                // Get the field
                var field = fields.getByTitle(formFields[i]);
                if (field) {
                    // Append the row
                    formRows += _formRowTemplate
                        .replace(/{{CSSField}}/g, css.Field ? css.Field : _formClasses.Field)
                        .replace(/{{CSSFieldDescription}}/g, css.FieldDescription ? css.FieldDescription : _formClasses.FieldDescription)
                        .replace(/{{CSSFieldLabel}}/g, css.FieldLabel ? css.FieldLabel : _formClasses.FieldLabel)
                        .replace(/{{CSSFieldRequired}}/g, field.Required ? " " + (css.FieldRequired ? css.FieldRequired : _formClasses.FieldRequired) : "")
                        .replace(/{{CSSRow}}/g, css.Row ? css.Row : _formClasses.Row)
                        .replace(/{{Description}}/g, field.Description)
                        .replace(/{{Field}}/g, BRAVO.JSLink.renderFieldHtml(ctx, field.InternalName, false))
                        .replace(/{{FieldName}}/g, field.InternalName)
                        .replace(/{{Label}}/g, field.Title);
                }
            }
        }

        // Return the form
        return _formTemplate
            .replace(/{{CSSForm}}/g, css.Form ? css.Form : _formClasses.Form)
            .replace(/{{Rows}}/g, formRows);
    };

    // Method to render the form buttons
    // ctx - The form context.
    var renderFormButtons = function (ctx, css) {
        // Default the css
        css = css ? css : _btnClasses;

        // Determine the form type
        var formType = getFormType(ctx);

        // Determine what to render, based on the form type
        var renderSaveButtonFl = formType == "NewForm" || formType == "EditForm";
        var renderCloseButtonFl = !renderSaveButtonFl;

        // Return the table
        return _btnFormTemplate
            .replace(/{{CSSButton}}/g, css.Form ? css.Button : _btnClasses.Button)
            .replace(/{{CSSForm}}/g, css.Form ? css.Form : _btnClasses.Form)
            .replace(/{{CSSRow}}/g, css.Row ? css.Row : _btnClasses.Row)
            .replace(/{{Save}}/g, renderSaveButtonFl ? _btnSaveTemplate.replace(/{{FormID}}/g, ctx.FormUniqueId) : "")
            .replace(/{{Cancel}}/g, _btnCancelTemplate.replace(/{{Text}}/g, renderCloseButtonFl ? "Close" : "Cancel"));
    };

    // Method to render a table.
    // ctx - The form context.
    // formFields - The fields to render in the table.
    var renderTable = function (ctx, formFields) {
        var tbody = "";

        // Default the form fields, if needed
        formFields = formFields ? formFields : getDefaultFormFields(ctx);

        // Get the current list fields
        var fields = getCurrentListFields();

        // Parse the form fields
        for (var i = 0; i < formFields.length; i++) {
            // Get the field
            var field = fields.getByTitle(formFields[i]);
            if (field) {
                // Append the row for this field
                tbody += _tblRowTemplate
                    .replace(/{{Label}}/g, field.Title + (field.Required ? '<span class="ms-formvalidation"> *</span>' : ""))
                    .replace(/{{Field}}/g, BRAVO.JSLink.renderFieldHtml(ctx, field.InternalName))
                    .replace(/{{FieldName}}/g, field.InternalName);
            }
        }

        // Return the table
        return _tblTemplate.replace(/{{TBODY}}/g, tbody);
    };

    // Method to render the table buttons
    // ctx - The form context.
    var renderTableButtons = function (ctx) {
        // Determine the form type
        var formType = getFormType(ctx);

        // Determine what to render, based on the form type
        var renderSaveButtonFl = formType == "NewForm" || formType == "EditForm";
        var renderCloseButtonFl = !renderSaveButtonFl;

        // Return the table
        return _btnTableTemplate
            .replace(/{{Save}}/g, renderSaveButtonFl ? _btnSaveTemplate.replace(/{{FormID}}/g, ctx.FormUniqueId) : "")
            .replace(/{{Cancel}}/g, _btnCancelTemplate.replace(/{{Text}}/g, renderCloseButtonFl ? "Close" : "Cancel"));
    };

    // Method to send an email
    var sendEmail = function (subject, body, to, from) {
        // Get the current web
        var web = getCurrentWeb();
        if (web) {
            // Ensure the email is an array
            to = typeof (to) === "string" ? [to] : to;

            // Send the email
            web.sendEmail({
                To: { results: to },
                From: from ? from : getCurrentUser().Email,
                Subject: subject,
                Body: body
            });
        }
    };

    // **********************************************************************************
    // Public Interface
    // **********************************************************************************

    return {
        addFieldNameAttribute: addFieldNameAttribute,
        addScript: addScript,
        addStyle: addStyle,
        addValidation: addValidation,
        defaultToCurrentUser: defaultToCurrentUser,
        disableEdit: disableEdit,
        disableQuickEdit: disableQuickEdit,
        getCurrentField: getCurrentField,
        getCurrentItem: getCurrentItem,
        getCurrentList: getCurrentList,
        getCurrentUser: getCurrentUser,
        getCurrentWeb: getCurrentWeb,
        getFieldDefaultHtml: getFieldDefaultHtml,
        getFieldMultiUserHtml: getFieldMultiUserHtml,
        getFieldUserHtml: getFieldUserHtml,
        getListView: getListView,
        getViewListItems: getViewListItems,
        getViewSelectedItems: getViewSelectedItems,
        getWebPart: getWebPart,
        hideField: hideField,
        inEditMode: inEditMode,
        removeField: removeField,
        renderFieldHtml: renderFieldHtml,
        renderForm: renderForm,
        renderFormButtons: renderFormButtons,
        renderTable: renderTable,
        renderTableButtons: renderTableButtons,
        sendEmail: sendEmail
    };
}();

// The modal dialog class.
BRAVO.ModalDialog = function () {
    // **********************************************************************************
    // Private Interface
    // **********************************************************************************

    // **********************************************************************************
    // Methods
    // **********************************************************************************

    // Method to close the modal dialog
    var close = function (dialogResult, returnVal) {
        // Ensure this is a dialog
        if (isDialog()) {
            // See if the dialog result is defined and is a number
            if (dialogResult != null && typeof (dialogResult) === "number") {
                // Close the dialog
                SP.UI.ModalDialog.commonModalDialogClose(dialogResult, returnVal);
            } else {
                // Close the dialog with the 'Cancel' result
                SP.UI.ModalDialog.commonModalDialogClose(SP.UI.DialogResult.cancel, returnVal);
            }
        }
        else {
            // Get the redirect url from the query string
            var redirectUrl = BRAVO.Core.getQueryStringValue("Source");
            if (redirectUrl && redirectUrl.length > 0) {
                // Redirect the user
                window.location.href = redirectUrl;
            }
            else {
                // Go back a page
                window.history.back();
            }
        }
    };

    // Method to get arguments passed to the dialog
    var getArguments = function () {
        // Return the arguments passed to the dialog
        return window.frameElement ? window.frameElement.dialogArgs : null;
    };

    // Method to determine if a dialog is open
    var isDialog = function () {
        return SP.UI.ModalDialog.get_childDialog() != null;
    };

    // Method to open a modal dialog
    var open = function (title, source, maximized, args, onClose) {
        if (arguments.length == 0) { return; }

        // Create the options
        var options = {
            args: args ? args : null,
            dialogReturnValueCallback: onClose ? onClose : null,
            showMaximized: maximized ? true : false,
            title: title
        };

        // See if the source is a string
        if (typeof (source) === "string") {
            // Set the url
            options.url = source;
        } else {
            // Set the html
            options.html = source;
        }

        // Open the dialog
        return SP.UI.ModalDialog.showModalDialog(options);
    };

    // Method to open a modal dialog with the defined 'html' input parameter
    var openHtml = function (title, html, maximized, args, onClose) {
        // Create an element to hold the html
        var div = document.createElement("div");
        div.innerHTML = html;

        // Open the dialog
        open(title, div, maximized, args, onClose);
    };

    // Method to refresh the current page if the result is 'OK'
    // If the return value is detected as a url, it will redirect to it.
    var refreshOnSuccess = function (dialogResult, returnVal) {
        // See if the 'OK' button was clicked
        if (dialogResult == SP.UI.DialogResult.OK) {
            // See if the return value is a url
            if (returnVal && typeof (returnVal) === "string" &&
                (returnVal.toLowerCase().indexOf("http") == 0 || returnVal.indexOf("/") == 0)) {
                // Redirect to the url
                document.location.href = returnVal;
            }

            // Refresh the page
            document.location.reload();
        }
    };

    // Method to show the wait screen id
    var showWaitScreen = function (title, message, height, width) {
        if (arguments.length == 0) { return; }

        // See if the height & width are defined
        if (height && width) {
            // Show the wait screen
            SP.UI.ModalDialog.showWaitScreenWithNoClose(title, message, height, width);
        } else {
            // Show the wait screen
            SP.UI.ModalDialog.showWaitScreenWithNoClose(title, message);
        }
    };

    // Method to update the wait screen
    var updateWaitScreen = function (title, message) {
        var oldData = { title: null, message: null };

        // Get the dialog
        var dialog = document.querySelector(".ms-dlgContent");
        if (dialog) {
            // Get the title
            var dlgTitle = dialog.querySelector(".ms-core-pageTitle");
            if (dlgTitle) {
                // Set the title
                oldData.title = dlgTitle.innerHTML;
                dlgTitle.innerHTML = title != null ? title : oldData.title;
            }

            // Get the message
            var dlgMessage = dialog.querySelector(".ms-textXLarge");
            if (dlgMessage) {
                // Set the message
                oldData.message = dlgMessage.innerHTML;
                dlgMessage.innerHTML = message != null ? message : oldData.message;
            }

            // Resize the screen
            SP.UI.ModalDialog.get_childDialog().autoSize();
        }

        // Return the old title and message
        return oldData;
    };

    // **********************************************************************************
    // Public Interface
    // **********************************************************************************
    return {
        // Methods
        close: close,
        getArguments: getArguments,
        isDialog: isDialog,
        open: open,
        openHtml: openHtml,
        refreshOnSuccess: refreshOnSuccess,
        showWaitScreen: showWaitScreen,
        updateWaitScreen: updateWaitScreen
    };
}();

// The initialization for the library
BRAVO.Init = function () {
    // Ensure the dependencies are loaded
    BRAVO.Core.loadDependencies(function () {
        // Ensure the dialog class is loaded
        SP.SOD.executeFunc("sp.ui.dialog.js", "SP.UI.ModalDialog", function () {
            // Notify scripts that the modal dialog class is loaded
            SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("bravo.modaldialog.js");

            // Notify scripts that the bravo library is loaded
            SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("bravo.js");
        });

        // Notify scripts that the core class is loaded
        SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("bravo.core.js");

        // Notify scripts that the js link class is loaded
        SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("bravo.jslink.js");
    });
}

// Write the javascript to the page. This will ensure it's called when MDS is enabled
document.write("<script type='text/javascript'>(function() { BRAVO.Init(); })();</script>");

/* Bravo Core Library v1.7.6 | (c) Bravo Consulting Group, LLC (Bravo) | bravocg.com */

// Promise Polyfill to provide access to promises for browsers without it implemented
// Taken from https://github.com/taylorhakes/promise-polyfill
(function (root) {

    // Store setTimeout reference so promise-polyfill will be unaffected by
    // other code modifying setTimeout (like sinon.useFakeTimers())
    var setTimeoutFunc = setTimeout;

    function noop() { }

    // Polyfill for Function.prototype.bind
    function bind(fn, thisArg) {
        return function () {
            fn.apply(thisArg, arguments);
        };
    }

    function Promise(fn) {
        if (typeof this !== 'object') throw new TypeError('Promises must be constructed via new');
        if (typeof fn !== 'function') throw new TypeError('not a function');
        this._state = 0;
        this._handled = false;
        this._value = undefined;
        this._deferreds = [];

        doResolve(fn, this);
    }

    function handle(self, deferred) {
        while (self._state === 3) {
            self = self._value;
        }
        if (self._state === 0) {
            self._deferreds.push(deferred);
            return;
        }
        self._handled = true;
        Promise._immediateFn(function () {
            var cb = self._state === 1 ? deferred.onFulfilled : deferred.onRejected;
            if (cb === null) {
                (self._state === 1 ? resolve : reject)(deferred.promise, self._value);
                return;
            }
            var ret;
            try {
                ret = cb(self._value);
            } catch (e) {
                reject(deferred.promise, e);
                return;
            }
            resolve(deferred.promise, ret);
        });
    }

    function resolve(self, newValue) {
        try {
            // Promise Resolution Procedure: https://github.com/promises-aplus/promises-spec#the-promise-resolution-procedure
            if (newValue === self) throw new TypeError('A promise cannot be resolved with itself.');
            if (newValue && (typeof newValue === 'object' || typeof newValue === 'function')) {
                var then = newValue.then;
                if (newValue instanceof Promise) {
                    self._state = 3;
                    self._value = newValue;
                    finale(self);
                    return;
                } else if (typeof then === 'function') {
                    doResolve(bind(then, newValue), self);
                    return;
                }
            }
            self._state = 1;
            self._value = newValue;
            finale(self);
        } catch (e) {
            reject(self, e);
        }
    }

    function reject(self, newValue) {
        self._state = 2;
        self._value = newValue;
        finale(self);
    }

    function finale(self) {
        if (self._state === 2 && self._deferreds.length === 0) {
            Promise._immediateFn(function () {
                if (!self._handled) {
                    Promise._unhandledRejectionFn(self._value);
                }
            });
        }

        for (var i = 0, len = self._deferreds.length; i < len; i++) {
            handle(self, self._deferreds[i]);
        }
        self._deferreds = null;
    }

    function Handler(onFulfilled, onRejected, promise) {
        this.onFulfilled = typeof onFulfilled === 'function' ? onFulfilled : null;
        this.onRejected = typeof onRejected === 'function' ? onRejected : null;
        this.promise = promise;
    }

    /**
     * Take a potentially misbehaving resolver function and make sure
     * onFulfilled and onRejected are only called once.
     *
     * Makes no guarantees about asynchrony.
     */
    function doResolve(fn, self) {
        var done = false;
        try {
            fn(function (value) {
                if (done) return;
                done = true;
                resolve(self, value);
            }, function (reason) {
                if (done) return;
                done = true;
                reject(self, reason);
            });
        } catch (ex) {
            if (done) return;
            done = true;
            reject(self, ex);
        }
    }

    Promise.prototype['catch'] = function (onRejected) {
        return this.then(null, onRejected);
    };

    Promise.prototype.then = function (onFulfilled, onRejected) {
        var prom = new (this.constructor)(noop);

        handle(this, new Handler(onFulfilled, onRejected, prom));
        return prom;
    };

    Promise.all = function (arr) {
        var args = Array.prototype.slice.call(arr);

        return new Promise(function (resolve, reject) {
            if (args.length === 0) return resolve([]);
            var remaining = args.length;

            function res(i, val) {
                try {
                    if (val && (typeof val === 'object' || typeof val === 'function')) {
                        var then = val.then;
                        if (typeof then === 'function') {
                            then.call(val, function (val) {
                                res(i, val);
                            }, reject);
                            return;
                        }
                    }
                    args[i] = val;
                    if (--remaining === 0) {
                        resolve(args);
                    }
                } catch (ex) {
                    reject(ex);
                }
            }

            for (var i = 0; i < args.length; i++) {
                res(i, args[i]);
            }
        });
    };

    Promise.resolve = function (value) {
        if (value && typeof value === 'object' && value.constructor === Promise) {
            return value;
        }

        return new Promise(function (resolve) {
            resolve(value);
        });
    };

    Promise.reject = function (value) {
        return new Promise(function (resolve, reject) {
            reject(value);
        });
    };

    Promise.race = function (values) {
        return new Promise(function (resolve, reject) {
            for (var i = 0, len = values.length; i < len; i++) {
                values[i].then(resolve, reject);
            }
        });
    };

    // Use polyfill for setImmediate for performance gains
    Promise._immediateFn = (typeof setImmediate === 'function' && function (fn) { setImmediate(fn); }) ||
        function (fn) {
            setTimeoutFunc(fn, 0);
        };

    Promise._unhandledRejectionFn = function _unhandledRejectionFn(err) {
        if (typeof console !== 'undefined' && console) {
            console.warn('Possible Unhandled Promise Rejection:', err); // eslint-disable-line no-console
        }
    };

    /**
     * Set the immediate function to execute callbacks
     * @param fn {function} Function to execute
     * @deprecated
     */
    Promise._setImmediateFn = function _setImmediateFn(fn) {
        Promise._immediateFn = fn;
    };

    /**
     * Change the function to execute on unhandled rejection
     * @param {function} fn Function to execute on unhandled rejection
     * @deprecated
     */
    Promise._setUnhandledRejectionFn = function _setUnhandledRejectionFn(fn) {
        Promise._unhandledRejectionFn = fn;
    };

    if (typeof module !== 'undefined' && module.exports) {
        module.exports = Promise;
    } else if (!root.Promise) {
        root.Promise = Promise;
    }

})(this);

// Fetch Polyfill to provide a wrapper around the XMLHttpRequest functionalities
// Taken from https://github.com/github/fetch
(function (self) {
    'use strict';

    if (self.fetch) {
        return
    }

    var support = {
        searchParams: 'URLSearchParams' in self,
        iterable: 'Symbol' in self && 'iterator' in Symbol,
        blob: 'FileReader' in self && 'Blob' in self && (function () {
            try {
                new Blob()
                return true
            } catch (e) {
                return false
            }
        })(),
        formData: 'FormData' in self,
        arrayBuffer: 'ArrayBuffer' in self
    }

    if (support.arrayBuffer) {
        var viewClasses = [
            '[object Int8Array]',
            '[object Uint8Array]',
            '[object Uint8ClampedArray]',
            '[object Int16Array]',
            '[object Uint16Array]',
            '[object Int32Array]',
            '[object Uint32Array]',
            '[object Float32Array]',
            '[object Float64Array]'
        ]

        var isDataView = function (obj) {
            return obj && DataView.prototype.isPrototypeOf(obj)
        }

        var isArrayBufferView = ArrayBuffer.isView || function (obj) {
            return obj && viewClasses.indexOf(Object.prototype.toString.call(obj)) > -1
        }
    }

    function normalizeName(name) {
        if (typeof name !== 'string') {
            name = String(name)
        }
        if (/[^a-z0-9\-#$%&'*+.\^_`|~]/i.test(name)) {
            throw new TypeError('Invalid character in header field name')
        }
        return name.toLowerCase()
    }

    function normalizeValue(value) {
        if (typeof value !== 'string') {
            value = String(value)
        }
        return value
    }

    // Build a destructive iterator for the value list
    function iteratorFor(items) {
        var iterator = {
            next: function () {
                var value = items.shift()
                return { done: value === undefined, value: value }
            }
        }

        if (support.iterable) {
            iterator[Symbol.iterator] = function () {
                return iterator
            }
        }

        return iterator
    }

    function Headers(headers) {
        this.map = {}

        if (headers instanceof Headers) {
            headers.forEach(function (value, name) {
                this.append(name, value)
            }, this)

        } else if (headers) {
            Object.getOwnPropertyNames(headers).forEach(function (name) {
                this.append(name, headers[name])
            }, this)
        }
    }

    Headers.prototype.append = function (name, value) {
        name = normalizeName(name)
        value = normalizeValue(value)
        var oldValue = this.map[name]
        this.map[name] = oldValue ? oldValue + ',' + value : value
    }

    Headers.prototype['delete'] = function (name) {
        delete this.map[normalizeName(name)]
    }

    Headers.prototype.get = function (name) {
        name = normalizeName(name)
        return this.has(name) ? this.map[name] : null
    }

    Headers.prototype.has = function (name) {
        return this.map.hasOwnProperty(normalizeName(name))
    }

    Headers.prototype.set = function (name, value) {
        this.map[normalizeName(name)] = normalizeValue(value)
    }

    Headers.prototype.forEach = function (callback, thisArg) {
        for (var name in this.map) {
            if (this.map.hasOwnProperty(name)) {
                callback.call(thisArg, this.map[name], name, this)
            }
        }
    }

    Headers.prototype.keys = function () {
        var items = []
        this.forEach(function (value, name) { items.push(name) })
        return iteratorFor(items)
    }

    Headers.prototype.values = function () {
        var items = []
        this.forEach(function (value) { items.push(value) })
        return iteratorFor(items)
    }

    Headers.prototype.entries = function () {
        var items = []
        this.forEach(function (value, name) { items.push([name, value]) })
        return iteratorFor(items)
    }

    if (support.iterable) {
        Headers.prototype[Symbol.iterator] = Headers.prototype.entries
    }

    function consumed(body) {
        if (body.bodyUsed) {
            return Promise.reject(new TypeError('Already read'))
        }
        body.bodyUsed = true
    }

    function fileReaderReady(reader) {
        return new Promise(function (resolve, reject) {
            reader.onload = function () {
                resolve(reader.result)
            }
            reader.onerror = function () {
                reject(reader.error)
            }
        })
    }

    function readBlobAsArrayBuffer(blob) {
        var reader = new FileReader()
        var promise = fileReaderReady(reader)
        reader.readAsArrayBuffer(blob)
        return promise
    }

    function readBlobAsText(blob) {
        var reader = new FileReader()
        var promise = fileReaderReady(reader)
        reader.readAsText(blob)
        return promise
    }

    function readArrayBufferAsText(buf) {
        var view = new Uint8Array(buf)
        var chars = new Array(view.length)

        for (var i = 0; i < view.length; i++) {
            chars[i] = String.fromCharCode(view[i])
        }
        return chars.join('')
    }

    function bufferClone(buf) {
        if (buf.slice) {
            return buf.slice(0)
        } else {
            var view = new Uint8Array(buf.byteLength)
            view.set(new Uint8Array(buf))
            return view.buffer
        }
    }

    function Body() {
        this.bodyUsed = false

        this._initBody = function (body) {
            this._bodyInit = body
            if (!body) {
                this._bodyText = ''
            } else if (typeof body === 'string') {
                this._bodyText = body
            } else if (support.blob && Blob.prototype.isPrototypeOf(body)) {
                this._bodyBlob = body
            } else if (support.formData && FormData.prototype.isPrototypeOf(body)) {
                this._bodyFormData = body
            } else if (support.searchParams && URLSearchParams.prototype.isPrototypeOf(body)) {
                this._bodyText = body.toString()
            } else if (support.arrayBuffer && support.blob && isDataView(body)) {
                this._bodyArrayBuffer = bufferClone(body.buffer)
                // IE 10-11 can't handle a DataView body.
                this._bodyInit = new Blob([this._bodyArrayBuffer])
            } else if (support.arrayBuffer && (ArrayBuffer.prototype.isPrototypeOf(body) || isArrayBufferView(body))) {
                this._bodyArrayBuffer = bufferClone(body)
            } else {
                throw new Error('unsupported BodyInit type')
            }

            if (!this.headers.get('content-type')) {
                if (typeof body === 'string') {
                    this.headers.set('content-type', 'text/plain;charset=UTF-8')
                } else if (this._bodyBlob && this._bodyBlob.type) {
                    this.headers.set('content-type', this._bodyBlob.type)
                } else if (support.searchParams && URLSearchParams.prototype.isPrototypeOf(body)) {
                    this.headers.set('content-type', 'application/x-www-form-urlencoded;charset=UTF-8')
                }
            }
        }

        if (support.blob) {
            this.blob = function () {
                var rejected = consumed(this)
                if (rejected) {
                    return rejected
                }

                if (this._bodyBlob) {
                    return Promise.resolve(this._bodyBlob)
                } else if (this._bodyArrayBuffer) {
                    return Promise.resolve(new Blob([this._bodyArrayBuffer]))
                } else if (this._bodyFormData) {
                    throw new Error('could not read FormData body as blob')
                } else {
                    return Promise.resolve(new Blob([this._bodyText]))
                }
            }

            this.arrayBuffer = function () {
                if (this._bodyArrayBuffer) {
                    return consumed(this) || Promise.resolve(this._bodyArrayBuffer)
                } else {
                    return this.blob().then(readBlobAsArrayBuffer)
                }
            }
        }

        this.text = function () {
            var rejected = consumed(this)
            if (rejected) {
                return rejected
            }

            if (this._bodyBlob) {
                return readBlobAsText(this._bodyBlob)
            } else if (this._bodyArrayBuffer) {
                return Promise.resolve(readArrayBufferAsText(this._bodyArrayBuffer))
            } else if (this._bodyFormData) {
                throw new Error('could not read FormData body as text')
            } else {
                return Promise.resolve(this._bodyText)
            }
        }

        if (support.formData) {
            this.formData = function () {
                return this.text().then(decode)
            }
        }

        this.json = function () {
            return this.text().then(JSON.parse)
        }

        return this
    }

    // HTTP methods whose capitalization should be normalized
    var methods = ['DELETE', 'GET', 'HEAD', 'OPTIONS', 'POST', 'PUT']

    function normalizeMethod(method) {
        var upcased = method.toUpperCase()
        return (methods.indexOf(upcased) > -1) ? upcased : method
    }

    function Request(input, options) {
        options = options || {}
        var body = options.body

        if (input instanceof Request) {
            if (input.bodyUsed) {
                throw new TypeError('Already read')
            }
            this.url = input.url
            this.credentials = input.credentials
            if (!options.headers) {
                this.headers = new Headers(input.headers)
            }
            this.method = input.method
            this.mode = input.mode
            if (!body && input._bodyInit != null) {
                body = input._bodyInit
                input.bodyUsed = true
            }
        } else {
            this.url = String(input)
        }

        this.credentials = options.credentials || this.credentials || 'omit'
        if (options.headers || !this.headers) {
            this.headers = new Headers(options.headers)
        }
        this.method = normalizeMethod(options.method || this.method || 'GET')
        this.mode = options.mode || this.mode || null
        this.referrer = null

        if ((this.method === 'GET' || this.method === 'HEAD') && body) {
            throw new TypeError('Body not allowed for GET or HEAD requests')
        }
        this._initBody(body)
    }

    Request.prototype.clone = function () {
        return new Request(this, { body: this._bodyInit })
    }

    function decode(body) {
        var form = new FormData()
        body.trim().split('&').forEach(function (bytes) {
            if (bytes) {
                var split = bytes.split('=')
                var name = split.shift().replace(/\+/g, ' ')
                var value = split.join('=').replace(/\+/g, ' ')
                form.append(decodeURIComponent(name), decodeURIComponent(value))
            }
        })
        return form
    }

    function parseHeaders(rawHeaders) {
        var headers = new Headers()
        rawHeaders.split(/\r?\n/).forEach(function (line) {
            var parts = line.split(':')
            var key = parts.shift().trim()
            if (key) {
                var value = parts.join(':').trim()
                headers.append(key, value)
            }
        })
        return headers
    }

    Body.call(Request.prototype)

    function Response(bodyInit, options) {
        if (!options) {
            options = {}
        }

        this.type = 'default'
        this.status = 'status' in options ? options.status : 200
        this.ok = this.status >= 200 && this.status < 300
        this.statusText = 'statusText' in options ? options.statusText : 'OK'
        this.headers = new Headers(options.headers)
        this.url = options.url || ''
        this._initBody(bodyInit)
    }

    Body.call(Response.prototype)

    Response.prototype.clone = function () {
        return new Response(this._bodyInit, {
            status: this.status,
            statusText: this.statusText,
            headers: new Headers(this.headers),
            url: this.url
        })
    }

    Response.error = function () {
        var response = new Response(null, { status: 0, statusText: '' })
        response.type = 'error'
        return response
    }

    var redirectStatuses = [301, 302, 303, 307, 308]

    Response.redirect = function (url, status) {
        if (redirectStatuses.indexOf(status) === -1) {
            throw new RangeError('Invalid status code')
        }

        return new Response(null, { status: status, headers: { location: url } })
    }

    self.Headers = Headers
    self.Request = Request
    self.Response = Response

    self.fetch = function (input, init) {
        return new Promise(function (resolve, reject) {
            var request = new Request(input, init)
            var xhr = new XMLHttpRequest()

            xhr.onload = function () {
                var options = {
                    status: xhr.status,
                    statusText: xhr.statusText,
                    headers: parseHeaders(xhr.getAllResponseHeaders() || '')
                }
                options.url = 'responseURL' in xhr ? xhr.responseURL : options.headers.get('X-Request-URL')
                var body = 'response' in xhr ? xhr.response : xhr.responseText
                resolve(new Response(body, options))
            }

            xhr.onerror = function () {
                reject(new TypeError('Network request failed'))
            }

            xhr.ontimeout = function () {
                reject(new TypeError('Network request failed'))
            }

            xhr.open(request.method, request.url, true)

            if (request.credentials === 'include') {
                xhr.withCredentials = true
            }

            if ('responseType' in xhr && support.blob) {
                xhr.responseType = 'blob'
            }

            request.headers.forEach(function (value, name) {
                xhr.setRequestHeader(name, value)
            })

            xhr.send(typeof request._bodyInit === 'undefined' ? null : request._bodyInit)
        })
    }
    self.fetch.polyfill = true
})(typeof self !== 'undefined' ? self : this);

// Polyfill to add array forEach support if it does not yet exist on the page
// Production steps of ECMA-262, Edition 5, 15.4.4.18
// Reference: http://es5.github.io/#x15.4.4.18
if (!Array.prototype.forEach) {

    Array.prototype.forEach = function (callback, thisArg) {

        var T, k;

        if (this === null) {
            throw new TypeError('this is null or not defined');
        }

        // 1. Let O be the result of calling toObject() passing the
        // |this| value as the argument.
        var O = Object(this);

        // 2. Let lenValue be the result of calling the Get() internal
        // method of O with the argument "length".
        // 3. Let len be toUint32(lenValue).
        var len = O.length >>> 0;

        // 4. If isCallable(callback) is false, throw a TypeError exception. 
        // See: http://es5.github.com/#x9.11
        if (typeof callback !== 'function') {
            throw new TypeError(callback + ' is not a function');
        }

        // 5. If thisArg was supplied, let T be thisArg; else let
        // T be undefined.
        if (arguments.length > 1) {
            T = thisArg;
        }

        // 6. Let k be 0
        k = 0;

        // 7. Repeat, while k < len
        while (k < len) {

            var kValue;

            // a. Let Pk be ToString(k).
            //    This is implicit for LHS operands of the in operator
            // b. Let kPresent be the result of calling the HasProperty
            //    internal method of O with argument Pk.
            //    This step can be combined with c
            // c. If kPresent is true, then
            if (k in O) {

                // i. Let kValue be the result of calling the Get internal
                // method of O with argument Pk.
                kValue = O[k];

                // ii. Call the Call internal method of callback with T as
                // the this value and argument list containing kValue, k, and O.
                callback.call(T, kValue, k, O);
            }
            // d. Increase k by 1.
            k++;
        }
        // 8. return undefined
    };
}

// Define a polyfill for the Object.keys function which allows for finding property keys
// From https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Object/keys
if (!Object.keys) {
    Object.keys = (function () {
        'use strict';
        var hasOwnProperty = Object.prototype.hasOwnProperty,
            hasDontEnumBug = !({ toString: null }).propertyIsEnumerable('toString'),
            dontEnums = [
                'toString',
                'toLocaleString',
                'valueOf',
                'hasOwnProperty',
                'isPrototypeOf',
                'propertyIsEnumerable',
                'constructor'
            ],
            dontEnumsLength = dontEnums.length;

        return function (obj) {
            if (typeof obj !== 'object' && (typeof obj !== 'function' || obj === null)) {
                throw new TypeError('Object.keys called on non-object');
            }

            var result = [], prop, i;

            for (prop in obj) {
                if (hasOwnProperty.call(obj, prop)) {
                    result.push(prop);
                }
            }

            if (hasDontEnumBug) {
                for (i = 0; i < dontEnumsLength; i++) {
                    if (hasOwnProperty.call(obj, dontEnums[i])) {
                        result.push(dontEnums[i]);
                    }
                }
            }
            return result;
        };
    }());
}
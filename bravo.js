"use strict";

/*
* Title: Bravo Core Library
* Source: https://github.com/bravocg/core
* Version: v1.6.2
* Author: Gunjan Datta
* Description: The Bravo core library translates the REST api as an object model.
* 
* Copyright © 2015 Bravo Consulting Group, LLC (Bravo). All Rights Reserved.
* Released under the MIT license.
*/

// Global variable
var BRAVO = BRAVO || {};

// **********************************************************************************
// Bravo Core Class
// This class converts the REST api as an object model.
// **********************************************************************************
BRAVO.Core = function () {
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
                    { name: "update", "function": function (data) { return this.executePost(null, null, data, true, this.__metadata.type, "MERGE"); } }
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
                    { name: "getListByTitle", "function": function (title) { title = encodeURIComponent(title); return this.executeGet("lists?$filter=Title eq '" + title + "'"); } },
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
            var metadataType = obj.__metadata.type;

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
            var promise = new BRAVO.Core.Promise();

            // Execute the request
            executeMethod(this, methodType, metadataType, funcName, data, headers).done(function (obj, response) {
                // Process the response and resolve the promise
                promise.resolve(processRequest(obj, response));
            });

            // Return the promise
            return promise;
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
                var promise = new BRAVO.Core.Promise();

                // Set the response type
                xhr.responseType = "arraybuffer";

                // Set the state change event
                xhr.onreadystatechange = function () {
                    // Ensure the request was successful
                    if (xhr.readyState == 4 && xhr.status == "200") {
                        // Resolve the promise
                        promise.resolve(obj, xhr.response);
                    }
                }

                // Execute the request
                xhr.send();

                // Return the promise
                return promise;
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
                var promise = new BRAVO.Core.Promise();

                // Set the state change event
                xhr.onreadystatechange = function () {
                    // See if the request has finished
                    if (xhr.readyState == 4) {
                        // Resolve the promise
                        promise.resolve(obj, xhr.response);
                    }
                }

                // Execute the request
                xhr.send(data);

                // Return the promise
                return promise;
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
                if (propertyValue == value) { return result; }
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
            var promise = new BRAVO.Core.Promise();

            // Execute the request
            executeRequest(obj, obj.RequestUrl).done(function (obj, response) {
                // Set the response
                obj.Response = JSON.parse(response);

                // Update the properties
                updateProperties(obj, obj.Response.d);

                // Resolve the promise
                promise.resolve(obj);
            });

            // Return the promise
            return promise;
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
        // Ensure the value exists
        if (value === undefined) { return; }

        // See if an update is needed
        if (this[name] != null && this[name] == value) { return; }

        // Store the current setting, and disable the asynchronous flag
        var isAsync = this.asyncFl;
        this.asyncFl = false;

        // Create the data to update the property
        var data = '{ "' + name + '": ' + (typeof (value) == "string" ? '"' + value + '"' : value) + ' }';

        // Execute the request
        var response = executeMethod(this, "MERGE", this.__metadata.type, null, JSON.parse(data), { "IF-MATCH": "*" });

        // Reset the asynchronous flag
        this.asyncFl = isAsync;

        // See if the update was successful
        if (response == "") {
            // Update the property in the current object
            this[name] = value;

            // Property was updated successfully
            return true;
        }

        // Error setting the property
        return false;
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
                // See if a result has been passed
                if (arguments.length == 4 && typeof (arguments[3]) === "object") {
                    // Set the response
                    obj.Response = arguments[3];

                    // Update the properties
                    updateProperties(obj, obj.Response.d);
                }
                else {
                    // Set the asynchronous flag
                    obj.asyncFl = arguments[2] && typeof (arguments[2]) === "boolean" ? arguments[2] : false;

                    // Determine the method type
                    var methodType = arguments[3] && typeof (arguments[3] === "string") ? arguments[3] : "GET";

                    // See if we are making an asynchronous request
                    if (obj.asyncFl) {
                        var promise = new BRAVO.Core.Promise();

                        // Execute the request
                        executeRequest(obj, obj.RequestUrl, methodType).done(function (obj, response) {
                            // Set the response
                            obj.Response = JSON.parse(response);

                            // Update the properties
                            updateProperties(obj, obj.Response.d);

                            // Resolve the promise
                            promise.resolve(obj);
                        });

                        // Return the promise
                        return promise;
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
        Promise: function () {
            return {
                _arguments: null,
                _callback: null,
                _resolveFl: false,
                done: function (callback) { this._callback = callback; if (this._callback && this._resolveFl) { this._callback.apply(this, this._arguments); } },
                resolve: function () { this._arguments = arguments; this._resolveFl = true; if (this._callback) { this._callback.apply(this, this._arguments); } }
            };
        },

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

// Ensure the sp class is loaded 
SP.SOD.executeFunc("sp.js", "SP.ClientContext", function () {
    // Notify scripts that this class is loaded
    SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("bravo.js");
});
// The BRAVO JSLink class
BRAVO.JSLink = function () {
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
            if (list.exists) {
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
            var listId = ctx && ctx.FormContext ? ctx.FormContext.listAttributes.Id : null;
            var listTitle = ctx && ctx.ListTitle ? ctx.ListTitle : ctx.ListName;

            // Get the list
            _currentList = listId ? web.getListById(listId) : web.getListByTitle(listTitle);
        }

        // Return the current list
        return _currentList;
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

    // Method to get the default field html
    var getFieldDefaultHtml = function (ctx, field, formType) {
        // Determine the field type
        var fieldType = field ? field.Type : (ctx.CurrentFieldSchema ? ctx.CurrentFieldSchema.Type : null);

        // Ensure the form type is set
        formType = formType ? formType : null;
        if (formType == null) {
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
        }

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

        // Return the item's field value html
        return fieldRenderer ? fieldRenderer.RenderField(ctx, field, ctx.CurrentItem, ctx.ListSchema) : ctx.CurrentItem[field.Name];
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

    // Method to get the list view element
    var getListView = function (ctx) { return document.querySelector("#WebPart" + (ctx.FormUniqueId || ctx.wpq) + " .ms-formtable"); }

    // Method to get the view's list items
    var getViewListItems = function (ctx) { return ctx.ListData ? ctx.ListData.Row : []; };

    // Method to get the view's selected items
    var getViewSelectedItems = function (ctx) { return SP.ListOperation.Selection.getSelectedItems(); };

    // Method to hide the row containing the field
    var hideField = function (ctx) {
        // Get the list view
        var listView = getListView(ctx);
        if (listView && listView.getAttribute("data-hide-row-event") == null) {
            // Add an onload event to it
            window.addEventListener("load", function () {
                // Query for the elements we need to hide
                var fieldElements = document.querySelectorAll(".hide-row");
                for (var i = 0; i < fieldElements.length; i++) {
                    // Find the parent row
                    var fieldElement = fieldElements[i];
                    while (fieldElement && fieldElement.nodeName.toLowerCase() != "tr") { fieldElement = fieldElement.parentNode; }

                    // Hide the row
                    if (fieldElement) { fieldElement.style.display = "none"; }
                }
            });

            // Set the attribute
            listView.setAttribute("data-hide-row-event", true);
        }

        // Create an empty element
        return "<div class='hide-row' />";
    }

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
        hideField: hideField,
        inEditMode: inEditMode,
        sendEmail: sendEmail
    };
}();

// Notify scripts that this class is loaded
SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("bravo.jslink.js");

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
        // See if the dialog result is defined and is a number
        if (dialogResult != null && typeof (dialogResult) === "number") {
            // Close the dialog
            SP.UI.ModalDialog.commonModalDialogClose(dialogResult, returnVal);
        } else {
            // Close the dialog with the 'Cancel' result
            SP.UI.ModalDialog.commonModalDialogClose(SP.UI.DialogResult.cancel, returnVal);
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

// Ensure the dialog class is loaded
SP.SOD.executeFunc("sp.ui.dialog.js", "SP.UI.ModalDialog", function () {
    // Notify scripts that this class is loaded
    SP.SOD.notifyScriptLoadedAndExecuteWaitingJobs("bravo.modaldialog.js");
});
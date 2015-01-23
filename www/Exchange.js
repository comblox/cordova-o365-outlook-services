// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

var Users = require('./Users').Users;
var UserFetcher = require('./Users').UserFetcher;

var Exchange = {
};

Exchange.Client = Client;

function DataContext (serviceRootUri, getAccessTokenFn) {
    this.serviceRootUri = serviceRootUri;
    this.getAccessTokenFn = getAccessTokenFn;
}

function Client(serviceRootUri, getAccessTokenFn) {
    this.context = new DataContext(serviceRootUri, getAccessTokenFn);
}

Client.prototype.getPath = function (prop) {
    return this.context.serviceRootUri + '/' + prop;
};

Object.defineProperty(Client.prototype, "me", {
    get: function () {
        if (this._me === undefined) {
            this._me = new UserFetcher(this.context, this.getPath("Me"), "me");
        }
        return this._me;
    },
    enumerable: true,
    configurable: true
});

Object.defineProperty(Client.prototype, "users", {
    get: function () {
        if (this._users === undefined) {
            this._users = new Users(this.context, this.getPath("Users"));
        }
        return this._users;
    },
    enumerable: true,
    configurable: true
});

module.exports = Exchange;

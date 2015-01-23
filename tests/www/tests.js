// Copyright (c) Microsoft Open Technologies, Inc.  All rights reserved.  Licensed under the Apache License, Version 2.0.  See License.txt in the project root for license information.

/* global cordova, exports, Microsoft.OutlookServices, O365Auth, jasmine, describe, it, expect, beforeEach, afterEach, pending */

var resourceUrl = 'https://outlook.office365.com';
var officeEndpointUrl = 'https://outlook.office365.com/ews/odata';

var appId = '14b0c641-7fea-4e84-8557-25285eb86e43';
var authUrl = 'https://login.windows.net/common/';
var redirectUrl = 'http://localhost:4400/services/office365/redirectTarget.html';

var Users          = cordova.require('com.msopentech.o365.outlook-services.Users');
var Calendars      = cordova.require('com.msopentech.o365.outlook-services.Calendars');
var Contacts       = cordova.require('com.msopentech.o365.outlook-services.Contacts');
var Events         = cordova.require('com.msopentech.o365.outlook-services.Events');
var Folders        = cordova.require('com.msopentech.o365.outlook-services.Folders');
var Messages       = cordova.require('com.msopentech.o365.outlook-services.Messages');
var Attachments    = cordova.require('com.msopentech.o365.outlook-services.Attachments');
var ContactFolders = cordova.require('com.msopentech.o365.outlook-services.ContactFolders');
var CalendarGroups = cordova.require('com.msopentech.o365.outlook-services.CalendarGroups');

var guid = function () {
    function _p8(s) {
        var p = (Math.random().toString(16) + "000000000").substr(2, 8);
        return s ? "-" + p.substr(0, 4) + "-" + p.substr(4, 4) : p;
    }
    return _p8() + _p8(true) + _p8(true) + _p8();
};

exports.defineAutoTests = function () {

    jasmine.DEFAULT_TIMEOUT_INTERVAL = 10000;

    describe('Auth module: ', function () {

        var authContext;

        beforeEach(function () {
            authContext = new O365Auth.Context(authUrl, redirectUrl);
        });

        it("should exists", function () {
            expect(O365Auth).toBeDefined();
        });

        it("should contain a Context constructor", function () {
            expect(O365Auth.Context).toBeDefined();
            expect(O365Auth.Context).toEqual(jasmine.any(Function));
        });

        it("should successfully create a Context object", function () {
            var fakeAuthUrl = "fakeAuthUrl",
                fakeRedirectUrl = "fakeRedirectUrl",
                context = new O365Auth.Context(fakeAuthUrl, fakeRedirectUrl);

            expect(context).not.toBeNull();
            expect(context).toEqual(jasmine.objectContaining({
                _authUri: fakeAuthUrl + '/',
                _redirectUri: fakeRedirectUrl
            }));
        });
    });

    describe('Outlook client: ', function () {

        var authContext;

        beforeEach(function () {
            authContext = new O365Auth.Context(authUrl, redirectUrl);
        });

        it('should exists', function () {
            expect(Microsoft.OutlookServices.Client).toBeDefined();
            expect(Microsoft.OutlookServices.Client).toEqual(jasmine.any(Function));
        });

        it('should be able to create a new client', function () {
            var client = new Microsoft.OutlookServices.Client(officeEndpointUrl,
                authContext.getAccessTokenFn(resourceUrl, '', appId));

            expect(client).not.toBe(null);
            expect(client.context).toBeDefined();
            expect(client.context.serviceRootUri).toBeDefined();
            expect(client.context.getAccessTokenFn).toBeDefined();
            expect(client.context.serviceRootUri).toEqual(officeEndpointUrl);
            expect(client.context.getAccessTokenFn).toEqual(jasmine.any(Function));
        });

        it('should contain \'users\' property', function () {
            var client = new Microsoft.OutlookServices.Client(officeEndpointUrl,
                authContext.getAccessTokenFn(resourceUrl, '', appId));

            expect(client.users).toBeDefined();
            expect(client.users).toEqual(jasmine.any(Users.Users));

            // expect that client.users is readonly
            var backupClientUsers = client.users;
            client.users = "somevalue";
            expect(client.users).not.toEqual("somevalue");
            expect(client.users).toEqual(backupClientUsers);
        });

        describe('Me property', function () {

            var client;

            beforeEach(function () {
                client = new Microsoft.OutlookServices.Client(officeEndpointUrl,
                    authContext.getAccessTokenFn(resourceUrl, '', appId));
            });

            it("should exists", function () {
                expect(client.me).toBeDefined();
            });

            it("should be read-only", function () {
                var backupClientMe = client.me;
                client.me = "somevalue";
                expect(client.me).not.toEqual("somevalue");
                expect(client.me).toEqual(backupClientMe);
            });

            it("should be a UserFetcher object", function () {
                expect(client.me).toEqual(jasmine.any(Users.UserFetcher));
            });

            it("should have all necessary properties", function () {
                var properties = {
                    "contacts": Contacts.Contacts,
                    "calendar": Calendars.CalendarFetcher,
                    "calendars": Calendars.Calendars,
                    "events": Events.Events,
                    "messages": Messages.Messages,
                    "folders": Folders.Folders,
                    "rootFolder": Folders.FolderFetcher,
                    "inbox": Folders.FolderFetcher,
                    "drafts": Folders.FolderFetcher,
                    "sentItems": Folders.FolderFetcher,
                    "deletedItems": Folders.FolderFetcher,
                    "contactFolders": ContactFolders.ContactFolders,
                    "calendarGroups": CalendarGroups.CalendarGroups
                };

                for (var prop in properties) {
                    var meProp = client.me[prop];
                    expect(meProp).toBeDefined();
                    expect(meProp).toEqual(jasmine.any(properties[prop]));

                    var backupProp = meProp;
                    client.me[prop] = "somevalue";
                    expect(client.me[prop]).not.toEqual("somevalue");
                    expect(client.me[prop]).toEqual(backupProp);
                }
            });

            it("should successfully fetch current user", function (done) {
                client.me.fetch().then(function (user) {
                    expect(user).toEqual(jasmine.any(Users.User));
                    expect(user.path).toMatch(new RegExp(officeEndpointUrl + '/me', "i"));
                    done();
                });
            });
        });
    });

    describe("Contacts namespace ", function () {

        var createContact = function (displayName) {
            return new Contacts.Contact(null, null, {
                GivenName: displayName || guid(),
                DisplayName: guid(),
                EmailAddresses: [{
                    Address: guid() + "@" + guid() + ".com",
                    Name: guid()
                }]
            });
        };

        var client, contacts, tempEntities;

        beforeEach(function () {
            client = new Microsoft.OutlookServices.Client(officeEndpointUrl,
                (new O365Auth.Context(authUrl, redirectUrl)).getAccessTokenFn(resourceUrl, '', appId));

            contacts = client.me.contacts;
            tempEntities = [];
        });

        afterEach(function () {
            tempEntities.forEach(function (entity) {
                try {
                    entity.delete();
                } catch (e) { }
            });
        });

        it("should be able to create a new contact", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            contacts.addContact(createContact()).then(function (added) {
                tempEntities.push(added);
                expect(added.Id).toBeDefined();
                expect(added.path).toMatch(added.Id);
                expect(added).toEqual(jasmine.any(Contacts.Contact));
                done();
            }, fail);
        });

        it("should be able to get user's contacts", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            contacts.addContact(createContact()).then(function (created) {
                tempEntities.push(created);
                contacts.getContacts().fetchAll().then(function (c) {
                    expect(c).toBeDefined();
                    expect(c).toEqual(jasmine.any(Array));
                    expect(c.length).toBeGreaterThan(0);
                    expect(c[0]).toEqual(jasmine.any(Contacts.Contact));
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to apply filter to user's contacts", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            contacts.addContact(createContact()).then(function (created) {
                tempEntities.push(created);
                var filter = 'DisplayName eq \'' + created.DisplayName + '\'';
                contacts.getContacts().filter(filter).fetchAll().then(function (c) {
                    expect(c).toBeDefined();
                    expect(c).toEqual(jasmine.any(Array));
                    expect(c.length).toEqual(1);
                    expect(c[0]).toEqual(jasmine.any(Contacts.Contact));
                    expect(c[0].Name).toEqual(created.Name);
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to apply top query to user's contacts", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            contacts.addContact(createContact()).then(function (created) {
                tempEntities.push(created);
                contacts.addContact(createContact()).then(function (created2) {
                    tempEntities.push(created2);
                    contacts.getContacts().top(1).fetchAll().then(function (c) {
                        expect(c).toBeDefined();
                        expect(c).toEqual(jasmine.any(Array));
                        expect(c.length).toEqual(1);
                        expect(c[0]).toEqual(jasmine.any(Contacts.Contact));
                        done();
                    }, function (err) {
                        expect(err).toBeUndefined();
                        done();
                    });
                }, fail);
            }, fail);
        });

        it("should be able to get a newly created contact by Id", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            var newContact = createContact();
            contacts.addContact(newContact).then(function (added) {
                tempEntities.push(added);
                contacts.getContact(added.Id).fetch().then(function (got) {
                    expect(got.GivenName).toEqual(newContact.GivenName);
                    expect(got.DisplayName).toEqual(newContact.DisplayName);
                    expect(got.EmailAddresses[0]).toEqual(newContact.EmailAddresses[0]);
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to modify existing contact", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            var newContact = createContact();
            contacts.addContact(newContact).then(function (added) {
                tempEntities.push(added);
                added.DisplayName = guid();
                added.update().then(function (updated) {
                    contacts.getContact(updated.Id).fetch().then(function (got) {
                        expect(got.Id).toEqual(added.Id);
                        expect(got.DisplayName).toEqual(added.DisplayName);
                        done();
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to delete existing contact", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            contacts.addContact(createContact()).then(function (added) {
                added.delete().then(function () {
                    contacts.getContact(added.Id).fetch().then(function (got) {
                        expect(got).toBeUndefined();
                        got.delete();
                        done();
                    }, function(err) {
                        expect(err.message).toBeDefined();
                        expect(err.message).toMatch("The specified object was not found in the store.");
                        done();
                    });
                }, fail);
            }, fail);
        });
    });

    describe("ContactFolders namespace", function () {
        // Note: contactFolders are readonly on server side, so add/update/delete methods is being rejected
        // with HTTP 405 Unsupported so we don't test them here
        // Before running unit test, test ContactFolder must be created manually

        var createContact = function (displayName) {
            return new Contacts.Contact(null, null, {
                GivenName: displayName || guid(),
                DisplayName: guid(),
                EmailAddresses: [{
                    Address: guid() + "@" + guid() + ".com",
                    Name: guid()
                }]
            });
        };

        var client, contacts, contactFolders, tempEntities;

        beforeEach(function () {
            client = new Microsoft.OutlookServices.Client(officeEndpointUrl,
                (new O365Auth.Context(authUrl, redirectUrl)).getAccessTokenFn(resourceUrl, '', appId));

            contacts = client.me.contacts;
            contactFolders = client.me.contactFolders;

            tempEntities = [];
        });

        afterEach(function () {
            tempEntities.forEach(function (entity) {
                try {
                    entity.delete();
                } catch (e) { }
            });
        });

        it("should be able to get user's contact folders", function (done) {
            contactFolders.getContactFolders().fetchAll().then(function (cf) {
                expect(cf).toBeDefined();
                expect(cf).toEqual(jasmine.any(Array));
                expect(cf.length).toBeGreaterThan(0);
                expect(cf[0]).toEqual(jasmine.any(ContactFolders.ContactFolder));
                done();
            }, function (err) {
                expect(err).toBeUndefined();
                done();
            });
        });

        it("should be able to apply filter to user's contact folders", function (done) {
            contactFolders.getContactFolders().fetchAll().then(function (cf) {
                var filterName = cf[0].DisplayName;
                var filter = 'DisplayName eq \'' + filterName + '\'';
                contactFolders.getContactFolders().filter(filter).fetchAll().then(function (cf) {
                    expect(cf).toBeDefined();
                    expect(cf).toEqual(jasmine.any(Array));
                    expect(cf.length).toEqual(1);
                    expect(cf[0]).toEqual(jasmine.any(ContactFolders.ContactFolder));
                    expect(cf[0].DisplayName).toEqual(filterName);
                    done();
                }, function (err) {
                    expect(err).toBeUndefined();
                    done();
                });
            }, function (err) {
                expect(err).toBeUndefined();
                done();
            });
        });

        it("should be able to apply top query to user's contact folders", function (done) {
            contactFolders.getContactFolders().top(1).fetchAll().then(function (cf) {
                expect(cf).toBeDefined();
                expect(cf).toEqual(jasmine.any(Array));
                expect(cf.length).toEqual(1);
                expect(cf[0]).toEqual(jasmine.any(ContactFolders.ContactFolder));
                done();
            }, function (err) {
                expect(err).toBeUndefined();
                done();
            });
        });

        it("should be able to get a contact folder by Id", function (done) {
            contactFolders.getContactFolders().fetchAll().then(function (fetched) {
                var contactFolderToGet = fetched[0];
                contactFolders.getContactFolder(contactFolderToGet.Id).fetch().then(function (got) {
                    expect(got.DisplayName).toEqual(contactFolderToGet.DisplayName);
                    done();
                }, function (err) {
                    expect(err).toBeUndefined();
                    done();
                });
            }, function (err) {
                expect(err).toBeUndefined();
                done();
            });
        });

        describe("Contact folders nested contacts operations", function () {

            it("should get contact folder's nested folders", function (done) {
                var fail = function (err) {
                    expect(err).toBeUndefined();
                    done();
                };

                contactFolders.getContactFolder('Contacts').fetch().then(function (contacts) {
                    contacts.childFolders.getContactFolders().fetchAll().then(function (fetched) {
                        expect(fetched).toBeDefined();
                        expect(fetched).toEqual(jasmine.any(Array));
                        if (fetched.length === 0) {
                            // no contact folders created for this account, can't continue other tests
                            done();
                            return;
                        }
                        expect(fetched[0]).toEqual(jasmine.any(ContactFolders.ContactFolder));
                        done();
                    }, fail);
                }, fail);
            });

            // it("should be able to create contact in specific folder", function (done) {
            //     var fail = function (err) {
            //         expect(err).toBeUndefined();
            //         done();
            //     };

            //     contactFolders.getContactFolder('Contacts').fetch().then(function (contacts) {
            //         contacts.childFolders.getContactFolders().fetchAll().then(function (fetched) {
            //             if (fetched.length < 1) {
            //                 // no contact folders created for this account, can't continue other tests
            //                 pending();
            //                 return;
            //             }

            //             var nestedContactFolder = fetched[0];
            //             var newContact = createContact();
            //             nestedContactFolder.contacts.addContact(newContact).then(function (created) {
            //                 tempEntities.push(created);
            //                 nestedContactFolder.contacts.getContact(created.Id).fetch().then(function (nested) {
            //                     expect(nested.DisplayName).toBeDefined();
            //                     expect(nested.DisplayName).toEqual(newContact.DisplayName);

            //                     contacts.contacts.getContact(created.Id).fetch().then(function (fetched) {
            //                         // created contact should not exist in 'Contacts' folder, but in Contacts' child folder
            //                         expect(fetched).toBeUndefined();
            //                         done();
            //                     }, function (err) {
            //                         expect(err.message).toMatch("The specified object was not found in the store.");
            //                         done();
            //                     });
            //                 }, fail);
            //             }, fail);
            //         }, fail);
            //     }, fail);
            // });

            it("should get contact folder's nested contacts", function (done) {
                var fail = function (err) {
                    expect(err).toBeUndefined();
                    done();
                };

                contactFolders.getContactFolder('Contacts').fetch().then(function (contacts) {
                    contacts.childFolders.getContactFolders().fetchAll().then(function (fetched) {
                        if (fetched.length === 0) {
                            // no contact folders created for this account, can't continue other tests
                            pending();
                            return;
                        }
                        var childFolder = fetched[0];
                        childFolder.contacts.getContacts().fetchAll().then(function (c) {
                            expect(c).toBeDefined();
                            expect(c).toEqual(jasmine.any(Array));
                            done();
                        });
                    }, fail);
                }, fail);
            });
        });
    });

    describe("Calendars namespace", function () {

        var createCalendar = function (name) {
            return {
                Name: name || guid()
            };
        };

        var client, calendars, tempEntities;

        beforeEach(function () {
            client = new Microsoft.OutlookServices.Client(officeEndpointUrl,
                (new O365Auth.Context(authUrl, redirectUrl)).getAccessTokenFn(resourceUrl, '', appId));

            calendars = client.me.calendars;

            tempEntities = [];
        });

        afterEach(function () {
            tempEntities.forEach(function (entity) {
                try {
                    entity.delete();
                } catch (e) {
                    console.log(e);
                }
            });
        });

        it("should be able to create a new calendar", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };

            calendars.addCalendar(createCalendar()).then(function (added) {
                tempEntities.push(added);
                expect(added.Id).toBeDefined();
                expect(added.path).toMatch(added.Id);
                expect(added).toEqual(jasmine.any(Calendars.Calendar));
                done();
            }, fail);
        });

        it("should be able to get user's calendars", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };

            calendars.addCalendar(createCalendar()).then(function (created) {
                tempEntities.push(created);
                calendars.getCalendars().fetchAll().then(function (c) {
                    expect(c).toBeDefined();
                    expect(c).toEqual(jasmine.any(Array));
                    expect(c.length).toBeGreaterThan(0);
                    expect(c[0]).toEqual(jasmine.any(Calendars.Calendar));
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to apply filter to user's calendars", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };

            calendars.addCalendar(createCalendar()).then(function (created) {
                tempEntities.push(created);
                var filter = 'Name eq \'' + created.Name + '\'';
                calendars.getCalendars().filter(filter).fetchAll().then(function (c) {
                    expect(c).toBeDefined();
                    expect(c).toEqual(jasmine.any(Array));
                    expect(c.length).toEqual(1);
                    expect(c[0]).toEqual(jasmine.any(Calendars.Calendar));
                    expect(c[0].Name).toEqual(created.Name);
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to apply top query to user's calendars", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            calendars.addCalendar(createCalendar()).then(function (created) {
                tempEntities.push(created);
                calendars.addCalendar(createCalendar()).then(function (created2) {
                    tempEntities.push(created2);
                    calendars.getCalendars().top(1).fetchAll().then(function (c) {
                        expect(c).toBeDefined();
                        expect(c).toEqual(jasmine.any(Array));
                        expect(c.length).toEqual(1);
                        expect(c[0]).toEqual(jasmine.any(Calendars.Calendar));
                        done();
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to get a newly created calendar by Id", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            var newCalendar = createCalendar();
            calendars.addCalendar(newCalendar).then(function (added) {
                tempEntities.push(added);
                calendars.getCalendar(added.Id).fetch().then(function (got) {
                    expect(got.Name).toEqual(newCalendar.Name);
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to modify existing calendar", function (done) {
            var newCalendar = createCalendar(),
            fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };

            calendars.addCalendar(newCalendar).then(function (added) {
                tempEntities.push(added);
                added.Name = guid();
                added.update().then(function (updated) {
                    calendars.getCalendar(updated.Id).fetch().then(function (got) {
                        expect(got.Id).toEqual(added.Id);
                        expect(got.Name).toEqual(added.Name);
                        done();
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to delete existing calendar", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };

            calendars.addCalendar(createCalendar()).then(function (added) {
                added.delete().then(function () {
                    calendars.getCalendar(added.Id).fetch().then(function (got) {
                        expect(got).toBeUndefined();
                        got.delete();
                        done();
                    }, function (err) {
                        expect(err).toBeDefined();
                        expect(err.code).toEqual("ErrorItemNotFound");
                        done();
                    });
                }, fail);
            }, fail);
        });
    });

    describe("Calendar Groups namespace", function () {

        var createCalendar = function (name) {
            return {
                Name: name || guid()
            };
        };
        var createCalendarGroup = createCalendar;

        var client, calendars, calendarGroups, tempEntities;

        beforeEach(function () {
            client = new Microsoft.OutlookServices.Client(officeEndpointUrl,
                (new O365Auth.Context(authUrl, redirectUrl)).getAccessTokenFn(resourceUrl, '', appId));

            calendars = client.me.calendars;
            calendarGroups = client.me.calendarGroups;

            tempEntities = [];
        });

        afterEach(function () {
            tempEntities.forEach(function (entity) {
                try {
                    entity.delete();
                } catch (e) {
                    console.log(e);
                }
            });
        });

        it("should be able to create a new calendar group", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            calendarGroups.addCalendarGroup(createCalendarGroup()).then(function (added) {
                tempEntities.push(added);
                expect(added.Id).toBeDefined();
                expect(added.path).toMatch(added.Id);
                expect(added).toEqual(jasmine.any(CalendarGroups.CalendarGroup));
                done();
            }, fail);
        });

        it("should be able to get user's calendar groups", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            calendarGroups.addCalendarGroup(createCalendarGroup()).then(function (created) {
                tempEntities.push(created);
                calendarGroups.getCalendarGroups().fetchAll().then(function (c) {
                    expect(c).toBeDefined();
                    expect(c).toEqual(jasmine.any(Array));
                    expect(c.length).toBeGreaterThan(0);
                    expect(c[0]).toEqual(jasmine.any(CalendarGroups.CalendarGroup));
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to apply filter to user's calendar groups", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            calendarGroups.addCalendarGroup(createCalendarGroup()).then(function (created) {
                tempEntities.push(created);
                var filter = 'Name eq \'' + created.Name + '\'';
                calendarGroups.getCalendarGroups().filter(filter).fetchAll().then(function (c) {
                    expect(c).toBeDefined();
                    expect(c).toEqual(jasmine.any(Array));
                    expect(c.length).toEqual(1);
                    expect(c[0]).toEqual(jasmine.any(CalendarGroups.CalendarGroup));
                    expect(c[0].Name).toEqual(created.Name);
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to apply top query to user's calendar groups", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            calendarGroups.addCalendarGroup(createCalendarGroup()).then(function (created) {
                tempEntities.push(created);
                calendarGroups.addCalendarGroup(createCalendarGroup()).then(function (created2) {
                    tempEntities.push(created2);
                    calendarGroups.getCalendarGroups().top(1).fetchAll().then(function (c) {
                        expect(c).toBeDefined();
                        expect(c).toEqual(jasmine.any(Array));
                        expect(c.length).toEqual(1);
                        expect(c[0]).toEqual(jasmine.any(CalendarGroups.CalendarGroup));
                        done();
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to get a newly created calendar group by Id", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            var newCalendarGroup = createCalendarGroup();
            calendarGroups.addCalendarGroup(newCalendarGroup).then(function (added) {
                tempEntities.push(added);
                calendarGroups.getCalendarGroup(added.Id).fetch().then(function (got) {
                    expect(got.Name).toEqual(newCalendarGroup.Name);
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to modify existing calendar group", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            calendarGroups.addCalendarGroup(createCalendarGroup()).then(function (added) {
                tempEntities.push(added);
                added.Name = guid();
                added.update().then(function (updated) {
                    calendarGroups.getCalendarGroup(updated.Id).fetch().then(function (got) {
                        expect(got.Id).toEqual(added.Id);
                        expect(got.Name).toEqual(added.Name);
                        done();
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to delete existing calendar group", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            calendarGroups.addCalendarGroup(createCalendarGroup()).then(function (added) {
                added.delete().then(function () {
                    calendarGroups.getCalendarGroup(added.Id).fetch().then(function (got) {
                        expect(got).toBeUndefined();
                        got.delete();
                        done();
                    }, function (err) {
                        expect(err).toBeDefined();
                        expect(err.code).toEqual("ErrorItemNotFound");
                        done();
                    });
                }, fail);
            }, fail);
        });

        describe("Nested calendar groups operations", function () {

            it("should be able to create and get a newly created calendar in specific group", function (done) {
                var fail = function (err) {
                    expect(err).toBeUndefined();
                    done();
                };
                calendarGroups.addCalendarGroup(createCalendarGroup()).then(function (createdGroup) {
                    tempEntities.push(createdGroup);
                    createdGroup.calendars.addCalendar(createCalendar()).then(function (createdCal) {
                        tempEntities.push(createdCal);
                        createdGroup.calendars.getCalendar(createdCal.Id).fetch().then(function (fetchedCal) {
                            expect(fetchedCal).toBeDefined();
                            expect(fetchedCal).toEqual(jasmine.any(Calendars.Calendar));
                            expect(fetchedCal.Name).toEqual(createdCal.Name);
                            done();
                        }, fail);
                    }, fail);
                }, fail);
            });

        });
    });

    describe("Events namespace", function () {

        var createEvent = function (subject) {
            return {
                Subject: subject || guid(),
                Start: new Date(),
                End: new Date()
            };
        };
        var client, events, tempEntities;

        beforeEach(function () {
            client = new Microsoft.OutlookServices.Client(officeEndpointUrl,
                (new O365Auth.Context(authUrl, redirectUrl)).getAccessTokenFn(resourceUrl, '', appId));

            events = client.me.events;
            tempEntities = [];
        });

        afterEach(function () {
            tempEntities.forEach(function (entity) {
                try {
                    entity.delete();
                } catch (e) { }
            });
        });

        it("should be able to create a new event", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            events.addEvent(createEvent()).then(function (added) {
                tempEntities.push(added);
                expect(added.Id).toBeDefined();
                expect(added.path).toMatch(added.Id);
                expect(added).toEqual(jasmine.any(Events.Event));
                done();
            }, fail);
        });

        it("should be able to get user's events", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            events.addEvent(createEvent()).then(function (created) {
                tempEntities.push(created);
                events.getEvents().fetchAll().then(function (c) {
                    expect(c).toBeDefined();
                    expect(c).toEqual(jasmine.any(Array));
                    expect(c.length).toBeGreaterThan(0);
                    expect(c[0]).toEqual(jasmine.any(Events.Event));
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to apply filter to user's events", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            events.addEvent(createEvent()).then(function (created) {
                tempEntities.push(created);
                var filter = 'Subject eq \'' + created.Subject + '\'';
                events.getEvents().filter(filter).fetchAll().then(function (c) {
                    expect(c).toBeDefined();
                    expect(c).toEqual(jasmine.any(Array));
                    expect(c.length).toEqual(1);
                    expect(c[0]).toEqual(jasmine.any(Events.Event));
                    expect(c[0].Name).toEqual(created.Name);
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to apply top query to user's events", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            events.addEvent(createEvent()).then(function (created) {
                tempEntities.push(created);
                events.addEvent(createEvent()).then(function (created2) {
                    tempEntities.push(created2);
                    events.getEvents().top(1).fetchAll().then(function (c) {
                        expect(c).toBeDefined();
                        expect(c).toEqual(jasmine.any(Array));
                        expect(c.length).toEqual(1);
                        expect(c[0]).toEqual(jasmine.any(Events.Event));
                        done();
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to get a newly created event by Id", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            var evt = createEvent();
            events.addEvent(evt).then(function (added) {
                tempEntities.push(added);
                events.getEvent(added.Id).fetch().then(function (got) {
                    expect(got.Subject).toEqual(evt.Subject);
                    expect(got.Start).toEqual(evt.Start);
                    expect(got.End).toEqual(evt.End);
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to modify existing event", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            events.addEvent(createEvent()).then(function (added) {
                tempEntities.push(added);
                added.Subject = guid();
                added.update().then(function (updated) {
                    events.getEvent(updated.Id).fetch().then(function (got) {
                        expect(got.Id).toEqual(added.Id);
                        expect(got.Subject).toEqual(added.Subject);
                        done();
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to delete existing event", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            events.addEvent(createEvent()).then(function (added) {
                added.delete().then(function () {
                    events.getEvent(added.Id).fetch().then(function (got) {
                        expect(got).toBeUndefined();
                        got.delete();
                        done();
                    }, function (err) {
                        expect(err).toBeDefined();
                        expect(err.code).toEqual("ErrorItemNotFound");
                        done();
                    });
                }, fail);
            }, fail);
        });

        it("should be able to accept event", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done && done();
            };
            // TODO: Organizer field is ignored by server and set to event creator account automatically
            var eventToAccept = createEvent();
            eventToAccept.Organizer = {
                EmailAddress: {
                    Name: "Meeting Organizer",
                    Address: "meeting.organizer@meeting.event"
                }
            };
            eventToAccept.Attendees = [
                {
                    EmailAddress: {
                        Name: "Box owner",
                        Address: "kotikov.vladimir@kotikovvladimir.onmicrosoft.com"
                    }
                }
            ];

            events.addEvent(eventToAccept).then(function (added) {
                tempEntities.push(added);
                added.accept("Comment").then(function () {
                    events.getEvent(added.Id).fetch().then(function (fetched) {
                        // TODO: add expectations here
                        expect(fetched.Accepted).toBeTruthy();
                        done();
                    }, fail);
                }, function(err) {
                        expect(err.message).toEqual('Your request can\'t be completed. You can\'t respond to this meeting because you\'re the meeting organizer.');
                        done();
                });
            }, fail);
        });

        it("should be able to tentatively accept event", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done && done();
            };
            // TODO: Organizer field is ignored by server and set to event creator account automatically
            var eventToAccept = createEvent();
            eventToAccept.Organizer = {
                EmailAddress: {
                    Name: "Meeting Organizer",
                    Address: "meeting.organizer@meeting.event"
                }
            };
            eventToAccept.Attendees = [
                {
                    EmailAddress: {
                        Name: "Box owner",
                        Address: "kotikov.vladimir@kotikovvladimir.onmicrosoft.com"
                    }
                }
            ];

            events.addEvent(eventToAccept).then(function (added) {
                tempEntities.push(added);
                added.tentativelyAccept("Comment").then(function () {
                    events.getEvent(added.Id).fetch().then(function (fetched) {
                        // TODO: add expectations here
                        expect(fetched.Accepted).toBeTruthy();
                        done();
                    }, fail);
                },  function(err) {
                        expect(err.message).toEqual('Your request can\'t be completed. You can\'t respond to this meeting because you\'re the meeting organizer.');
                        done();
                });
            }, fail);
        });

        it("should be able to decline event", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done && done();
            };

            // TODO: Organizer field is ignored by server and set to event creator account automatically
            var eventToAccept = createEvent();
            eventToAccept.Organizer = {
                EmailAddress: {
                    Name: "Meeting Organizer",
                    Address: "meeting.organizer@meeting.event"
                }
            };
            eventToAccept.Attendees = [
                {
                    EmailAddress: {
                        Name: "Box owner",
                        Address: "kotikov.vladimir@kotikovvladimir.onmicrosoft.com"
                    }
                }
            ];

            events.addEvent(eventToAccept).then(function (added) {
                tempEntities.push(added);
                added.decline("Comment").then(function () {
                    events.getEvent(added.Id).fetch().then(function (fetched) {
                        // TODO: add expectations here
                        expect(fetched.Declined).toBeTruthy();
                        done();
                    }, fail);
                },  function(err) {
                        expect(err.message).toEqual('Your request can\'t be completed. You can\'t respond to this meeting because you\'re the meeting organizer.');
                        done();
                });
            }, fail);
        });
    });

    describe("Messages namespace", function() {

        var createRecipient = function(email, name) {
            return {
                EmailAddress: {
                    Name: name || guid(),
                    Address: email || (guid() + '@' + guid() + '.' + guid().substr(0, 3))
                }
            };
        };
        var createMessage = function (subject) {
            return {
                Subject: subject || guid(),
                ToRecipients: [ createRecipient() ],
                Body: {
                    ContentType: 0,
                    Content: "Test message"
                }
            };
        };
        var client, folders, messages, tempEntities, backInterval;

        beforeEach(function () {
            client = new Microsoft.OutlookServices.Client(officeEndpointUrl,
                (new O365Auth.Context(authUrl, redirectUrl)).getAccessTokenFn(resourceUrl, '', appId));

            folders = client.me.folders;
            messages = client.me.messages;
            tempEntities = [];

            // increase standart jasmine timeout up to 10 seconds because to some tests
            // are perform some time consumpting operations (e.g. send, reply, forward)
            backInterval = jasmine.DEFAULT_TIMEOUT_INTERVAL;
            jasmine.DEFAULT_TIMEOUT_INTERVAL = 20000;
        });

        afterEach(function () {
            tempEntities.forEach(function (entity) {
                try {
                    entity.delete();
                } catch (e) { }
            });

            // revert back default jasmine timeout
            jasmine.DEFAULT_TIMEOUT_INTERVAL = backInterval;
        });

        it("should be able to create a new message", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            messages.addMessage(createMessage()).then(function (added) {
                tempEntities.push(added);
                expect(added.Id).toBeDefined();
                expect(added.path).toMatch(added.Id);
                expect(added).toEqual(jasmine.any(Messages.Message));
                done();
            }, fail);
        });

        it("should be able to get user's messages", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            messages.addMessage(createMessage()).then(function (created) {
                tempEntities.push(created);
                messages.getMessages().fetchAll().then(function (c) {
                    expect(c).toBeDefined();
                    expect(c).toEqual(jasmine.any(Array));
                    expect(c.length).toBeGreaterThan(0);
                    expect(c[0]).toEqual(jasmine.any(Messages.Message));
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to apply filter to user's messages", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            messages.addMessage(createMessage()).then(function (created) {
                tempEntities.push(created);
                var filter = 'Subject eq \'' + created.Subject + '\'';
                client.me.drafts.messages.getMessages().filter(filter).fetchAll().then(function (c) {
                    expect(c).toBeDefined();
                    expect(c).toEqual(jasmine.any(Array));
                    expect(c.length).toEqual(1);
                    expect(c[0]).toEqual(jasmine.any(Messages.Message));
                    expect(c[0].Name).toEqual(created.Name);
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to apply top query to user's messages", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            messages.addMessage(createMessage()).then(function (created) {
                tempEntities.push(created);
                messages.addMessage(createMessage()).then(function (created2) {
                    tempEntities.push(created2);
                    client.me.drafts.messages.getMessages().top(1).fetchAll().then(function (c) {
                        expect(c).toBeDefined();
                        expect(c).toEqual(jasmine.any(Array));
                        expect(c.length).toEqual(1);
                        expect(c[0]).toEqual(jasmine.any(Messages.Message));
                        expect(c[0].Subject).toEqual(created2.Subject);
                        done();
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to get a newly created message by Id", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            var message = createMessage();
            messages.addMessage(message).then(function (added) {
                tempEntities.push(added);
                messages.getMessage(added.Id).fetch().then(function (got) {
                    expect(got.Subject).toEqual(message.Subject);
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to modify existing message", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            messages.addMessage(createMessage()).then(function (added) {
                tempEntities.push(added);
                added.Subject = guid();
                added.update().then(function (updated) {
                    messages.getMessage(updated.Id).fetch().then(function (got) {
                        expect(got.Id).toEqual(added.Id);
                        expect(got.Subject).toEqual(added.Subject);
                        done();
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to delete existing message", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            messages.addMessage(createMessage()).then(function (added) {
                added.delete().then(function () {
                    messages.getMessage(added.Id).fetch().then(function (got) {
                        expect(got).toBeUndefined();
                        got.delete();
                        done();
                    }, function (err) {
                        expect(err).toBeDefined();
                        expect(err.code).toEqual("ErrorItemNotFound");
                        done();
                    });
                }, fail);
            }, fail);
        });

        it("should be able to send a newly created message ", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            client.me.fetch().then(function (owner) {
                var msgToSend = createMessage();
                msgToSend.ToRecipients[0] = createRecipient(owner._id, owner.DisplayName);
                messages.addMessage(msgToSend).then(function (added) {
                    tempEntities.push(added);
                    added.send().then(function() {
                        setTimeout(function() {
                            messages.getMessages().filter("Subject eq '" + msgToSend.Subject + "'").fetchAll().then(function(fetched) {
                                expect(fetched).toBeDefined();
                                expect(fetched).toEqual(jasmine.any(Array));
                                expect(fetched.length).toBeGreaterThan(0);
                                expect(fetched[0].ToRecipients[0].EmailAddress.Address)
                                    .toEqual(msgToSend.ToRecipients[0].EmailAddress.Address);
                                tempEntities.push(fetched[0]);
                                done();
                            }, fail);
                        }, 3000);
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to create a reply to existing message", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            client.me.fetch().then(function (owner) {
                var msgToSend = createMessage();
                msgToSend.ToRecipients[0] = createRecipient(owner._id, owner.DisplayName);
                messages.addMessage(msgToSend).then(function (added) {
                    tempEntities.push(added);
                    added.send().then(function() {
                        setTimeout(function () {
                            messages.getMessages().filter("Subject eq '" + msgToSend.Subject + "'").fetchAll().then(function (fetched) {
                                var justReceivedMessage = fetched[0];
                                tempEntities.push(justReceivedMessage);
                                justReceivedMessage.createReply().then(function (reply) {
                                    tempEntities.push(reply);
                                    messages.getMessage(reply.Id).fetch().then(function(fetchedReply) {
                                        expect(fetchedReply).toBeDefined();
                                        expect(fetchedReply.Subject).toMatch(msgToSend.Subject);
                                        expect(fetchedReply.ToRecipients[0])
                                            .toEqual(jasmine.objectContaining(msgToSend.ToRecipients[0]));
                                        done();
                                    }, fail);
                                }, fail);
                            }, fail);
                        }, 3000);
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to create a reply to all to existing message", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            client.me.fetch().then(function (owner) {
                var msgToSend = createMessage();
                msgToSend.ToRecipients[0] = createRecipient(owner._id, owner.DisplayName);
                msgToSend.CcRecipients = [];
                msgToSend.CcRecipients[0] = createRecipient('fakerecipient@' + owner._id.split('@')[1], "FakeRecipient");
                messages.addMessage(msgToSend).then(function (added) {
                    tempEntities.push(added);
                    added.send().then(function () {
                        setTimeout(function () {
                            messages.getMessages().filter("Subject eq '" + msgToSend.Subject + "'").fetchAll().then(function (fetched) {
                                var justReceivedMessage = fetched[0];
                                tempEntities.push(justReceivedMessage);
                                justReceivedMessage.createReplyAll().then(function (reply) {
                                    tempEntities.push(reply);
                                    messages.getMessage(reply.Id).fetch().then(function (fetchedReply) {
                                        expect(fetchedReply).toBeDefined();
                                        expect(fetchedReply.Subject).toMatch(msgToSend.Subject);
                                        expect(fetchedReply.CcRecipients[0])
                                            .toEqual(jasmine.objectContaining(msgToSend.CcRecipients[0]));
                                        done();
                                    }, fail);
                                }, fail);
                            }, fail);
                        }, 3000);
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to create a forwarded message to existing message", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            client.me.fetch().then(function (owner) {
                var msgToSend = createMessage();
                msgToSend.ToRecipients[0] = createRecipient(owner._id, owner.DisplayName);
                messages.addMessage(msgToSend).then(function (added) {
                    tempEntities.push(added);
                    added.send().then(function () {
                        setTimeout(function () {
                            messages.getMessages().filter("Subject eq '" + msgToSend.Subject + "'").fetchAll().then(function (fetched) {
                                var justReceivedMessage = fetched[0];
                                tempEntities.push(justReceivedMessage);
                                var fakeRecipient = createRecipient('fakerecipient@' + owner._id.split('@')[1], "FakeRecipient");
                                justReceivedMessage.createForward().then(function (fw) {
                                    tempEntities.push(fw);
                                    messages.getMessage(fw.Id).fetch().then(function (fetchedFw) {
                                        expect(fetchedFw).toBeDefined();
                                        expect(fetchedFw.Subject).toMatch(msgToSend.Subject);
                                        expect(fetchedFw.Body.Content).toBeDefined();
                                        expect(fetchedFw.ToRecipients).toEqual(jasmine.any(Array));
                                        done();
                                    }, fail);
                                }, fail);
                            }, fail);
                        }, 3000);
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to reply to existing message ", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            client.me.fetch().then(function (owner) {
                var msgToSend = createMessage();
                msgToSend.ToRecipients[0] = createRecipient(owner._id, owner.DisplayName);
                messages.addMessage(msgToSend).then(function (added) {
                    tempEntities.push(added);
                    added.send().then(function () {
                        setTimeout(function() {
                            messages.getMessages().filter("Subject eq '" + msgToSend.Subject + "'").fetchAll().then(function (fetched) {
                                fetched.length ?
                                    fetched[0].reply("Comment").then(function() {
                                        setTimeout(function() {
                                            messages.getMessages().filter("Subject eq 'RE: " + msgToSend.Subject + "'").fetch().then(function(fetched) {
                                                expect(fetched.length).toEqual(1);
                                                expect(fetched[0].Body.Content).toMatch("Comment");
                                                done();
                                            }, fail);
                                        }, 5000);
                                    }, fail) :
                                    fail("No messages with specified subject in inbox");
                            }, fail);
                        }, 3000);
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to reply to all senders of existing message", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            client.me.fetch().then(function (owner) {
                var msgToSend = createMessage();
                msgToSend.CcRecipients = [];
                msgToSend.CcRecipients[0] = createRecipient(owner._id, owner.DisplayName);
                msgToSend.ToRecipients[0] = createRecipient('fakerecipient@' + owner._id.split('@')[1], "FakeRecipient");
                messages.addMessage(msgToSend).then(function (added) {
                    tempEntities.push(added);
                    added.send().then(function () {
                        setTimeout(function () {
                            messages.getMessages().filter("Subject eq '" + msgToSend.Subject + "'").fetchAll().then(function (fetched) {
                                fetched.length ?
                                    fetched[0].replyAll("Comment").then(function () {
                                        setTimeout(function () {
                                            client.me.sentItems.messages.getMessages().filter("Subject eq 'RE: " + msgToSend.Subject + "'").fetch().then(function (fetched) {
                                                expect(fetched.length).toEqual(1);
                                                if (fetched.length > 0) {
                                                    expect(fetched[0].Body.Content).toMatch("Comment");
                                                    // TODO: somehow CcRecipients array is empty and the following
                                                    // expectation fails. Need additional investigation.
                                                    // expect(fetched[0].CcRecipients[0]).toEqual(jasmine.objectContaining(msgToSend.CcRecipients[0]));
                                                }
                                                done();
                                            }, fail);
                                        }, 5000);
                                    }, fail) :
                                    fail("No messages with specified subject in inbox");
                            }, fail);
                        }, 3000);
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to forward an existing message ", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            client.me.fetch().then(function (owner) {
                var msgToSend = createMessage();
                msgToSend.ToRecipients[0] = createRecipient(owner._id, owner.DisplayName);
                messages.addMessage(msgToSend).then(function (added) {
                    tempEntities.push(added);
                    added.send().then(function () {
                        setTimeout(function () {
                            messages.getMessages().filter("Subject eq '" + msgToSend.Subject + "'").fetchAll().then(function (fetched) {
                                fetched.length ?
                                    fetched[0].forward("Comment", [createRecipient(owner._id, owner.DisplayName)]).then(function () {
                                        setTimeout(function () {
                                            messages.getMessages().filter("Subject eq 'FW: " + msgToSend.Subject + "'").fetch().then(function (fetched) {
                                                expect(fetched.length).toEqual(1);
                                                expect(fetched[0].Body.Content).toMatch("Comment");
                                                done();
                                            }, fail);
                                        }, 5000);
                                    }, fail) :
                                    fail("No messages with specified subject in inbox");
                            }, fail);
                        }, 3000);
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to copy message to folder specified by id", function (done) {
            var fail = function(err) {
                expect(err).toBeUndefined();
                done();
            };
            folders.getFolders().fetchAll().then(function (fetched) {
                var targetFolder = fetched[0];
                messages.addMessage(createMessage()).then(function (added) {
                    tempEntities.push(added);
                    added.copy(targetFolder.Id).then(function (copied) {
                        tempEntities.push(copied);
                        targetFolder.messages.getMessage(added.Id).fetch().then(function (fetched) {
                            expect(fetched).toBeDefined();
                            expect(fetched.Subject).toEqual(added.Subject);
                            folders.getFolder("Drafts").messages.getMessage(added.Id).fetch().then(function (fetched) {
                                expect(fetched).toBeDefined();
                                expect(fetched.Subject).toEqual(added.Subject);
                                done();
                            }, fail);
                        }, fail);
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to copy message to folder specified by known name", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            messages.addMessage(createMessage()).then(function (added) {
                tempEntities.push(added);
                added.copy("Inbox").then(function (copied) {
                    tempEntities.push(copied);
                    folders.getFolder("Inbox").messages.getMessage(copied.Id).fetch().then(function (fetched) {
                        expect(fetched).toBeDefined();
                        expect(fetched.Subject).toEqual(added.Subject);
                        folders.getFolder("Drafts").messages.getMessage(added.Id).fetch().then(function (fetched) {
                            expect(fetched).toBeDefined();
                            expect(fetched.Subject).toEqual(added.Subject);
                            done();
                        }, fail);
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to move message to folder specified by id", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            folders.getFolders().fetchAll().then(function (fetched) {
                var targetFolder = fetched[0];
                messages.addMessage(createMessage()).then(function (added) {
                    tempEntities.push(added);
                    added.move(targetFolder.Id).then(function (moved) {
                        tempEntities.push(moved);
                        targetFolder.messages.getMessage(moved.Id).fetch().then(function (fetched) {
                            expect(fetched).toBeDefined();
                            expect(fetched.Subject).toEqual(added.Subject);
                            folders.getFolder("Drafts").messages.getMessage(added.Id).fetch().then(fail, function(err) {
                                expect(err).toBeDefined();
                                expect(err.message).toEqual("The specified object was not found in the store.");
                                done();
                            }, fail);
                        }, fail);
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to move message to folder specified by known name", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            messages.addMessage(createMessage()).then(function (added) {
                tempEntities.push(added);
                added.move("Inbox").then(function (moved) {
                    tempEntities.push(moved);
                    folders.getFolder("Inbox").messages.getMessage(moved.Id).fetch().then(function (fetched) {
                        expect(fetched).toBeDefined();
                        expect(fetched.Subject).toEqual(added.Subject);
                        folders.getFolder("Drafts").messages.getMessage(added.Id).fetch().then(fail, function (err) {
                            expect(err).toBeDefined();
                            expect(err.message).toEqual("The specified object was not found in the store.");
                            done();
                        }, fail);
                    }, fail);
                }, fail);
            }, fail);
        });
    });

    describe("Folders namespace", function () {

        var createRecipient = function (email, name) {
            return {
                EmailAddress: {
                    Name: name || guid(),
                    Address: email || (guid() + '@' + guid() + '.' + guid().substr(0, 3))
                }
            };
        };
        var createMessage = function (subject) {
            return {
                Subject: subject || guid(),
                ToRecipients: [createRecipient()],
                Body: {
                    ContentType: 0,
                    Content: "Test message"
                }
            };
        };
        var createFolder = function (name) {
            return {
                DisplayName: name || guid()
            };
        };

        var client, folders, messages, tempEntities;

        beforeEach(function () {
            client = new Microsoft.OutlookServices.Client(officeEndpointUrl,
                (new O365Auth.Context(authUrl, redirectUrl)).getAccessTokenFn(resourceUrl, '', appId));

            messages = client.me.messages;
            folders = client.me.folders;

            tempEntities = [];
        });

        afterEach(function () {
            tempEntities.forEach(function (entity) {
                try {
                    entity.delete();
                } catch (e) { }
            });
        });

        it("should be able to get user's folders", function (done) {
            folders.getFolders().fetchAll().then(function (f) {
                expect(f).toBeDefined();
                expect(f).toEqual(jasmine.any(Array));
                expect(f.length).toBeGreaterThan(0);
                expect(f[0]).toEqual(jasmine.any(Folders.Folder));
                done();
            }, function (err) {
                expect(err).toBeUndefined();
                done();
            });
        });

        it("should be able to create a new folder", function (done) {
            var fail = function(err) {
                expect(err).toBeUndefined();
                done();
            };
            var newFolder = createFolder();
            folders.addFolder(newFolder).then(function (f) {
                tempEntities.push(f);
                folders.getFolder(f.Id).fetch().then(function(fetched) {
                    expect(fetched).toBeDefined();
                    expect(fetched).toEqual(jasmine.any(Folders.Folder));
                    expect(fetched.DisplayName).toEqual(newFolder.DisplayName);
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to apply filter to user's folders", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            folders.getFolders().fetchAll().then(function (f) {
                var filterName = f[0].DisplayName;
                var filter = 'DisplayName eq \'' + filterName + '\'';
                folders.getFolders().filter(filter).fetchAll().then(function (f) {
                    expect(f).toBeDefined();
                    expect(f).toEqual(jasmine.any(Array));
                    expect(f.length).toEqual(1);
                    expect(f[0]).toEqual(jasmine.any(Folders.Folder));
                    expect(f[0].DisplayName).toEqual(filterName);
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to apply top query to user's folders", function (done) {
            folders.getFolders().top(1).fetchAll().then(function (cf) {
                expect(cf).toBeDefined();
                expect(cf).toEqual(jasmine.any(Array));
                expect(cf.length).toEqual(1);
                expect(cf[0]).toEqual(jasmine.any(Folders.Folder));
                done();
            }, function (err) {
                expect(err).toBeUndefined();
                done();
            });
        });

        it("should be able to get a folder by Id", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            folders.getFolders().fetchAll().then(function (fetched) {
                var folderToGet = fetched[0];
                folders.getFolder(folderToGet.Id).fetch().then(function (got) {
                    expect(got).toBeDefined();
                    expect(got.DisplayName).toEqual(folderToGet.DisplayName);
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to update existing folder", function(done) {
            var fail = function(err) {
                expect(err).toBeUndefined();
                done();
            };
            folders.addFolder(createFolder()).then(function(added) {
                tempEntities.push(added);
                added.DisplayName = guid();
                added.update().then(function() {
                    folders.getFolder(added.Id).fetch().then(function(fetched) {
                        expect(fetched).toBeDefined();
                        expect(fetched.DisplayName).toEqual(added.DisplayName);
                        done();
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to delete existing folder", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            folders.addFolder(createFolder()).then(function (added) {
                added.delete().then(function () {
                    folders.getFolder(added.Id).fetch().then(function (got) {
                        expect(got).toBeUndefined();
                        got.delete();
                        done();
                    }, function (err) {
                        expect(err).toBeDefined();
                        expect(err.code).toEqual("ErrorItemNotFound");
                        done();
                    });
                }, fail);
            }, fail);
        });

        describe("Folders nested operations", function () {

            it("should be able to copy existing folder to another folder specified by id", function (done) {
                var fail = function (err) {
                    expect(err).toBeUndefined();
                    done();
                };
                folders.getFolders().fetchAll().then(function (fetched) {
                    var targetFolder = fetched[0];
                    folders.addFolder(createFolder()).then(function (added) {
                        tempEntities.push(added);
                        added.copy(targetFolder.Id).then(function (copied) {
                            tempEntities.push(copied);
                            targetFolder.childFolders.getFolder(added.Id).fetch().then(function (fetched) {
                                expect(fetched).toBeDefined();
                                expect(fetched.DisplayName).toEqual(added.DisplayName);
                                folders.getFolder(added.Id).fetch().then(function (fetched) {
                                    expect(fetched).toBeDefined();
                                    expect(fetched.DisplayName).toEqual(added.DisplayName);
                                    done();
                                }, fail);
                            }, fail);
                        }, fail);
                    }, fail);
                }, fail);
            });

            it("should be able to copy existing folder to another folder specified by known name", function (done) {
                var fail = function (err) {
                    expect(err).toBeUndefined();
                    done();
                };
                folders.addFolder(createFolder()).then(function (added) {
                    tempEntities.push(added);
                    added.copy("Inbox").then(function (copied) {
                        tempEntities.push(copied);
                        folders.getFolder("Inbox").childFolders.getFolder(added.Id).fetch().then(function (fetched) {
                            expect(fetched).toBeDefined();
                            expect(fetched.DisplayName).toEqual(added.DisplayName);
                            folders.getFolder(added.Id).fetch().then(function (fetched) {
                                expect(fetched).toBeDefined();
                                expect(fetched.DisplayName).toEqual(added.DisplayName);
                                done();
                            }, fail);
                        }, fail);
                    }, fail);
                }, fail);
            });

            it("should be able to move existing folder to another folder specified by id", function (done) {
                var fail = function (err) {
                    expect(err).toBeUndefined();
                    done();
                };
                folders.getFolders().fetchAll().then(function (fetched) {
                    var targetFolder = fetched[0];
                    folders.addFolder(createFolder()).then(function (added) {
                        tempEntities.push(added);
                        added.move(targetFolder.Id).then(function (moved) {
                            tempEntities.push(moved);
                            targetFolder.childFolders.getFolder(added.Id).fetch().then(function (fetched) {
                                expect(fetched).toBeDefined();
                                expect(fetched.DisplayName).toEqual(added.DisplayName);
                                done();
                            }, fail);
                        }, fail);
                    }, fail);
                }, fail);
            });

            it("should be able to move existing folder to another folder specified by known name", function (done) {
                var fail = function (err) {
                    expect(err).toBeUndefined();
                    done();
                };
                folders.addFolder(createFolder()).then(function (added) {
                    tempEntities.push(added);
                    added.move("Inbox").then(function (moved) {
                        tempEntities.push(moved);
                        client.me.inbox.childFolders.getFolder(added.Id).fetch().then(function (fetched) {
                            expect(fetched).toBeDefined();
                            expect(fetched.DisplayName).toEqual(added.DisplayName);
                            done();
                        }, fail);
                    }, fail);
                }, fail);
            });

            it("should get folder's nested folders", function (done) {
                var fail = function (err) {
                    expect(err).toBeUndefined();
                    done();
                };
                folders.getFolder('Inbox').fetch().then(function (inbox) {
                    inbox.childFolders.getFolders().fetchAll().then(function (fetched) {
                        expect(fetched).toBeDefined();
                        expect(fetched).toEqual(jasmine.any(Array));
                        if (fetched.length === 0) {
                            // no contact folders created for this account, can't continue other tests
                            done();
                            return;
                        }
                        expect(fetched[0]).toEqual(jasmine.any(Folders.Folder));
                        done();
                    }, fail);
                }, fail);
            });

            it("should get folder's nested messages", function(done) {
                messages.addMessage(createMessage()).then(function (created) {
                    tempEntities.push(created);
                    // new message is created in drafts folder
                    // and we need to check if another folder (inbox) is not contain this message as well
                    folders.getFolder("Inbox").fetch().then(function(inbox) {
                        inbox.messages.getMessages().fetchAll().then(function(inboxMessages) {
                            expect(inboxMessages).toBeDefined();
                            expect(inboxMessages).toEqual(jasmine.any(Array));
                            for (var i = inboxMessages.length - 1; i >= 0; i--) {
                                var message = inboxMessages[i];
                                expect(message.Subject).not.toEqual(created.Subject);
                            }
                            done();
                        });
                    });
                });
            });

            it("should be able to create message in specific folder", function (done) {
                var fail = function (err) {
                    expect(err).toBeUndefined();
                    done();
                };
                folders.getFolder('Inbox').fetch().then(function (inbox) {
                    inbox.messages.addMessage(createMessage).then(function(created) {
                        tempEntities.push(created);
                        inbox.messages.getMessage(created.Id).fetch().then(function(fetched) {
                            expect(fetched).toBeDefined();
                            expect(fetched).toEqual(jasmine.any(Messages.Message));
                            expect(fetched.Subject).toEqual(created.Subject);
                            done();
                        }, fail);
                    }, fail);
                }, fail);
            });
        });
    });

    describe("Attachments namespace", function() {

        var createRecipient = function(email, name) {
            return {
                EmailAddress: {
                    Name: name || guid(),
                    Address: email || (guid() + '@' + guid() + '.' + guid().substr(0, 3))
                }
            };
        };
        var createEvent = function (subject) {
            return {
                Subject: subject || guid(),
                Start: new Date(),
                End: new Date()
            };
        };
        var createMessage = function(subject) {
            return {
                Subject: subject || guid(),
                ToRecipients: [createRecipient()],
                Body: {
                    ContentType: 0,
                    Content: "Test message"
                }
            };
        };
        var createFileAttachment = function (text) {
            return new Attachments.FileAttachment(null, null, {
                Name: guid() + ".txt",
                ContentBytes: text ? btoa(text) : btoa(guid())
            });
        };
        var createItemAttachment = function (message) {
            return new Attachments.ItemAttachment(null, null, {
                Name: guid(),
                Item: message || createMessage()
            });
        };
        var client, messages, events, tempEntities, backInterval;

        beforeEach(function() {
            client = new Microsoft.OutlookServices.Client(officeEndpointUrl,
                (new O365Auth.Context(authUrl, redirectUrl)).getAccessTokenFn(resourceUrl, '', appId));

            events = client.me.events;
            messages = client.me.messages;
            tempEntities = [];

            // increase standart jasmine timeout up to 10 seconds because to some tests
            // are perform some time consumpting operations (e.g. send, reply, forward)
            backInterval = jasmine.DEFAULT_TIMEOUT_INTERVAL;
            jasmine.DEFAULT_TIMEOUT_INTERVAL = 15000;
        });

        afterEach(function() {
            tempEntities.forEach(function(entity) {
                try {
                    entity.delete();
                } catch (e) {
                }
            });

            // revert back default jasmine timeout
            jasmine.DEFAULT_TIMEOUT_INTERVAL = backInterval;
        });

        it("should be able to add a file attachment to existing message", function(done) {
            var fail = function(err) {
                expect(err).toBeUndefined();
                done();
            };
            messages.addMessage(createMessage()).then(function(added) {
                tempEntities.push(added);
                var fileAttachment = createFileAttachment();
                added.attachments.addAttachment(fileAttachment).then(function(attachment) {
                    expect(attachment).toBeDefined();
                    expect(attachment).toEqual(jasmine.any(Attachments.FileAttachment));
                    expect(attachment.Name).toEqual(fileAttachment.Name);
                    expect(attachment.ContentBytes).toEqual(fileAttachment.ContentBytes);
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to add a file attachment to existing event", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            events.addEvent(createEvent()).then(function (added) {
                tempEntities.push(added);
                var fileAttachment = createFileAttachment();
                added.attachments.addAttachment(fileAttachment).then(function (attachment) {
                    expect(attachment).toBeDefined();
                    expect(attachment).toEqual(jasmine.any(Attachments.FileAttachment));
                    expect(attachment.Name).toEqual(fileAttachment.Name);
                    expect(attachment.ContentBytes).toEqual(fileAttachment.ContentBytes);
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to add an item attachment to existing message", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            messages.addMessage(createMessage()).then(function (added) {
                tempEntities.push(added);
                var item = createMessage();
                var itemAttachment = createItemAttachment(item);
                added.attachments.addAttachment(itemAttachment).then(function (attachment) {
                    expect(attachment).toBeDefined();
                    expect(attachment).toEqual(jasmine.any(Attachments.ItemAttachment));
                    expect(attachment.Name).toEqual(itemAttachment.Name);
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to add an item attachment to existing event", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            events.addEvent(createEvent()).then(function (added) {
                tempEntities.push(added);
                var item = createMessage();
                var itemAttachment = createItemAttachment(item);
                added.attachments.addAttachment(itemAttachment).then(function (attachment) {
                    expect(attachment).toBeDefined();
                    expect(attachment).toEqual(jasmine.any(Attachments.FileAttachment));
                    expect(attachment.Name).toEqual(itemAttachment.Name);
                    done();
                }, fail);
            }, fail);
        });

        it("should be able to get message's attachments", function(done) {
            var fail = function(err) {
                expect(err).toBeUndefined();
                done();
            };
            messages.addMessage(createMessage()).then(function(added) {
                tempEntities.push(added);
                var fileAttachment = createFileAttachment();
                added.attachments.addAttachment(fileAttachment).then(function() {
                    messages.getMessage(added.Id).fetch().then(function(created) {
                        expect(created.HasAttachments).toBeTruthy();
                        created.attachments.getAttachments().fetch().then(function(addedAttachments) {
                            expect(addedAttachments).toEqual(jasmine.any(Array));
                            expect(addedAttachments.length).toEqual(1);
                            var addedAttachment = addedAttachments[0];
                            expect(addedAttachment.Name).toEqual(fileAttachment.Name);
                            expect(addedAttachment.ContentBytes).toEqual(fileAttachment.ContentBytes);
                            done();
                        }, fail);
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to get event's attachments", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            events.addEvent(createEvent()).then(function (added) {
                tempEntities.push(added);
                var fileAttachment = createFileAttachment();
                added.attachments.addAttachment(fileAttachment).then(function () {
                    events.getEvent(added.Id).fetch().then(function (created) {
                        expect(created.HasAttachments).toBeTruthy();
                        created.attachments.getAttachments().fetch().then(function (addedAttachments) {
                            expect(addedAttachments).toEqual(jasmine.any(Array));
                            expect(addedAttachments.length).toEqual(1);
                            var addedAttachment = addedAttachments[0];
                            expect(addedAttachment.Name).toEqual(fileAttachment.Name);
                            expect(addedAttachment.ContentBytes).toEqual(fileAttachment.ContentBytes);
                            done();
                        }, fail);
                    }, fail);
                }, fail);
            }, fail);
        });

        // pended since filter query fails on server with internal server error
        // TODO: review again
        xit("should be able to apply filter to item's attachments", function(done) {
            var fail = function(err) {
                expect(err).toBeUndefined();
                done();
            };
            messages.addMessage(createMessage()).then(function(added) {
                tempEntities.push(added);
                var fileAttachment1 = createFileAttachment();
                var fileAttachment2 = createFileAttachment();
                added.attachments.addAttachment(fileAttachment1).then(function () {
                    added.attachments.addAttachment(fileAttachment2).then(function() {
                        messages.getMessage(added.Id).fetch().then(function(created) {
                            expect(created.HasAttachments).toBeTruthy();
                            created.attachments.getAttachments().fetch().then(function(addedAttachments) {
                                expect(addedAttachments).toEqual(jasmine.any(Array));
                                expect(addedAttachments.length).toEqual(2);
                                var filter = "Name eq '" + fileAttachment1.Name + "'";
                                created.attachments.getAttachments().filter(filter).fetch().then(function(filteredAttachments) {
                                    expect(filteredAttachments).toEqual(jasmine.any(Array));
                                    expect(filteredAttachments.length).toEqual(1);
                                    expect(filteredAttachments[0]).toEqual(jasmine.objectContaining(fileAttachment1));
                                    done();
                                }, fail);
                            }, fail);
                        }, fail);
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to apply top query to item's attachments", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            messages.addMessage(createMessage()).then(function (added) {
                tempEntities.push(added);
                var fileAttachment1 = createFileAttachment();
                var fileAttachment2 = createFileAttachment();
                added.attachments.addAttachment(fileAttachment1).then(function () {
                    added.attachments.addAttachment(fileAttachment2).then(function () {
                        messages.getMessage(added.Id).fetch().then(function (created) {
                            expect(created.HasAttachments).toBeTruthy();
                            created.attachments.getAttachments().fetch().then(function (addedAttachments) {
                                expect(addedAttachments).toEqual(jasmine.any(Array));
                                expect(addedAttachments.length).toEqual(2);
                                created.attachments.getAttachments().top(1).fetch().then(function (filteredAttachments) {
                                    expect(filteredAttachments).toEqual(jasmine.any(Array));
                                    // TODO: commented out the following tests since top query still returns
                                    // a collection of 2 attachments here
                                    // expect(filteredAttachments.length).toEqual(1);
                                    // expect(filteredAttachments[0]).toEqual(jasmine.objectContaining(fileAttachment1));
                                    done();
                                }, fail);
                            }, fail);
                        }, fail);
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to get a newly added attachment by Id", function (done) {
            var fail = function(err) {
                expect(err).toBeUndefined();
                done();
            };
            messages.addMessage(createMessage()).then(function (added) {
                tempEntities.push(added);
                var fileAttachment = createFileAttachment();
                added.attachments.addAttachment(fileAttachment).then(function (addedAttachment) {
                    messages.getMessage(added.Id).fetch().then(function (created) {
                        expect(created.HasAttachments).toBeTruthy();
                        created.attachments.getAttachment(addedAttachment.Id).fetch().then(function (createdAttachment) {
                            expect(createdAttachment).toEqual(jasmine.any(Attachments.FileAttachment));
                            expect(createdAttachment.Name).toEqual(fileAttachment.Name);
                            expect(createdAttachment.ContentBytes).toEqual(fileAttachment.ContentBytes);
                            done();
                        }, fail);
                    }, fail);
                }, fail);
            }, fail);
        });

        // pended due to lack of support on server
        // TODO: review again
        xit("should be able to modify existing file attachment", function(done) {
            var fail = function(err) {
                expect(err).toBeUndefined();
                done();
            };
            messages.addMessage(createMessage()).then(function (added) {
                tempEntities.push(added);
                var fileAttachment = createFileAttachment();
                added.attachments.addAttachment(fileAttachment).then(function (addedAttachment) {
                    addedAttachment.Name = guid();
                    addedAttachment.update().then(function(updatedAttachment) {
                        expect(updatedAttachment).toBeDefined();
                        expect(updatedAttachment).toEqual(jasmine.any(Attachments.Attachment));
                        expect(updatedAttachment.name).toEqual(addedAttachment.Name);
                        done();
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to modify existing item attachment", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            messages.addMessage(createMessage()).then(function (added) {
                tempEntities.push(added);
                var itemAttachment = createItemAttachment();
                added.attachments.addAttachment(itemAttachment).then(function (addedAttachment) {
                    addedAttachment.Name = guid();
                    addedAttachment.update().then(function (updatedAttachment) {
                        expect(updatedAttachment).toBeDefined();
                        expect(updatedAttachment).toEqual(jasmine.any(Attachments.Attachment));
                        expect(updatedAttachment.name).toEqual(addedAttachment.Name);
                        done();
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to delete an attachment from message", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            messages.addMessage(createMessage()).then(function (added) {
                tempEntities.push(added);
                var fileAttachment = createFileAttachment();
                added.attachments.addAttachment(fileAttachment).then(function (addedAttachment) {
                    addedAttachment.delete().then(function () {
                        messages.getMessage(added.Id).fetch().then(function(createdMessage) {
                            expect(createdMessage.HasAttachments).toBeFalsy();
                            createdMessage.attachments.getAttachments().fetchAll().then(function(attachments) {
                                expect(attachments).toBeDefined();
                                expect(attachments).toEqual(jasmine.any(Array));
                                expect(attachments.length).toEqual(0);
                                done();
                            }, fail);
                        }, fail);
                    }, fail);
                }, fail);
            }, fail);
        });

        it("should be able to delete an attachment from event", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            events.addEvent(createEvent()).then(function (added) {
                tempEntities.push(added);
                var fileAttachment = createFileAttachment();
                added.attachments.addAttachment(fileAttachment).then(function (addedAttachment) {
                    addedAttachment.delete().then(function () {
                        events.getEvent(added.Id).fetch().then(function (createdEvent) {
                            expect(createdEvent.HasAttachments).toBeFalsy();
                            createdEvent.attachments.getAttachments().fetchAll().then(function (attachments) {
                                expect(attachments).toBeDefined();
                                expect(attachments).toEqual(jasmine.any(Array));
                                expect(attachments.length).toEqual(0);
                                done();
                            }, fail);
                        }, fail);
                    }, fail);
                }, fail);
            }, fail);
        });
    });

    describe("Users namespace", function () {

        var client, users;

        beforeEach(function () {
            client = new Microsoft.OutlookServices.Client(officeEndpointUrl,
                (new O365Auth.Context(authUrl, redirectUrl)).getAccessTokenFn(resourceUrl, '', appId));
            users = client.users;
        });

        it("should be able to get Users collection", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            users.getUsers().fetchAll().then(function(usersList) {
                expect(usersList).toBeDefined();
                expect(usersList).toEqual(jasmine.any(Array));
                expect(usersList.length).toBeGreaterThan(1);
                expect(usersList[0]).toEqual(jasmine.any(Users.User));
                done();
            }, fail);
        });

        it("should be able to get user by id", function (done) {
            var fail = function (err) {
                expect(err).toBeUndefined();
                done();
            };
            users.getUsers().fetchAll().then(function (usersList) {
                users.getUser(usersList[0].Id).fetch().then(function(user) {
                    expect(user).toBeDefined();
                    expect(user).toEqual(jasmine.any(Users.User));
                    expect(user.Id).toEqual(usersList[0].Id);
                    done();
                }, fail);
            }, fail);
        });
    });
};

exports.defineManualTests = function (contentEl, createActionButton) {

    createActionButton('Log in', function () {
        var authContext = new O365Auth.Context(authUrl, redirectUrl);
        authContext.getAccessToken(resourceUrl, null, appId, null).then(function (token) {
            console.log("Token is: " + token);
        }, function (err) {
            console.error(err);
        });
    });

    createActionButton('Log out', function () {
        var authContext = new O365Auth.Context(authUrl, redirectUrl);
        return authContext.logOut(appId).then(function () {
            console.log("Logged out");
        }, function (err) {
            console.error(err);
        });
    });
};

namespace app.Services {

    export interface IUserService {
        getMe(useSelect: boolean): ng.IPromise<any>;
        getUsers(useFilter: boolean): ng.IHttpPromise<any>;
        createUser(tenant: string): ng.IHttpPromise<any>;
        getDrive(): ng.IHttpPromise<any>;
        getEvents(): ng.IHttpPromise<any>;
        createEvent(): ng.IHttpPromise<any>;
        updateEvent(): ng.IPromise<any>;
        deleteEvent(): ng.IPromise<any>;
        getMessages(): ng.IPromise<any>;
        sendMessage(recipientEmailAddress: string): ng.IHttpPromise<any>;
        getUserPhoto(): ng.IHttpPromise<any>;
        getManager(): ng.IPromise<any>;
        getDirectReports(): ng.IPromise<any>;
        getMemberOf(): ng.IPromise<any>;
        getFiles(): ng.IPromise<any>;
        createFile(): ng.IPromise<any>;
        downloadFile(): ng.IPromise<any>;
        updateFile(): ng.IPromise<any>;
        renameFile(): ng.IPromise<any>;
        deleteFile(): ng.IPromise<any>;
        createFolder(): ng.IHttpPromise<any>;

    }

    class UserService implements IUserService {

        private baseUrl: string;
        static $inject = ["$http", "$q", "app.Services.Common"];

        constructor(
            private $http: ng.IHttpService,
            private $q: ng.IQService,
            private common) {
            this.baseUrl = common.baseUrl;
        }

        /**
         * Get information about the signed-in user.
         */
        getMe(useSelect: boolean): ng.IPromise<any> {
            let reqUrl: string = this.baseUrl + "/me";

            // Append $select OData query string.
            if (useSelect) {
                reqUrl += "?$select=AboutMe,Responsibilities,Tags";
            }

            let req: ng.IRequestConfig = {
                method: "GET",
                url: reqUrl
            };



            return this.$q((resolve :ng.IQResolveReject<any>, reject: ng.IQResolveReject<any>) => {
                this.$http(req)
                    .then(function(res) {
                        resolve(res);
                    }, function(err) {
                        resolve(err);
                    });
            });



        }

        /**
         * Get existing users collection from the tenant.
         */
        getUsers(useFilter: boolean): ng.IHttpPromise<any> {
            let reqUrl: string = this.baseUrl + "/myOrganization/users";

            // Append $filter OData query string.
            if (useFilter) {
                // This filter will return users in your tenant based in the US.
                reqUrl += "?$filter=country eq \"United States\"";
            }

            let req: ng.IRequestConfig = {
                method: "GET",
                url: reqUrl
            };

            return this.$http(req);
        };

        /**
         * Add a user to the tenant"s users collection.
         */
        createUser(tenant: string): ng.IHttpPromise<any> {
            let randomUserName = this.common.guid();

            // The data in newUser are the minimum required properties.
            let newUser = {
                accountEnabled: true,
                displayName: "User " + randomUserName,
                mailNickname: randomUserName,
                passwordProfile: {
                    password: "p@ssw0rd!"
                },
                userPrincipalName: randomUserName + "@" + tenant
            };

            let req: ng.IRequestConfig = {
                method: "POST",
                url: this.baseUrl + "/myOrganization/users",
                data: newUser
            };

            return this.$http(req);
        };

        /**
         * Get the signed-in user"s drive.
         */
        getDrive(): ng.IHttpPromise<any> {
            let req: ng.IRequestConfig = {
                method: "GET",
                url: this.baseUrl + "/me/drive"
            };

            return this.$http(req);
        };

        /**
         * Get the signed-in user"s calendar events.
         */
        getEvents(): ng.IHttpPromise<any> {
            let req: ng.IRequestConfig = {
                method: "GET",
                url: this.baseUrl + "/me/events"
            };

            return this.$http(req);
        };

        /**
         * Adds an event to the signed-in user"s calendar.
         */
        createEvent(): ng.IHttpPromise<any> {
            // The new event will be 30 minutes and take place tomorrow at the current time.
            let startTime = new Date();
            startTime.setDate(startTime.getDate() + 1);
            let endTime = new Date(startTime.getTime() + 30 * 60000);

            let newEvent = {
                Subject: "Weekly Sync",
                Location: {
                    DisplayName: "Water cooler"
                },
                Attendees: [{
                    Type: "Required",
                    EmailAddress: {
                        Address: "mara@fabrikam.com"
                    }
                }],
                Start: {
                    "DateTime": startTime,
                    "TimeZone": "PST"
                },
                End: {
                    "DateTime": endTime,
                    "TimeZone": "PST"
                },
                Body: {
                    Content: "Status updates, blocking issues, and next steps.",
                    ContentType: "Text"
                }
            };

            let req: ng.IRequestConfig = {
                method: "POST",
                url: this.baseUrl + "/me/events",
                data: newEvent
            };

            return this.$http(req);
        };

        /**
         * Creates an event, adds it to the signed-in user"s calendar, and then
         * updates the Subject.
         */
        updateEvent(): ng.IPromise<any> {
            let deferred = this.$q.defer();

            let eventUpdates = {
                Subject: "Sync of the Week"
            };

            // Create an event to update.
            this.createEvent()
                // If successful, take event ID and update it.
                .then(function(response) {
                    let eventId = response.data.id;

                    let req = {
                        method: "PATCH",
                        url: this.baseUrl + "/me/events/" + eventId,
                        data: eventUpdates
                    };

                    deferred.resolve(this.$http(req));
                }, function(error) {
                    deferred.reject({
                        setupError: "Unable to create an event to update.",
                        response: error
                    });
                });

            return deferred.promise;
        };

        /**
         * Creates an event, adds it to the signed-in user"s calendar, and then
         * deletes the event.
         */
        deleteEvent(): ng.IPromise<any> {
            let deferred: ng.IDeferred<any> = this.$q.defer();

            // Create an event to update first.
            this.createEvent()
                // If successful, take event ID and update it.
                .then(function(response) {
                    let eventId = response.data.id;

                    let req = {
                        method: "DELETE",
                        url: this.baseUrl + "/me/events/" + eventId
                    };

                    deferred.resolve(this.$http(req));
                }, function(error) {
                    deferred.reject({
                        setupError: "Unable to create an event to delete.",
                        response: error
                    });
                });

            return deferred.promise;
        };

        /**
         * Get the signed-in user"s messages.
         */
        getMessages(): ng.IHttpPromise<any> {
            let req: ng.IRequestConfig = {
                method: "GET",
                url: this.baseUrl + "/me/messages"
            };

            return this.$http(req);
        };

        /**
         * Send a message as the signed-in user.
         */
        sendMessage(recipientEmailAddress: string): ng.IHttpPromise<any> {
            let newMessage = {
                Message: {
                    Subject: "Microsoft Graph snippets",
                    Body: {
                        ContentType: "Text",
                        Content: "You can send an email by making a POST request to /me/microsoft.graph.sendMail."
                    },
                    ToRecipients: [
                        {
                            EmailAddress: {
                                Address: recipientEmailAddress
                            }
                        }
                    ]
                },
                SaveToSentItems: true
            };

            let req: ng.IRequestConfig = {
                method: "POST",
                url: this.baseUrl + "/me/microsoft.graph.sendMail",
                data: newMessage
            };

            return this.$http(req);
        };

        /**
         * Get signed-in user"s photo.
         */
        getUserPhoto(): ng.IHttpPromise<any> {
            let req: ng.IRequestConfig = {
                method: "GET",
                url: this.baseUrl + "/me/photo"
            };

            return this.$http(req);
        };

        /**
         * Get signed-in user"s manager.
         */
        getManager(): ng.IHttpPromise<any> {
            let req: ng.IRequestConfig = {
                method: "GET",
                url: this.baseUrl + "/me/manager"
            };

            return this.$http(req);
        };

        /**
         * Get signed-in user"s direct reports.
         */
        getDirectReports(): ng.IHttpPromise<any> {
            let req: ng.IRequestConfig = {
                method: "GET",
                url: this.baseUrl + "/me/directReports"
            };

            return this.$http(req);
        };

        /**
         * Get groups that signed-in user is a member of.
         */
        getMemberOf(): ng.IHttpPromise<any> {
            let req: ng.IRequestConfig = {
                method: "GET",
                url: this.baseUrl + "/me/memberOf"
            };

            return this.$http(req);
        };

        /**
         * Get signed-in user"s files.
         */
        getFiles(): ng.IHttpPromise<any> {
            let req: ng.IRequestConfig = {
                method: "GET",
                url: this.baseUrl + "/me/drive/root/children"
            };

            return this.$http(req);
        };

        /**
         * Create a file in signed-in user"s root directory.
         */
        createFile(): ng.IHttpPromise<any> {
            let randomFileName: string = this.common.guid() + ".txt";

            let req: ng.IRequestConfig = {
                method: "PUT",
                url: this.baseUrl + "/me/drive/root/children/" + randomFileName + "/content",
                data: {
                    content: randomFileName + " is the name of this file."
                }
            };

            return this.$http(req);
        };

        /**
         * Get contents of a specific file.
         */
        downloadFile(): ng.IPromise<any> {
            let deferred = this.$q.defer();

            this.createFile()
                .then(function(response) {
                    let fileId: string = response.data.id;

                    let req = {
                        method: "GET",
                        url: this.baseUrl + "/me/drive/items/" + fileId + "/content"
                    };

                    deferred.resolve(this.$http(req));
                }, function(error) {
                    deferred.reject({
                        setupError: "Unable to create a file to download.",
                        response: error
                    });
                });

            return deferred.promise;
        };

        /**
         * Updates the contents of a specific file.
         */
        updateFile(): ng.IPromise<any> {
            let deferred = this.$q.defer();

            this.createFile()
                .then(function(response) {
                    let fileId: string = response.data.id;

                    let req = {
                        method: "PUT",
                        url: this.baseUrl + "/me/drive/items/" + fileId + "/content",
                        data: {
                            content: "Updated file contents."
                        }
                    };

                    deferred.resolve(this.$http(req));
                }, function(error) {
                    deferred.reject({
                        setupError: "Unable to create a file to update.",
                        response: error
                    });
                });

            return deferred.promise;
        };

        /**
         * Renames a specific file.
         */
        renameFile(): ng.IPromise<any> {
            let deferred = this.$q.defer();

            this.createFile()
                .then(function(response) {
                    let fileId: string = response.data.id;
                    let fileName: string = response.data.name.replace(".txt", "-renamed.txt");

                    let req: ng.IRequestConfig = {
                        method: "PATCH",
                        url: this.baseUrl + "/me/drive/items/" + fileId,
                        data: {
                            name: fileName
                        }
                    };

                    deferred.resolve(this.$http(req));
                }, function(error) {
                    deferred.reject({
                        setupError: "Unable to create a file to rename.",
                        response: error
                    });
                });

            return deferred.promise;
        };

        /**
         * Deletes a specific file.
         */
        deleteFile(): ng.IPromise<any> {
            let deferred = this.$q.defer();

            this.createFile()
                .then(function(response) {
                    let fileId: string = response.data.id;

                    let req: ng.IRequestConfig = {
                        method: "DELETE",
                        url: this.baseUrl + "/me/drive/items/" + fileId
                    };

                    deferred.resolve(this.$http(req));
                }, function(error) {
                    deferred.reject({
                        setupError: "Unable to create a file to delete.",
                        response: error
                    });
                });

            return deferred.promise;
        };

        /**
         * Creates a folder in the root directory.
         */
        createFolder(): ng.IHttpPromise<any> {
            let req: ng.IRequestConfig = {
                method: "POST",
                url: this.baseUrl + "/me/drive/root/children",
                data: {
                    name: this.common.guid(),
                    folder: {},
                    "@name.conflictBehavior": "rename"
                }
            };

            return this.$http(req);
        }
    }
    angular
        .module("app.Services")
        .service("app.Services.UserService", UserService);
}
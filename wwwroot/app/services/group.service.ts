namespace app.Services {
    'use strict';

    export interface IGroupService {
        getGroups(): ng.IHttpPromise<any>;
        createGroup(): ng.IHttpPromise<any>;
		getGroup(): ng.IPromise<any>;
		updateGroup(): ng.IPromise<any>;
		deleteGroup(): ng.IPromise<any>;
		getMembers(): ng.IPromise<any>;
		getOwners(): ng.IPromise<any>;
    }

    class GroupService implements IGroupService {
        private baseUrl: string;

        static $inject = ["$http", "$q", "app.Services.Common"];

        constructor(
            private $http: ng.IHttpService,
            private $q: ng.IQService,
            private common: app.Services.ICommonService) {

            this.baseUrl = common.baseUrl;
        }

		/**
		 * Gets the list of groups that exist in the tenant.
		 */
        public getGroups(): ng.IHttpPromise<any> {
            var req = {
                method: 'GET',
                url: this.baseUrl + '/myOrganization/groups'
            };

            return this.$http(req);
        }

		/**
		 * Creates a new security group in the tenant.
		 */
        public createGroup(): ng.IHttpPromise<any> {
            var uuid = this.common.newGuid();

            var newGroup = {
                displayName: uuid,
                mailEnabled: false, // Set to true for mail-enabled groups.
                mailNickname: uuid,
                securityEnabled: true // Set to true for security-enabled groups. Do not set this property if creating an Office 365 group.
            };

            var req = {
                method: 'POST',
                url: this.baseUrl + '/myOrganization/groups',
                data: newGroup
            };

            return this.$http(req);
        };

		/**
		 * Gets information about a specific group.
		 */
        getGroup(): ng.IPromise<any> {
            return this.getGroupsSetup('');
        };

		/**
		 * Updates the description of a specific group.
		 */
        public updateGroup(): ng.IPromise<any> {
            var deferred = this.$q.defer();

            // You can only update groups created via the Microsoft Graph API, so to make sure we have one,
            // we'll create it here and then update its description.
            this.createGroup()
                .then(function(response) {
                    var groupId = response.data.id;
                    this.$log.debug('Group "' + groupId + '" was created. Updating the group\'s description...');

                    var groupUpdates = {
                        description: 'This is a group.'
                    };

                    var req = {
                        method: 'PATCH',
                        url: this.baseUrl + '/myOrganization/groups/' + groupId,
                        data: groupUpdates
                    };

                    deferred.resolve(this.$http(req));
                }, function(error) {
                    deferred.reject({
                        setupError: 'Unable to create a new group to update.',
                        response: error
                    });
                });

            return deferred.promise;
        };

		/**
		 * Deletes a specific group.
		 */
        public deleteGroup(): ng.IPromise<any> {
            var deferred = this.$q.defer();

            // You can only delete groups created via the Microsoft Graph API, so to make sure we have one,
            // we'll create it here and then delete its description.
            this.createGroup()
                .then(function(response) {
                    var groupId = response.data.id;
                    this.$log.debug('Group "' + groupId + '" was created. Deleting the group...');

                    var req = {
                        method: 'DELETE',
                        url: this.baseUrl + '/myOrganization/groups/' + groupId
                    };

                    deferred.resolve(this.$http(req));
                }, function(error) {
                    deferred.reject({
                        setupError: 'Unable to create a new group to delete.',
                        response: error
                    });
                });

            return deferred.promise;
        };

		/**
		 * Gets members of a specific group.
		 */
        getMembers(): ng.IPromise<any> {
            return this.getGroupsSetup('members');
        };

		/**
		 * Gets owners of a specific group.
		 */
        public getOwners(): ng.IPromise<any> {
            return this.getGroupsSetup('owners');
        };

		/**
		 * Several snippets require a group ID to work. This method does the setup work
		 * required to get the ID of a group in the tenant and then makes a request to the
		 * desired navigation property.
		 */
        private getGroupsSetup(endpoint): ng.IPromise<any> {
            var deferred = this.$q.defer();

            this.getGroups()
                .then(function(response) {
                    // Check to make sure at least 1 group is returned.
                    if (response.data.value.length >= 1) {
                        var groupId = response.data.value[0].id;

                        var req = {
                            method: 'GET',
                            url: this.baseUrl + '/myOrganization/groups/' + groupId + '/' + endpoint
                        };

                        deferred.resolve(this.$http(req));
                    }
                    else {
                        deferred.reject({
                            setupError: 'Tenant doesn\'t have any groups.',
                            response: response
                        });
                    }
                }, function(error) {
                    deferred.reject({
                        setupError: 'Unable to get list of tenant\'s groups.',
                        response: error
                    });
                });

            return deferred.promise;
        };

    };
    angular
        .module('app.Services')
        .service('app.Services.GroupService', GroupService);

}
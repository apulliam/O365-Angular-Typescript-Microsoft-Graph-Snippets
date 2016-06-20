namespace app.Services {

    export interface IDriveService {
        getDrives(): ng.IHttpPromise<any>;
    }


    class DriveService {
        static $inject = ["$http", "app.Services.Common"];
        private baseUrl: string;

        constructor(
            private $http: ng.IHttpService,
            private common: ICommonService) {

            this.baseUrl = common.baseUrl;
        }
        /**
         * Gets the list of drives that exist in the tenant.
         */
        getDrives(): ng.IHttpPromise<any> {

            let req: ng.IRequestConfig = {
                method: "GET",
                url: this.baseUrl + "/myOrganization/drives"
            };

            return this.$http(req);
        }
    }

    angular
        .module("app.Services")
        .service("app.Services.DriveService", DriveService);
}

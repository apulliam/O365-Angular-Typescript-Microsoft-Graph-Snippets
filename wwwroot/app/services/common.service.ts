namespace app.Services {

    export interface ICommonService {
        baseUrl: string;
        newGuid(): string;
    };

    class CommonService implements ICommonService {

        public baseUrl: string;

        /**
		 * Constructor for the snippet object.
		 */
        constructor() {
            // Properties
            this.baseUrl = "https://graph.microsoft.com/v1.0";
        }

		/**
		 * Random GUID generator. Copied from Stack Overflow user "broofa".
		 * http://stackoverflow.com/users/109538/broofa
		 * http://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript
		 */
        public newGuid(): string {
            function s4() {
                return Math.floor((1 + Math.random()) * 0x10000)
                    .toString(16)
                    .substring(1);
            }
            return s4() + s4() + "-" + s4() + "-" + s4() + "-" +
                s4() + "-" + s4() + s4() + s4();
        };
    }

    angular
        .module("app.Services")
        .service("app.Services.Common", CommonService);

}


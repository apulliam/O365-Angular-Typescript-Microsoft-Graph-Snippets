namespace app.Main {

    interface ISnippetGroup {
        groupTitle: string;
        snippets: Snippet[];

    }

    interface ISnippet {
        title: string;
        description: string;
        documentationUrl: string;
        apiUrl: string;
        requireAdmin: boolean;
        run(run: () => ng.IPromise<any> | ng.IHttpPromise<any>): void;
        request: any;
        response: any;
        setupError: any;
    }

    /////////////////////////
    // Snippet             //
    // -------             //
    // Title               //
    // Description         //
    // Documenation URL    //
    // API URL             //
    // Require admin?      //
    // Snippet code        //
    /////////////////////////



    class Snippet implements ISnippet {

        public request: any;
        public response: any;
        public setupError: any;

        constructor(
            public title: string,
            public description: string,
            public documentationUrl: string,
            public apiUrl: string,
            public requireAdmin: boolean,
            public run: (run: (useSelect?: boolean) => ng.IPromise<any> | ng.IHttpPromise<any>) => void) {
        }


    }

    interface IMainControllerScope {
        setActive(title: string): string | void;
        activeSnippet: ISnippet;
        laddaLoading: boolean;
    }


    /**
     * The MainController code.
     */
    class MainController implements IMainControllerScope {


        ////////////////////////////////////////////////
        // All of the snippets that fall under the    //
        // "drives" tenant-level resource collection. //
        ////////////////////////////////////////////////
        private driveSnippets: ISnippetGroup = {
            groupTitle: "drives",
            snippets: [

                /////////////////////////////////
                //       DRIVES SNIPPETS       //
                /////////////////////////////////
                new Snippet(
                    "GET myOrganization/drives",
                    "Gets  all of the drives in your tenant.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/drive_get",
                    this.common.baseUrl + "/myOrganization/drives",
                    false,
                    () => this.doSnippet(() => this.drives.getDrives())
                )
            ]
        };

        ////////////////////////////////////////////////
        // All of the snippets that fall under the    //
        // "groups" tenant-level resource collection. //
        ////////////////////////////////////////////////
        private groupSnippets: ISnippetGroup = {
            groupTitle: "groups",

            snippets: [
                /////////////////////////////////
                //       GROUPS SNIPPETS       //
                /////////////////////////////////
                new Snippet(
                    "GET myOrganization/groups",
                    "Gets all of the groups in your tenant\"s directory.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/group_list",
                    this.common.baseUrl + "/myOrganization/groups",
                    false,
                    () => this.doSnippet(() => this.groups.getGroups())),
                new Snippet(
                    "POST myOrganization/groups",
                    "Adds a new security group to the tenant.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/group_post_groups",
                    this.common.baseUrl + "/myOrganization/groups",
                    false,
                    () => this.doSnippet(() => this.groups.createGroup())),
                new Snippet(
                    "GET myOrganization/groups/{Group.id}",
                    "Gets information about a group in the tenant by ID.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/group_get",
                    this.common.baseUrl + "/myOrganization/groups/{Group.id}",
                    false,
                    () => this.doSnippet(() => this.groups.getGroup())),
                new Snippet(
                    "PATCH myOrganization/groups/{Group.id}",
                    "Adds a new group to the tenant, then updates the description of that group.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/group_update",
                    this.common.baseUrl + "/myOrganization/groups/{Group.id}",
                    false,
                    () => this.doSnippet(() => this.groups.updateGroup())),
                new Snippet(
                    "DELETE myOrganization/groups/{Group.id}",
                    "Adds a new group to the tenant, then deletes the group.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/group_delete",
                    this.common.baseUrl + "/myOrganization/groups/{Group.id}",
                    false,
                    () => this.doSnippet(() => this.groups.deleteGroup())),
                new Snippet(
                    "GET myOrganization/groups/{Group.id}/members",
                    "Gets the members of a group.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/group_list_members",
                    this.common.baseUrl + "/myOrganization/groups/{Group.id}/members",
                    false,
                    () => this.doSnippet(() => this.groups.getMembers())),
                new Snippet(
                    "GET myOrganization/groups/{Group.id}/owners",
                    "Gets the owners of a group.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/group_list_owners",
                    this.common.baseUrl + "/myOrganization/groups/{Group.id}/owners",
                    false,
                    () => this.doSnippet(() => this.groups.getOwners()))
            ]
        };

        ////////////////////////////////////////////////
        // All of the snippets that fall under the    //
        // "users" tenant-level resource collection.  //
        ////////////////////////////////////////////////
        public userSnippets: ISnippetGroup = {
            groupTitle: "users",
            snippets: [
                ///////////////////////////////
                //       USER SNIPPETS       //
                ///////////////////////////////
                new Snippet(
                    "GET myOrganization/users",
                    "Gets all of the users in your tenant\"s directory.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/user_list",
                    this.common.baseUrl + "/myOrganization/users",
                    false,
                    () => this.doSnippet(() => this.users.getUsers(false))),
                new Snippet(
                    "GET myOrganization/users?$filter=country eq \"United States\"",
                    "Gets all of the users in your tenant\"s directory who are from the United States, using $filter.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/user_list",
                    this.common.baseUrl + "/myOrganization/users?$filter=country eq \"United States\"",
                    false,
                    () => this.doSnippet(() => this.users.getUsers(true))),
                new Snippet(
                    "POST myOrganization/users",
                    "Adds a new user to the tenant\"s directory.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/user_post_users",
                    this.common.baseUrl + "/myOrganization/users",
                    true,
                    () => this.doSnippet(() => this.users.createUser(this.tenant))),
                new Snippet(
                    "GET me",
                    "Gets information about the signed-in user.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/user_get",
                    this.common.baseUrl + "/me",
                    false,
                    () => this.doSnippet(() => this.users.getMe(false))),
                new Snippet(
                    "GET me?$select=AboutMe,Responsibilities,Tags",
                    "Gets select information about the signed-in user, using $select.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/user_get",
                    this.common.baseUrl + "/me?$select=AboutMe,Responsibilities,Tags",
                    false,
                    () => this.doSnippet(() => this.users.getMe(true))),
                //////////////////////////////////////
                //        USER/DRIVE SNIPPETS       //
                //////////////////////////////////////
                new Snippet(
                    "GET me/drive",
                    "Gets the signed-in user\"s drive.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/drive_get",
                    this.common.baseUrl + "/me/drive",
                    false,
                    () => this.doSnippet(() => this.users.getDrive())),
                ///////////////////////////////////////
                //        USER/EVENTS SNIPPETS       //
                ///////////////////////////////////////
                new Snippet(
                    "GET me/events",
                    "Gets the signed-in user\"s calendar events.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/user_list_events",
                    this.common.baseUrl + "/me/events",
                    false,
                    () => this.doSnippet(() => this.users.getEvents())),
                new Snippet(
                    "POST me/events",
                    "Adds an event to the signed-in user\"s calendar.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/user_post_events",
                    this.common.baseUrl + "/me/events",
                    false,
                    () => this.doSnippet(() => this.users.createEvent())),
                new Snippet(
                    "PATCH me/events/{Event.id}",
                    "Adds an event to the signed-in user\"s calendar, then updates the subject of the event.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/event_update",
                    this.common.baseUrl + "/me/events/{Event.id}",
                    false,
                    () => this.doSnippet(() => this.users.updateEvent())),
                new Snippet(
                    "DELETE me/events/{Event.id}",
                    "Adds an event to the signed-in user\"s calendar, then deletes the event.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/event_delete",
                    this.common.baseUrl + "/me/events/{Event.id}",
                    false,
                    () => this.doSnippet(() => this.users.deleteEvent())),
                //////////////////////////////////////////
                //        USER/MESSAGES SNIPPETS        //
                //////////////////////////////////////////
                new Snippet(
                    "GET me/messages",
                    "Gets the signed-in user\"s emails.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/user_list_messages",
                    this.common.baseUrl + "/me/messages",
                    false,
                    () => this.doSnippet(() => this.users.getMessages())),
                new Snippet(
                    "POST me/microsoft.graph.sendMail",
                    "Sends an email as the signed-in user and saves a copy to their Sent Items folder.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/user_sendmail",
                    this.common.baseUrl + "/me/microsoft.graph.sendMail",
                    false,
                    () => this.doSnippet(() => this.users.sendMessage(this.adalAuthenticationService.userInfo.userName))),
                //////////////////////////////////////////
                //          USER/FILES SNIPPETS         //
                //////////////////////////////////////////
                new Snippet(
                    "GET me/drive/root/children",
                    "Gets files from the signed-in user\"s root directory.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/item_list_children",
                    this.common.baseUrl + "/me/drive/root/children",
                    false,
                    () => this.doSnippet(() => this.users.getFiles())),
                new Snippet(
                    "PUT me/drive/root/children/{FileName}/content",
                    "Creates a file with content in the signed-in user\"s root directory.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/item_uploadcontent",
                    this.common.baseUrl + "/me/drive/root/children/{FileName}/content",
                    false,
                    () => this.doSnippet(() => this.users.createFile())),
                new Snippet(
                    "GET me/drive/items/{File.id}/content",
                    "Downloads a file.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/item_downloadcontent",
                    this.common.baseUrl + "/me/drive/items/{File.id}/content",
                    false,
                    () => this.doSnippet(() => this.users.downloadFile())),
                new Snippet(
                    "PUT me/drive/items/{File.id}/content",
                    "Updates the contents of a file.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/item_downloadcontent",
                    this.common.baseUrl + "/me/drive/items/{File.id}/content",
                    false,
                    () => this.doSnippet(() => this.users.updateFile())),
                new Snippet(
                    "PATCH me/drive/items/{File.id}",
                    "Renames a file.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/item_update",
                    this.common.baseUrl + "/me/drive/items/{File.id}",
                    false,
                    () => this.doSnippet(() => this.users.renameFile())),
                new Snippet(
                    "DELETE me/drive/items/{File.id}",
                    "Deletes a file.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/item_delete",
                    this.common.baseUrl + "/me/drive/items/{File.id}",
                    false,
                    () => this.doSnippet(() => this.users.deleteFile())),
                new Snippet(
                    "POST me/drive/root/children",
                    "Creates a folder in the signed-in user\"s root directory.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/item_post_children",
                    this.common.baseUrl + "/me/drive/root/children",
                    false,
                    () => this.doSnippet(() => this.users.createFolder())),
                ///////////////////////////////////////////////
                //        MISCELLANEOUS USER SNIPPETS        //
                ///////////////////////////////////////////////
                new Snippet(
                    "GET me/manager",
                    "Gets the signed-in user\"s manager.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/user_list_manager",
                    this.common.baseUrl + "/me/manager",
                    false,
                    () => this.doSnippet(() => this.users.getManager())),
                new Snippet(
                    "GET me/directReports",
                    "Gets the signed-in user\"s direct reports.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/user_list_directreports",
                    this.common.baseUrl + "/me/directReports",
                    false,
                    () => this.doSnippet(() => this.users.getDirectReports())),
                new Snippet(
                    "GET me/photo",
                    "Gets the signed-in user\"s photo.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/profilephoto_get",
                    this.common.baseUrl + "/me/photo",
                    false,
                    () => this.doSnippet(() => this.users.getUserPhoto())),
                new Snippet(
                    "GET me/memberOf",
                    "Gets the groups that the signed-in user is a member of.",
                    "http://graph.microsoft.io/docs/api-reference/v1.0/api/user_list_memberof",
                    this.common.baseUrl + "/me/memberOf",
                    false,
                    () => this.doSnippet(() => this.users.getMemberOf())
                )
            ]
        };

        private snippetGroups: ISnippetGroup[] = [
            this.userSnippets,
            this.groupSnippets,
            this.driveSnippets
        ];


        // Properties
        public activeSnippet: ISnippet;

        public laddaLoading: boolean;

        // Methods
        static $inject = [
            "$q",
            "adalAuthenticationService",
            "app.Services.Common",
            "app.Services.UserService",
            "app.Services.GroupService",
            "app.Services.DriveService"
        ];

        constructor(
            private $q: ng.IQService,
            private adalAuthenticationService,
            private common: app.Services.ICommonService,
            private users: app.Services.IUserService,
            private groups: app.Services.IGroupService,
            private drives: app.Services.IDriveService) {

            this.activate();
            this.main = this;
        }

        private main: MainController;  // save this for promise callbacks

        private tenant: string;

        /**
         * This function does any initialization work the
         * controller needs.
         */
        private activate(): void {
            if (this.adalAuthenticationService.userInfo.isAuthenticated) {
                this.activeSnippet = this.snippetGroups[0].snippets[0];

                this.tenant = this.adalAuthenticationService.userInfo.userName.split("@")[1];
            }
        }

        /**
         * Takes in a snippet, starts animation, executes snippet and handles response,
         * then stops animation.
         */
        private doSnippet(run: () => ng.IPromise<any> | ng.IHttpPromise<any>, useSelect?: boolean): void {
            // Starts button animation.
            this.laddaLoading = true;

            // Clear old data.
            // this.activeSnippet.request = null;
            // this.activeSnippet.response = null;
            // this.activeSnippet.setupError = null;

            run()
                .then((response: any): void => {

                    // Format request and response.
                    let request = response.config;
                    response = {
                        status: response.status,
                        statusText: response.statusText,
                        data: response.data
                    };

                    // Attach response to view model.
                    this.activeSnippet.request = request;
                    this.activeSnippet.response = response;
                }, (error: any): void => {

                    // If a snippet requires setup (i.e. creating an event to update, creating a file
                    // to delete, etc.) and it fails, handle that error message differently.
                    if (error.setupError) {
                        // Extract setup error message.
                        this.activeSnippet.setupError = error.setupError;

                        // Extract response data.
                        this.activeSnippet.response = {
                            status: error.response.data,
                            statusText: error.response.statusText,
                            data: error.response.data
                        };

                        return;
                    }

                    // Format request and response.
                    let request = error.config;
                    error = {
                        status: error.status,
                        statusText: error.statusText,
                        data: error.data
                    };

                    // Attach response to view model.
                    this.activeSnippet.request = request;
                    this.activeSnippet.response = error;
                })
                .finally(():void => {
                     let main = this;
                    // Stops button animation.
                    main.laddaLoading = false;
                });
        }

        /**
         * Sets class of list item in the sidebar.
         */
        setActive(title: string): string | void {
            if (!this.adalAuthenticationService.userInfo.isAuthenticated) {
                return;
            }

            if (title === this.activeSnippet.title) {
                return "active";
            }
            else {
                return "";
            }
        }
    }

    angular
        .module("app")
        .controller("app.MainController", MainController);

}

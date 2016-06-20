namespace app.Main {

    interface INavbarScope {
        isCollapsed: boolean;
        isopen: boolean;
        activate(): void;
        connect(): void;
        disconnect(): void;
    }
    /**
     * The NavbarController code.
     */

     class NavbarController implements INavbarScope {
        static $inject = ["adalAuthenticationService", "$log"];

        constructor(
            private adalAuthenticationService,
            private $log: ng.ILogService) {
                this.activate();
        }

        // Properties
        isCollapsed: boolean;
        isopen: boolean = false;

        // Methods
        /**
         * This function does any initialization work the
         * controller needs.
         */
        activate(): void {
            this.isCollapsed = true;
        };

        /**
         * Expose the login method to the view.
         */
        connect(): void {
            this.$log.debug("Connecting to Office 365...");
            this.adalAuthenticationService.login();
        };

        /**
         * Expose the logOut method to the view.
         */
        disconnect(): void {
            this.$log.debug("Disconnecting from Office 365...");
            this.adalAuthenticationService.logOut();
        };

        /**
         * Event listener for dropdown menu in navbar.
         */
        toggleDropdown($event: ng.IAngularEvent) {
            $event.preventDefault();
            $event.stopPropagation();
            this.isopen = !this.isopen;
        };
    }

      angular
        .module("app")
        .controller("app.NavbarController", NavbarController);
}

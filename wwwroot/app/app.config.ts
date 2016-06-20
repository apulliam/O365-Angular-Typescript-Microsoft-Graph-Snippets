((): void => {


    config.$inject = ["$routeProvider", "adalAuthenticationServiceProvider", "$httpProvider", "cfpLoadingBarProvider", "localStorageServiceProvider"];

    function config($routeProvider: ng.route.IRouteProvider,
                    adalAuthenticationServiceProvider,
                    $httpProvider: ng.IHttpProvider,
                    cfpLoadingBarProvider: ng.loadingBar.ILoadingBarProvider,
                    localStorageServiceProvider: ng.local.storage.ILocalStorageServiceProvider) {
        // Configure the routes.
        $routeProvider
            .when("/", {
                templateUrl: "app/main/main.html",
                controller: "app.MainController",
                controllerAs: "vm"
            })
            .otherwise({
                redirectTo: "/"
            });

        // Allow cross domain requests to be made.
        // ToDo: AGP->Is this supported anymore?  $httpProvider.defaults.useXDomain = true;
        delete $httpProvider.defaults.headers.common["X-Requested-With"];

        // Configure ADAL JS.
        adalAuthenticationServiceProvider.init(
            {
                clientId: clientId,
                endpoints: {
                    "https://graph.microsoft.com": "https://graph.microsoft.com"
                }
            },
            $httpProvider
        );

        // Remove spinner from loading bar.
        cfpLoadingBarProvider.includeSpinner = false;

        // Local storage configuration.
        localStorageServiceProvider.setPrefix("tsMsGraph");
    }

     angular
        .module("app").
        config(config);
})();
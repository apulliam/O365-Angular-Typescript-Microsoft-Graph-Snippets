((): void => {
    "use strict";

    console.log("Registering main app module");
    angular
        .module("app", [
            "app.Services",
            "ngRoute",
            "ui.bootstrap",
            "AdalAngular",
            "angular-loading-bar",
            "ladda",
            "LocalStorageModule"
        ]);
})();


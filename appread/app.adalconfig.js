(function () {
  'use strict';

  var officeAddin = angular.module('officeAddin').config(['$httpProvider', 'adalAuthenticationServiceProvider', 'appId', function($httpProvider, adalProvider, appId){
      
        var adalConfig = {
          tenant: 'common',
          clientId: appId,
          endpoints: {
            "https://graph.microsoft.com" :{
                scope:["mail.readWrite mail.send"] 
            } 
          }, 
          scope:["mail.readWrite mail.send"]
       }
       
    }])
})();
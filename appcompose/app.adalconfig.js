(function () {
  'use strict';

  var officeAddin = angular.module('officeAddin').config(['$httpProvider', 'adalAuthenticationServiceProvider', 'appId', function($httpProvider, adalProvider, appId){
      
        var adalConfig = {
          tenant: 'common',
          /*clientId: '8a7b8264-7be1-4cbc-9178-5941c91c4c86',*/
          clientId: appId,
        /*  extraQueryParameter: 'nux=1',*/
          endpoints: {
            "https://graph.microsoft.com" :{
                scope:["mail.readWrite mail.send"] 
            } 
          }, 
          scope:["mail.readWrite mail.send"]
       }
          // cacheLocation: 'localStorage', // enable this for IE, as sessionStorage does not work for localhost. 
       
       /* adalProvider.init(adalConfig, $httpProvider);*/
    }])
})();
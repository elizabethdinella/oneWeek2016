(function(){
  'use strict';

  angular.module('officeAddin')
         .controller('homeController', ['$scope', '$http', function($scope, $http) {
                
           $scope.getMessages = function(){
               return $http.get("https://graph.microsoft.com/beta/me/messages");
           }
  
         }]);
});

  /**
   * Controller constructor
   */
/*
  function homeController(dataService){
    var vm = this;  // jshint ignore:line
    vm.title = 'home controller';
    vm.dataObject = {};
    vm.messages= "messages";

    getDataFromService();
   /* vm.messages = getMessages();
    console.log(vm.messages);*/
      
/*
    function getDataFromService(){
      dataService.getData()
        .then(function(response){
          vm.dataObject = response;
        });
    }
   
   function getMessages(){
       return $http.get("https://graph.microsoft.com/beta/me/messages");
   }
     
  }

})();*/


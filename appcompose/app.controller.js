  angular.module('officeAddin')
         .controller('homeController', ['$scope', '$http', function($scope, $http) {
             
             
           /*$scope.messages = "testMessages";  */
           
           $scope.getMessages = function(){
               return $http.get("https://graph.microsoft.com/beta/me/messages");
           }
           
           $scope.getMessages().success(function(results, status, headers){
                  $scope.messages = results;                   
           }).error(function(err, status){
                   console.log("ERROR");      
           });
             
           console.log($scope.messages);
                
 }]);
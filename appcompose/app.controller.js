  angular.module('officeAddin')
         .controller('homeController', ['$scope', '$http', function($scope, $http) {
            
        
         $scope.getSentiment = function(){
            $scope.getMessageBody();
            var url = 'https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/sentiment';
            var text = "messageBody"
             
            if($scope.messageBody){
               text = $scope.messageBody;  
            }
             
            var postData = {
                  "documents": [
                    {
                      "language": "en",
                      "id": "string",
                      "text": text
                    }
                  ]
            };
             
            return $http({
                url: url,
                method: 'POST',
                headers: {
                    'Content-type': 'application/json',
                    'Ocp-Apim-Subscription-Key':  '1c510246edaf4ca48f5bd1ab7766e771'
                },
                
                data: postData
            });
         }
         
         $scope.handleSentimentResult = function() {
             $scope.getSentiment().success(function(results, status){
                  if(!$scope.messageBody){
                      $scope.sentimentScore = 0.5;
                      return;
                  }
                  $scope.sentimentScore = results.documents[0].score;
                  console.log($scope.sentimentScore);
             }).error(function(err, status){
                  console.log(err); 
             });
         }
         
         $scope.getMessageBody = function(){
             Office.context.mailbox.item.body.getAsync(
                  "text",
                  { asyncContext:"This is passed to the callback" },
                      function callback(result) {
                            $scope.messageBody = result.value; 
                 });
         }
         
         
         
        $scope.bingSearch = function(){
            
            var breedIndex = Math.floor((Math.random() * (dogBreeds.length - 1)) + 1);  
            var breed = dogBreeds[breedIndex];
            var url = 'https://api.datamarket.azure.com/Bing/Search/v1/Image?Query=%27dog%27' + breed + "%27cute%27";

            return $http({
                url: url,
                method: 'GET',
                headers: {
                    'Authorization': 'Basic ' + 'OjYyRTNnVzRvQXBLRk93dEtVU0h0SXJhaTFOZ2twNElVZEJ3ckNmMmp0TlU='
                }
            });
        }
       
        $scope.handleSearchResults = function(){
            $scope.bingSearch().success(function(results, status){
                console.log("search successful!");
                $scope.resultObj = results;
            }).error(function(err, status){
                console.log("error searching");
                console.log(err);
            });
            
        }
        
        function callback(result) {
          if (result.error) {
            showMessage(result.error);
          } else {
            showMessage("Attachment added");
          }
        }

        $scope.addAttachment = function(source){
          var options = { 'asyncContext': { var1: 1, var2: 2 } };
          var attachmentURL = source;
          Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
        }
        
 }]);
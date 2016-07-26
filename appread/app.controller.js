  angular.module('officeAddin')
         .controller('homeController', ['$scope', '$http', function($scope, $http) {
            
         $scope.sentimentScore = .1;
         $scope.resultsObj = [];
             
        $scope.bingSearch = function(searchTerm){
  
            var url = 'https://api.datamarket.azure.com/Bing/Search/v1/Image?Query=%27dog%27' + searchTerm + "%27";

            return $http({
                url: url,
                method: 'GET',
                headers: {
                    'Authorization': 'Basic ' + 'OjYyRTNnVzRvQXBLRk93dEtVU0h0SXJhaTFOZ2twNElVZEJ3ckNmMmp0TlU='
                }
            });
        }
       
        $scope.handleSearchResults = function(searchTerm){
            $scope.bingSearch(searchTerm).success(function(results, status){
            
                console.log("success!");
                console.log(results);
                $scope.resultsObj.push(results.d.results[0].MediaUrl);
                
            }).error(function(err, status){
                
                console.log("error");
                console.log(err);
                
            });
            
        }
             
         $scope.getSentimentScore = function(){
             return $scope.sentimentScore;
         }
         
         
         $scope.getKeyWords = function(){
            var url = 'https://westus.api.cognitive.microsoft.com/text/analytics/v2.0/keyPhrases';
            var text = " "
             
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
         
         $scope.handleKeyWordsResult = function() {
             $scope.keyPhrases = [];
             $scope.getKeyWords().success(function(results, status){
                  $scope.keyPhrases = results.documents[0].keyPhrases;
                  console.log("key phrases");
                  console.log($scope.keyPhrases);
                  for(var i=0; i<$scope.keyPhrases.length; i++){
                     $scope.handleSearchResults($scope.keyPhrases[i]);
                 }
             }).error(function(err, status){
                  console.log(err);
             });
         }
        
        
         $scope.getSentiment = function(){
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
                  var from = Office.context.mailbox.item.from;
                  var prefix = "It looks like " + from.displayName + " ";
                  var badMessage = prefix + "isn't in a good mood, send them a dog to make them feel better!"
                  var goodMessage = prefix + "is already happy, but send them a dog to make them even happier!"
                  var displayMessage = "";
                  if($scope.sentimentScore >= .5){
                      displayMessage = goodMessage;
                  }else{
                      displayMessage = badMessage;
                  }
                  document.getElementById('message').innerText += displayMessage;
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
                            $scope.handleSentimentResult();
                 });
         }
         
         $scope.constructReply = function(attatchURL){
             Office.context.mailbox.item.displayReplyForm(
            {
              'htmlBody' : 'hi',
              'attachments' :
              [
                {
                  'type' : Office.MailboxEnums.AttachmentType.File,
                  'name' : 'dog',
                  'url' :  attatchURL
                }
              ]
            });
         }
         

             
        var item;

        Office.initialize = function () {
            item = Office.context.mailbox.item;
            // Checks for the DOM to load using the jQuery ready function.
            $(document).ready(function () {
                // After the DOM is loaded, app-specific code can run.
                // Get all the recipients of the composed item.
                $scope.getMessageBody();
                //getAllRecipients();
                
                console.log($scope.sentimentScore);
            });
        }
             
        
 }]);
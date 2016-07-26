  angular.module('officeAddin')
         .controller('homeController', ['$scope', '$http', function($scope, $http) {
            
             $scope.sentimentScore = .1;
             
             
         $scope.getSentimentScore = function(){
             return $scope.sentimentScore;
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
                   /* 'Access-Control-Allow-Origin':  'https://westus.api.cognitive.microsoft.com',*/
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
         
         $scope.constructReply = function(){
             Office.context.mailbox.item.displayReplyForm(
            {
              'htmlBody' : 'hi',
              'attachments' :
              [
                {
                  'type' : Office.MailboxEnums.AttachmentType.File,
                  'name' : 'dog',
                  'url' : 'http://i.imgur.com/sRgTlGR.jpg'
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
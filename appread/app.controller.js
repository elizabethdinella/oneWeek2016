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
                  var badMessage = "It looks like ___ isn't in a good mood, send them a dog to make them feel better!"
                  var goodMessage = "It looks like ___  is already happy, but send them a dog to make them even happier!"
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

/*        // Get the email addresses of all the recipients of the composed item.
        function getAllRecipients() {
            // Local objects to point to recipients of either
            // the appointment or message that is being composed.
            // bccRecipients applies to only messages, not appointments.
            var toRecipients, ccRecipients, bccRecipients;
            // Verify if the composed item is an appointment or message.
            if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
                toRecipients = item.requiredAttendees;
                ccRecipients = item.optionalAttendees;
            }
            else {
                toRecipients = item.to;
                ccRecipients = item.cc;
                bccRecipients = item.bcc;
            }*/

            // Use asynchronous method getAsync to get each type of recipients
            // of the composed item. Each time, this example passes an anonymous 
            // callback function that doesn't take any parameters.
            /*toRecipients.getAsync(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Async call to get to-recipients of the item completed.
                    // Display the email addresses of the to-recipients. 
                    write ('To-recipients of the item:');
                    displayAddresses(asyncResult);
                    $scope.getTime(asyncResult.value[0].emailAddress);
                }    
            }); // End getAsync for to-recipients.

            // Get any cc-recipients.
            ccRecipients.getAsync(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Async call to get cc-recipients of the item completed.
                    // Display the email addresses of the cc-recipients.
                    write ('Cc-recipients of the item:');
                    displayAddresses(asyncResult);
                }
            }); // End getAsync for cc-recipients.

            // If the item has the bcc field, i.e., item is message,
            // get any bcc-recipients.
            if (bccRecipients) {
                bccRecipients.getAsync(function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Async call to get bcc-recipients of the item completed.
                    // Display the email addresses of the bcc-recipients.
                    write ('Bcc-recipients of the item:');
                    displayAddresses(asyncResult);
                }

                }); // End getAsync for bcc-recipients.
             }
        }*/

        // Recipients are in an array of EmailAddressDetails
        // objects passed in asyncResult.value.
        function displayAddresses (asyncResult) {
            for (var i=0; i<asyncResult.value.length; i++)
                write (asyncResult.value[i].emailAddress);
        }

        // Writes to a div with id='message' on the page.
        function write(message){
            document.getElementById('message').innerText += message; 
        }
             
        
 }]);
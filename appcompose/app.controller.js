  angular.module('officeAddin')
         .controller('homeController', ['$scope', '$http', function($scope, $http) {
             
             
        $scope.bingSearch = function(){
  
            var url = 'https://api.datamarket.azure.com/Bing/Search/v1/Image?Query=%27dog%27';

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
            
                console.log("success!");
                console.log(results);
                $scope.resultObj = results;
                
                
            }).error(function(err, status){
                
                console.log("error");
                console.log(err);
                
            });
            
        }
        
        $scope.getTime = function(){
              $http.get("https://graph.microsoft.com/v1.0/me/messages");         
        }
             
        function callback(result) {
          if (result.error) {
            showMessage(result.error);
          } else {
            showMessage("Attachment added");
          }
        }

        $scope.addAttachment = function(source){
          // The values in asyncContext can be accessed in the callback
          var options = { 'asyncContext': { var1: 1, var2: 2 } };

          var attachmentURL = source;
          Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
        }
        
        
        var item;

        Office.initialize = function () {
            item = Office.context.mailbox.item;
            document.getElementById('message').innerText += "INIT"; 
            // Checks for the DOM to load using the jQuery ready function.
            $(document).ready(function () {
                // After the DOM is loaded, app-specific code can run.
                // Get all the recipients of the composed item.
                getAllRecipients();
            });
        }

        // Get the email addresses of all the recipients of the composed item.
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
            }

            // Use asynchronous method getAsync to get each type of recipients
            // of the composed item. Each time, this example passes an anonymous 
            // callback function that doesn't take any parameters.
            toRecipients.getAsync(function (asyncResult) {
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
        }

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
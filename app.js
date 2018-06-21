// Import Modules (Packages)
var restify = require('restify');
var builder = require('botbuilder');
var http = require('http');
var https = require('https');
var Readable = require('stream').Readable;
var Writable = require('stream').Writable;
var fs = require('fs');
const delay = require('delay');
var async = require("async");
var azure = require('botbuilder-azure'); 
var azureStore = require('azure-storage');

// Initialize Global Variables
//Azure table storage connection details
var tableName = 'ContractDetailsTable'; // Table Name in Azure Table Storage
var storageName = 'myazurebotstoragetable';// storageName - Obtained from Azure Portal
// Key to the Azure storage - Obtained from Azure Portal
var storageKey = 'A3wvHL8VcC4fYKnNj6bvLjmvdGKQUZGtdf9UKOPI+kY3kUA9kEy8kobCUPEk3z3RvJLET+IwUQBEGGVua3z12Q=='; 

// Create an inMemory storage object for the bot
var inMemoryStorage = new builder.MemoryBotStorage();

// Objects created to connect to Azure Storage Table
var azureTableClient = new azure.AzureTableClient(tableName, storageName, storageKey);
var tableStorage = new azure.AzureBotStorage({gzipData: false}, azureTableClient);
var tableSvc = azureStore.createTableService(storageName, storageKey);

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
   console.log('%s listening to %s', server.name, server.url); 
});

// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword
});

// Listen for chat messages from users 
server.post('/api/messages', connector.listen());

// Create the bot and receive messages from the user and respond with Info fetched
// This is a bot that fetches the contract details of the user.
var bot = new builder.UniversalBot(connector, [
    function (session) {  
       var chatInput = ''
       chatInput = session.message.text 
       session.send("Hi.Welcome to the Help Desk!");
       builder.Prompts.text(session, "How may i help you!");      
    },function (session, results) {
        var userRequest = results.response;
        session.send("Processing your request..Please Hang On!");
        // REST API call to LUIS ai for Language Understanding, to know the user's intent
        var luisAPIHostPath = '/luis/v2.0/apps/b730d38a-ad3d-427c-a93c-6d9e609dc83c?subscription-key=d7c93c26d2a2442f85b87b74589b9b10&verbose=true&timezoneOffset=0&q='
        var uriPath = ''
        
        uriPath = encodeURI(luisAPIHostPath.concat(userRequest));
        
        var extServerOptions = {
          host: 'westus.api.cognitive.microsoft.com',
          path: uriPath,
          method: 'GET'
        };
  
        https.request(extServerOptions, function (res) {
          res.setEncoding('utf8');
          var body = '';
          var luisResp = '';
  
          res.on('data', function(chunk){
              body += chunk;
          }); 
  
          res.on('end', function () {          
              try {

                if(body != ''){
                   luisResp = JSON.parse(body); 

                   var intentAPIResp = JSON.parse(JSON.stringify(luisResp));
                   var intentsArray = [];
                   intentsArray = intentAPIResp.intents
                   var intentScore = 0
  
                   for(var i=0;i<intentsArray.length;i++){
                          if(intentsArray[i].intent == 'ContractDetails'){
                              intentScore =  intentsArray[i].score 
                          } 
                   }

                   console.log('intentScore intentScore intentScore %s',intentScore); 
                    if(intentScore > 0.96) {            
                       session.send("I sense that you need to know your Contract/Party related details!");
                       builder.Prompts.text(session, "Please enter the Contract Number!");         
                    } else {
                      session.send("At the momment we can only provide your contract details! Thank You! Bye.");
                      session.endDialog();
                    }
  
                } else{
                    session.send("Unable to process your request at the momment. Please try after sometime! Bye.");
                    session.endDialog();
                }             
              } catch (error) {
                  if (error instanceof SyntaxError) {
                      console.log(`SyntaxError instance -  Error Name is ${error.name}: Error Name is ${error.message}`);
                  } else {
                      console.log(`Error Name is ${error.name}: Error Name is ${error.message}`);
                  }
              }      
              //Write the consumed API data to a file 
            //   var writableStream = fs.createWriteStream('APIWrite.txt');         
            //   writableStream.write(JSON.stringify(luisResp));         
          }); 
      }).end();
    },
    function (session, results) {
        session.dialogData.contractNumber = results.response;  
        var rowKeyStr = ''
        rowKeyStr = session.dialogData.contractNumber
        session.dialogData.contractNumStr = rowKeyStr

        if(isNaN(rowKeyStr)){
            builder.Prompts.text(session, "Please enter a valid number as contract number and try again!"); 
        } else{
             //Azure Storage database call to fetch contract details
        tableSvc.retrieveEntity(tableName, 'Party',rowKeyStr, function(error, result, response){       
                              
        if(!error){
        // result contains the entity. Assign to variables
           var resName = result.Name._
           var resDelin = result.DelinquencyStatus._
           var resInvDueAmt =  result.InvoiceDueAmount._
           var resInvDueDate =  result.InvoiceDueDate._
        //    var myJSONObj = { "name":resName,  "delin":resDelin, "invDueAmt":resInvDueAmt, "invDueDate": resInvDueDate};
        //    var myJSONStr =  JSON.stringify(myJSONObj);
    
           // Display the contract details to user in chat.
            session.send(`Welcome : ${resName} !!<br/>         Contract/Party Details - <br/>Contract Number: ${session.dialogData.contractNumber} <br/>Delinquency Status: ${resDelin} <br/>Invoice Due Amount: ${resInvDueAmt} <br/>Invoice Due Date: ${resInvDueDate}`);
            session.send("Thanks for contacting help desk. Good Bye!");
            session.endDialog();         
        } else{
            session.send(`The Contract - ${session.dialogData.contractNumber} doesn't exist. Please try again!`);
            session.send("Good Bye!");
            session.endDialog();           
           }
        });
      }             
  }, 
    function (session, results) {
        session.dialogData.contractNum = results.response;  
        var rowKeyStr = ''
        rowKeyStr = session.dialogData.contractNum
        if(isNaN(rowKeyStr)){
            session.send("Sorry You have entered invalid Contract Number again! Bye.");
            session.endDialog(); 
        } else {
            session.send("We are fetching your contract/party related details..Please Hang On!");
            //Azure Storage database call to fetch contract details
           tableSvc.retrieveEntity(tableName, 'Party',rowKeyStr, function(error, result, response){       
                                    
          if(!error){
          // result contains the entity. Assign to variables
             var resName = result.Name._
             var resDelin = result.DelinquencyStatus._
             var resInvDueAmt =  result.InvoiceDueAmount._
             var resInvDueDate =  result.InvoiceDueDate._
          //    var myJSONObj = { "name":resName,  "delin":resDelin, "invDueAmt":resInvDueAmt, "invDueDate": resInvDueDate};
          //    var myJSONStr =  JSON.stringify(myJSONObj);
      
             // Display the contract details to user in chat.
              session.send(`Welcome : ${resName} !!<br/>         Contract/Party Details - <br/>Contract Number: ${session.dialogData.contractNum} <br/>Delinquency Status: ${resDelin} <br/>Invoice Due Amount: ${resInvDueAmt} <br/>Invoice Due Date: ${resInvDueDate}`);
              session.send("Thanks for contacting help desk. Good Bye!");
              session.endDialog();         
          } else{
              session.send(`The Contract - ${session.dialogData.contractNum} doesn't exist. Please try again!`);
              session.send("Good Bye!");
              session.endDialog();           
             }
         });
      }   
}
]).set('storage', inMemoryStorage); // Register in-memory storage 




/*-----------------------------------------------------------------------------
A simple echo bot for the Microsoft Bot Framework. 
-----------------------------------------------------------------------------*/

var restify = require('restify');
var builder = require('botbuilder');
var azure = require('azure');
var botbuilder_azure = require("botbuilder-azure");
var middleware = require('./middleware.js')
var conf = require('./config.json')

// Setup Restify Server
var server = restify.createServer();
server.listen(process.env.port || process.env.PORT || 3978, function () {
   console.log('%s listening to %s', server.name, server.url); 
});
  
// Create chat connector for communicating with the Bot Framework Service
var connector = new builder.ChatConnector({
    appId: process.env.MicrosoftAppId,
    appPassword: process.env.MicrosoftAppPassword,
    openIdMetadata: process.env.BotOpenIdMetadata 
});

// Listen for messages from users 
server.post('/api/messages', connector.listen());

//Set BOT state storage to CosmosDB
var documentDbOptions = {
    host: 'https://serivice-desk-store.documents.azure.com', 
    masterKey: conf.docdb_master_key, 
    database: 'euc-personal-assistant',   
    collection: 'botdata'
};

var docDbClient = new botbuilder_azure.DocumentDbClient(documentDbOptions);
var cosmosStorage = new botbuilder_azure.AzureBotStorage({ gzipData: false }, docDbClient)

// Create your bot with a function to receive messages from the user
var connectionName = "MicrosoftGraphAuth"
var bot = new builder.UniversalBot(connector, function (session) {
    if (session.message.text == 'signout') {
        // It is important to have a SignOut intent
        connector.signOutUser(session.message.address, connectionName,  (err, result) => {
            if (!err) {
                session.send('You are signed out.');
            } else {
                session.send('There was a problem signing you out.');                
            }
        });
    } else {
        // First check whether the Azure Bot Service already has a token for this user
        connector.getUserToken(session.message.address, connectionName, undefined, (err, result) => {
            if (result) {
                // If there is already a token, the bot can use it directly
                session.send('You are already signed in with token: ' + result.token);
            } else {
                // If there not is already a token, the bot can send an OAuthCard to have the user log in
                if (!session.userData.activeSignIn) {
                    session.send("Hello! Let's get you signed in!");
                    builder.OAuthCard.create(connector, session, connectionName, "Please sign in", "Sign in", (createSignInErr, signInMessage) =>
                    {
                        if (signInMessage) {
                            session.send(signInMessage);
                            session.userData.activeSignIn = true;
                        } else {
                            session.send("Something went wrong trying to sign you in.");
                        }     
                    });
                } else {
                    // Some clients require a 6 digit code validation so we can check that here
                    session.send("Let's see if that code works...");
                    connector.getUserToken(session.message.address, connectionName, session.message.text, (err2, tokenResponse) => {
                        if (tokenResponse) {
                            session.send('It worked! You are now signed in with token: ' + tokenResponse.token);
                            session.userData.activeSignIn = false;
                        } else {
                            session.send("Hmm, that code wasn't right");
                        }
                    });
                }
            }
        });
    }
})
bot.set('storage', cosmosStorage);
bot.use({
    botbuilder: function (session, next) {
        middleware.logIncomingMessage(session, next);
    },
    send: function (event, next) {
        middleware.logOutgoingMessage(event, next);
    }
})

//create servicebus connection
var sb_connection = conf.service_bus
var serviceBusService = azure.createServiceBusService(sb_connection);
serviceBusService.createQueueIfNotExists('send-to-bot', function(error){
    if(error){
        console.log("Failed to Crete Queue")
        process.exit(1)
    }
});

// LUIS Setup

var LUIS_URL = conf.luis
var recognizer = new builder.LuisRecognizer(LUIS_URL);
bot.recognizer(recognizer);

// Intercept trigger event (ActivityTypes.Trigger)
bot.on('trigger', function (message) {
    // handle message from trigger function
    var queuedMessage = message.value;
    var reply = new builder.Message()
        .address(queuedMessage.address)
        .text('This is coming from the trigger: ' + queuedMessage.text);
    bot.send(reply);
});

bot.on("event", (event) => {         // Handle 'event' activities
    if (event.name == 'tokens/response') {
        // received a TokenResponse, which is how the Azure Bot Service responds with the user token after an OAuthCard
        bot.loadSession(event.address, (err, session) => {
            let tokenResponse = event.value;
            session.send('You are now signed in with token: ' + tokenResponse.token);
            session.userData.activeSignIn = false;
        });
    }
});

connector.onInvoke((event, cb) => {
    if (event.name == 'signin/verifyState') {
        // received a MS Team's code verification Invoke Activity
        bot.loadSession(event.address, (err, session) => {
            let verificationCode = event.value.state;
            // Get the user token using the verification code sent by MS Teams
            connector.getUserToken(session.message.address, connectionName, verificationCode, (err, result) => {
                session.send('You are now signed in with token: ' + result.token);
                session.userData.activeSignIn = false;
                cb(undefined, {}, 200);
            });
        });
    } else {
        cb(undefined, {}, 200);
    }
});
const {
  ConnectorClient,
  MicrosoftAppCredentials,
} = require("botframework-connector");
const axios = require("axios");
const mongoose = require("mongoose");
const querystring = require("querystring");
const { BotFrameworkAdapter } = require("botbuilder");

const {
  TeamsActivityHandler,
  CardFactory,
  MessageFactory,
  TurnContext,
  TeamsInfo,
  ActivityHandler,
} = require("botbuilder");
// const { TeamsBot } = require("./teamsBot");
// const bot = new TeamsBot();

const cardTools = require("@microsoft/adaptivecards-tools");
// const rawLearnCard = require("./adaptiveCards/learn.json");
const rawLearnCard = require("./adaptiveCards/scrumfeedback.json");

const connectToBot = async function (usersData, params, context) {
  var credentials = new MicrosoftAppCredentials(
    "a40494cb-f6c9-41da-88de-43b61b59dca6",
    "aaZ8Q~azBSyV0uZtxdh_eYjskN8_kEjBKw4eCbYj"
  );
  var botId = "a40494cb-f6c9-41da-88de-43b61b59dca6";
  var tenantID = "71120744-5a74-413a-bab3-674fdcfaf5e5";

  // var credentials = new MicrosoftAppCredentials(
  //   "273863b0-a2d7-4016-8422-23a212fab473",
  //   "E~s7Q~Nv7URnJY8YGvr5DyIZ~5DTcaglVVjXS"
  // );

  // var recipientId =
  //   "29:1qE9H30mee_yXdjtSBceSgaQppn88q9-LEsvbPOo-qGdTkKobET11mbPsHjRoFtbCWfqz3rZOqgRYggXcsAXqAA";

  var recipientId =
    "29:1Ffz6pSggSxFvzK4PzI9Qfo4B5iKO8R-ExFiE_FEF76aEyIHMU2ck5QcPZ4WZvLWHG55A4mpB62cbl6TwpIlPJA";
  var recipientId2 = "XR70ML@24mlk8.onmicrosoft.com";
  // var recipientId2 = "2c5a1610-b41f-11ec-84bd-e9aa38c51682";

  var recipientId1 =
    "29:16QkUvUkj_9Mz-GAReosDGwXnDJoUZB5eHTdIjA0d2trQ-9TX2tI0vQdmTilKhGXtyhjA9k-A5narEYbGgqg9_w";

  //   var recipientId1 =
  //     "29:1P198ONF5HxqXNjY7bI9YKE2CqiwR0ewJZlijwT--ACk7Q_pM1vrBXvf_wED9mwsprxA0XIP8NKBDplNncze2Gw";

  // var recipientId2 = "xr70ml@24mlk8.onmicrosoft.com"

  var client = new ConnectorClient(credentials, {
    baseUri: "https://smba.trafficmanager.net/amer/",
  });

  let usersObj;
  // if (usersData) {
  //   usersObj = { emails: usersData };
  // }

  usersObj = {
    emails: [
      {
        emailId: "lahiru@bankofsingapore.com",
        value: "lahiru@bankofsingapore.com",
      },
      {
        emailId: "vijaya@bankofsingapore.com",
        value: "vijaya@bankofsingapore.com",
      },
      {
        emailId: "sridhar.kamath@bankofsingapore.com",
        value: "sridhar.kamath@bankofsingapore.com",
      },
      {
        emailId: "kai@bankofsingapore.com",
        value: "kai@bankofsingapore.com",
      },
      {
        emailId: "test123@bankofsingapore.com",
        value: "test123@bankofsingapore.com",
      },
    ],
    ratings: [
      "Colloboration",
      "Respect",
      "Courageous",
      "Focus",
      "Receptiveness",
      "Commitment",
    ],
  };

  const card = cardTools.AdaptiveCards.declare(rawLearnCard).render(usersObj);
  var members = [{ id: recipientId }, { id: recipientId1 }];

  // IsGroup set to true if this is not a direct message (default is false)
  for (var i = 0; i < members.length; i++) {
    var conversationResponse = await client.conversations.createConversation({
      bot: { id: botId },
      members: [{ id: members[i].id }],
      isGroup: false,
      tenantId: tenantID,
    });

    console.log("conversationResponse>>>>", conversationResponse);
    var activityResponse = await client.conversations.sendToConversation(
      conversationResponse.id,
      {
        type: "message",
        attachments: [
          {
            contentType: "application/vnd.microsoft.card.adaptive",
            content: card,
          },
        ],
        from: { id: botId },
        recipient: { id: members[i].id },
      }
    );
  }
  // var conversationResponse = await client.conversations.createConversation({
  //   bot: { id: botId },
  //   members: [{ id: recipientId2 }],
  //   isGroup: false,
  //   tenantId: tenantID,
  // });

  // console.log("conversationResponse>>>>", conversationResponse);
  // var activityResponse = await client.conversations.sendToConversation(
  //   conversationResponse.id,
  //   {
  //     type: "message",
  //     attachments: [
  //       {
  //         contentType: "application/vnd.microsoft.card.adaptive",
  //         content: card,
  //       },
  //     ],
  //     from: { id: botId },
  //     recipient: { id: recipientId },
  //   }
  // );
  //   bot.run(context);
  console.log("activityResponse", activityResponse);
  console.log("Sent reply with ActivityId:", activityResponse.id);
};

// class TeamsBot extends ActivityHandler {
//   constructor(conversationReferences) {
//     super();
//     this.onConversationUpdate(async (context, next) => {
//       console.log("hey");
//       // this.addConversationReference(context.activity);

//       await next();
//     });
//   }
// }

// const onAdaptiveCardInvoke = async function (context, invokeValue) {
//   console.log("on invoke", invokeValue.action.data);
//   mongoose.connect("mongodb://localhost:27017/test", {
//     useNewUrlParser: true,
//     useUnifiedTopology: true,
//   });
//   var db = mongoose.connection;
//   db.on("error", console.error.bind(console, "connection error"));
//   db.once("open", function (callback) {
//     console.log("Connection succeeded.");
//   });
//   var Schema = mongoose.Schema;

//   var feedbackSchema = new Schema({
//     email: String,
//     comments: String,
//   });
//   var Feedback = mongoose.model("Feedback", feedbackSchema);
//   var feedback = new Feedback({
//     email: invokeValue.action.data.CompactSelectVal,
//     comments: invokeValue.action.data.MultiLineVal,
//   });
//   feedback.save(function (error) {
//     console.log("Your bee has been saved!");
//     if (error) {
//       console.error(error);
//     }
//   });
//   await context.sendActivity({
//     type: "message",
//     id: context.activity.replyToId,
//     text: "We appreciate the time you took to share your feedback.",
//   });
//   return { statusCode: 200 };
//   // The verb "userlike" is sent from the Adaptive Card defined in adaptiveCards/learn.json
//   if (invokeValue.action.verb === "userlike") {
//     this.likeCountObj.likeCount++;
//     const card = cardTools.AdaptiveCards.declare(rawLearnCard).render(
//       this.likeCountObj
//     );
//     await context.updateActivity({
//       type: "message",
//       id: context.activity.replyToId,
//       attachments: [CardFactory.adaptiveCard(card)],
//     });
//     return { statusCode: 200 };
//   }
// };

module.exports.connectToBot = connectToBot;

const axios = require("axios");
const mongoose = require("mongoose");
var {
  ConnectorClient,
  MicrosoftAppCredentials,
} = require("botframework-connector");

const querystring = require("querystring");
const {
  TeamsActivityHandler,
  CardFactory,
  MessageFactory,
  TurnContext,
  TeamsInfo,
  ActivityHandler,
} = require("botbuilder");
const rawWelcomeCard = require("./adaptiveCards/welcome.json");
// const rawLearnCard = require("./adaptiveCards/learn.json");
const rawLearnCard = require("./adaptiveCards/scrumfeedback.json");

const cardTools = require("@microsoft/adaptivecards-tools");
mongoose.connect("mongodb://localhost:27017/test", {
  useNewUrlParser: true,
  useUnifiedTopology: true,
});
var db = mongoose.connection;
db.on("error", console.error.bind(console, "connection error"));
db.once("open", function (callback) {
  console.log("Connection succeeded.");
});
var Schema = mongoose.Schema;

var feedbackSchema = new Schema({
  email: String,
  collaboration: String,
  teamculture: String,
  courageous: String,
  qualityOfDilevery: String,
  commitment: String,
  recepitveness: String,
  comments: String,
});
var Feedback = mongoose.model("Feedback", feedbackSchema);
class TeamsBot extends ActivityHandler {
  constructor(conversationReferences) {
    super();
    this.likeCountObj = {
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
      ],
    };

    // Dependency injected dictionary for storing ConversationReference objects used in NotifyController to proactively message users
    this.conversationReferences = conversationReferences;

    // this.onConversationUpdate(async (context, next) => {
    //   this.addConversationReference(context.activity, context);

    //   await next();
    // });

    // record the likeCount
    // this.likeCountObj = { likeCount: 0 };
    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");
      // mongoose.connect("mongodb://localhost:27017/test", {
      //   useNewUrlParser: true,
      //   useUnifiedTopology: true,
      // });
      // var db = mongoose.connection;
      // db.on("error", console.error.bind(console, "connection error"));
      // db.once("open", function (callback) {
      //   console.log("Connection succeeded.");
      // });
      // var Schema = mongoose.Schema;

      // var feedbackSchema = new Schema({
      //   email: String,
      //   comments: String,
      // });
      // var Feedback = mongoose.model("Feedback", feedbackSchema);
      // var feedback = new Feedback({
      //   email: invokeValue.action.data.CompactSelectVal,
      //   comments: invokeValue.action.data.MultiLineVal,
      // });
      // feedback.save(function (error) {
      //   console.log("Your bee has been saved!");
      //   if (error) {
      //     console.error(error);
      //   }
      // });
      // const text = context.activity.text.trim().toLocaleLowerCase();
      // if (text.includes("learn")) {
      //   // await this.messageAllMembersAsync(context);
      // }
      // const teamsChannelId = context.activity.channelId; //msteams
      // const message = MessageFactory.text(
      //   "This will be the first message in a new thread"
      // );
      // const newConversation = await this.teamsCreateConversation(
      //   context,
      //   teamsChannelId,
      //   message
      // );
      // const newConversation = TurnContext.getConversationReference(
      //   context.activity
      // );

      // await context.adapter.continueConversationAsync(
      //   process.env.BOT_ID,
      //   newConversation[0],
      //   async (turnContext) => {
      //     await turnContext.sendActivity(
      //       MessageFactory.text(
      //         "This will be the first response to the new thread"
      //       )
      //     );
      //   }
      // );
      // await next();

      // if (text.includes('message')) {
      //   await this.messageAllMembersAsync(context);
      // }

      let txt = context.activity.text;
      console.log("Text Received.", context);

      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
      );
      console.log("removedMentionText", removedMentionText);
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }
      // Trigger command by IM text
      switch (txt) {
        case "welcome": {
          const card = cardTools.AdaptiveCards.declare(rawWelcomeCard).render();
          await context.sendActivity({
            type: "message",
            attachments: [
              {
                contentType: "application/vnd.microsoft.card.adaptive",
                content: card,
              },
            ],
          });

          break;
        }
        case "learn": {
          console.log("inside learn switch");
          // this.likeCountObj.likeCount = 0;
          // const card = cardTools.AdaptiveCards.declare(rawLearnCard).render(
          //   this.likeCountObj
          // );
          // await context.sendActivity({
          //   attachments: [CardFactory.adaptiveCard(card)],
          // });
          const card = cardTools.AdaptiveCards.declare(rawLearnCard).render(
            this.likeCountObj
          );
          await context.sendActivity({
            type: "message",
            attachments: [
              {
                contentType: "application/vnd.microsoft.card.adaptive",
                content: card,
              },
            ],
          });
          // await this.addConversationReference(context.activity, context);
          break;
        }
        /**
         * case "yourCommand": {
         *   await context.sendActivity(`Add your response here!`);
         *   break;
         * }
         */
      }
      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });

    // this.onConversationUpdate(async (context, next) => {
    //   this.addConversationReference(context.activity);

    //   await next();
    // });

    // Listen to MembersAdded event, view https://docs.microsoft.com/en-us/microsoftteams/platform/resources/bot-v3/bots-notifications for more events
    this.onMembersAdded(async (context, next) => {
      console.log("request to add members");

      const membersAdded = context.activity.membersAdded;
      console.log("new member added", membersAdded);

      for (let cnt = 0; cnt < membersAdded.length; cnt++) {
        if (membersAdded[cnt].id !== context.activity.recipient.id) {
          const welcomeMessage =
            "Welcome to the Proactive Bot sample.  Navigate to http://localhost:3978/api/notify to proactively message everyone who has previously messaged this bot.";
          await context.sendActivity(welcomeMessage);
          break;
        }
      }

      // By calling next() you ensure that the next BotHandler is run.
      await next();
    });
  }

  async messageAllMembersAsync(context) {
    console.log("inside messageAllMembersAsync");
    const members = await this.getPagedMembers(context);
    console.log("members>>>>>>>", members);
    await Promise.all(
      members.map(async (member) => {
        const message = MessageFactory.text(
          `Hello ${member.givenName} ${member.surname}. I'm a Teams conversation bot.`
        );
        console.log("message>>>>>>>", message);
        const convoParams = {
          members: [member],
          tenantId: context.activity.channelData.tenant.id,
          activity: context.activity,
        };
        console.log("convoParams>>>>>>>", convoParams);
        console.log("context.adapter>>>>>>>", context);
        console.log();
        await context.adapter.createConversationAsync(
          context.activity.channelId,
          context.activity.serviceUrl,
          process.env.BOT_ID,
          convoParams,
          null,
          async (context) => {
            console.log("inside createConversationAsync>>>>>>>");
            const ref = TurnContext.getConversationReference(context.activity);

            await context.adapter.continueConversationAsync(
              process.env.BOT_ID,
              ref,
              async (context) => {
                await context.sendActivity("All messages have been sent.");
              }
            );
          }
        );
      })
    );

    // await context.sendActivity(
    //   MessageFactory.text("All messages have been sent.")
    // );
  }

  // async connectToBot() {
  //   var credentials = new MicrosoftAppCredentials(
  //     "97bf0eaa-0698-4ba7-b1de-29cb2b86400f",
  //     "t1h7Q~PsAabYfc5Jfi2R2d1JeiMIeVN6UrAuz"
  //   );

  //   var botId = "97bf0eaa-0698-4ba7-b1de-29cb2b86400f";
  //   var recipientId =
  //     "29:1qE9H30mee_yXdjtSBceSgaQppn88q9-LEsvbPOo-qGdTkKobET11mbPsHjRoFtbCWfqz3rZOqgRYggXcsAXqAA";

  //   var client = new ConnectorClient(credentials, {
  //     baseUri: "https://smba.trafficmanager.net/amer/",
  //   });

  //   var conversationResponse = await client.conversations.createConversation({
  //     bot: { id: botId },
  //     members: [{ id: recipientId }],
  //     isGroup: false,
  //     tenantId: "71120744-5a74-413a-bab3-674fdcfaf5e5",
  //   });

  //   console.log("conversationResponse", conversationResponse);
  //   var activityResponse = await client.conversations.sendToConversation(
  //     conversationResponse.id,
  //     {
  //       type: "message",
  //       from: { id: botId },
  //       recipient: { id: recipientId },
  //       text: "This a message from Bot Connector Client (NodeJS)",
  //     }
  //   );

  //   console.log("Sent reply with ActivityId:", activityResponse.id);
  // }

  async getPagedMembers(context) {
    let continuationToken;
    const members = [];

    do {
      const page = await TeamsInfo.getPagedMembers(
        context,
        100,
        continuationToken
      );

      continuationToken = page.continuationToken;

      members.push(...page.members);
    } while (continuationToken !== undefined);

    return members;
  }

  async teamsCreateConversation(context, teamsChannelId, message) {
    console.log("Inside teamsCreateConversation");
    const conversationParameters = {
      isGroup: false,
      channelData: {
        channel: {
          id: teamsChannelId,
        },
      },

      activity: {
        type: "message",
        text: "This will be the first message in a new thread",
      },
    };
    console.log("conversationParameters", conversationParameters);
    console.log("context>>>>>>>>>>", JSON.stringify(context));

    const connectorFactory = context.turnState.get(
      context.adapter.ConnectorFactoryKey
    );
    console.log("connectorFactory", connectorFactory);

    const connectorClient = await connectorFactory.create(
      "https://smba.trafficmanager.net/amer/"
    );

    console.log("connectorClient", connectorClient);

    const conversationResourceResponse =
      await connectorClient.conversations.createConversation(
        conversationParameters
      );
    const conversationReference = TurnContext.getConversationReference(
      context.activity
    );
    // conversationReference.conversation.id = conversationResourceResponse.id;
    conversationReference.conversation.id =
      "a:1H0K8w32hvLWLXAJifSb4Nryu2V6u5y0IjCedAPXdmiUwyTFQ_V9q3Oj6IT2pySP1trx9lsEVDTSM-O8iq1Qgl1_d4xzUJM4mmMi_KSzoXzyLpyPhQ_AIdexgAblAfZ4V";
    return [conversationReference, conversationResourceResponse.activityId];
  }

  async addConversationReference(activity, context) {
    console.log("Inside conversationReference");
    const conversationReference =
      TurnContext.getConversationReference(activity);
    console.log("conversationReference>>>>>>", conversationReference);
    this.conversationReferences[conversationReference.conversation.id] =
      conversationReference;

    const convoParams = {
      members: [
        {
          id: "29:1qE9H30mee_yXdjtSBceSgaQppn88q9-LEsvbPOo-qGdTkKobET11mbPsHjRoFtbCWfqz3rZOqgRYggXcsAXqAA",
          name: "Deepika Ashok Chalpe",
          aadObjectId: "825bd6a4-acdc-47c6-a6f1-4f4d8bac27f9",
        },
      ],
      tenantId: conversationReference.conversation.tenantId,
      activity: conversationReference.activityId,
    };

    await context.adapter.createConversationAsync(
      process.env.BOT_ID,
      conversationReference.channelId,
      conversationReference.serviceUrl,
      null,
      convoParams,
      async (context) => {
        console.log("inside createConversationAsync");
        const ref = TurnContext.getConversationReference(context.activity);
        console.log("after ref", ref);

        await context.adapter.continueConversationAsync(
          process.env.BOT_ID,
          ref,
          async (context) => {
            await context.sendActivity("Hey Deepikaaaa");
          }
        );
      }
    );
  }
  // Invoked when an action is taken on an Adaptive Card. The Adaptive Card sends an event to the Bot and this
  // method handles that event.
  async onAdaptiveCardInvoke(context, invokeValue) {
    console.log("on invoke", invokeValue);
    const propertyNames = Object.keys(invokeValue.action.data);
    let allArray = [];
    const ratingArray = [];
    let currentName = "";
    let currentObject = {};
    const obj = invokeValue.action.data;
    let dKey = "";
    for (const propertyName of propertyNames) {
      let temp = propertyName;
      let a = propertyName.substring(0, propertyName.length - 1);
      let b = propertyName.slice(propertyName.length - 1);

      if (currentName === "" || currentName === a) {
        currentName = a;
        dKey = "";
        if (b == 1) {
          dKey = "collaboration";
        } else if (b == 2) {
          dKey = "teamculture";
        } else if (b == 3) {
          dKey = "courageous";
        } else if (b == 4) {
          dKey = "qualityOfDilevery";
        } else if (b == 5) {
          dKey = "commitment";
        } else if (b == 6) {
          dKey = "receptiveness";
        } else {
          dKey = "comments";
        }
        currentObject[dKey] = obj[temp];
      } else {
        dKey = "";
        if (b == 1) {
          dKey = "collaboration";
        } else if (b == 2) {
          dKey = "teamculture";
        } else if (b == 3) {
          dKey = "courageous";
        } else if (b == 4) {
          dKey = "qualityOfDilevery";
        } else if (b == 5) {
          dKey = "commitment";
        } else if (b == 6) {
          dKey = "receptiveness";
        } else {
          dKey = "comments";
        }
        currentObject["email"] = currentName;
        currentObject[dKey] = obj[temp];
        allArray.push(currentObject);
        currentObject = {};
        currentName = a;
      }
    }
    // mongoose.connect("mongodb://localhost:27017/test", {
    //   useNewUrlParser: true,
    //   useUnifiedTopology: true,
    // });
    // var db = mongoose.connection;
    // db.on("error", console.error.bind(console, "connection error"));
    // db.once("open", function (callback) {
    //   console.log("Connection succeeded.");
    // });
    // var Schema = mongoose.Schema;

    // var feedbackSchema = new Schema({
    //   email: String,
    //   comments: String,
    // });
    // var Feedback = mongoose.model("Feedback", feedbackSchema);
    var resFeedbackArray = [];
    var ratings = [];
    // resFeedbackArray.push({
    //   email: invokeValue.action.data.
    //   ratings
    // })
    for (var i = 0; i < allArray.length; i++) {
      var feedback = new Feedback(allArray[i]);
      feedback.save(function (error) {
        console.log("Your bee has been saved!");
        if (error) {
          console.error(error);
        }
      });
    }

    await context.sendActivity({
      type: "message",
      id: context.activity.replyToId,
      text: "We appreciate the time you took to share your feedback.",
    });
    if (invokeValue.action.verb === "userlike") {
      this.likeCountObj.likeCount++;
      const card = cardTools.AdaptiveCards.declare(rawLearnCard).render(
        this.likeCountObj
      );
      await context.updateActivity({
        type: "message",
        id: context.activity.replyToId,
        attachments: [CardFactory.adaptiveCard(card)],
      });
      return { statusCode: 200 };
    }
    // The verb "userlike" is sent from the Adaptive Card defined in adaptiveCards/learn.json
  }

  // Messaging extension Code
  // Action.
  handleTeamsMessagingExtensionSubmitAction(context, action) {
    switch (action.commandId) {
      case "createCard":
        return createCardCommand(context, action);
      case "shareMessage":
        return shareMessageCommand(context, action);
      default:
        throw new Error("NotImplemented");
    }
  }

  // Search.
  async handleTeamsMessagingExtensionQuery(context, query) {
    const searchQuery = query.parameters[0].value;
    const response = await axios.get(
      `http://registry.npmjs.com/-/v1/search?${querystring.stringify({
        text: searchQuery,
        size: 8,
      })}`
    );

    const attachments = [];
    response.data.objects.forEach((obj) => {
      const heroCard = CardFactory.heroCard(obj.package.name);
      const preview = CardFactory.heroCard(obj.package.name);
      preview.content.tap = {
        type: "invoke",
        value: { name: obj.package.name, description: obj.package.description },
      };
      const attachment = { ...heroCard, preview };
      attachments.push(attachment);
    });

    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: attachments,
      },
    };
  }

  async handleTeamsMessagingExtensionSelectItem(context, obj) {
    return {
      composeExtension: {
        type: "result",
        attachmentLayout: "list",
        attachments: [CardFactory.heroCard(obj.name, obj.description)],
      },
    };
  }

  // Link Unfurling.
  handleTeamsAppBasedLinkQuery(context, query) {
    const attachment = CardFactory.thumbnailCard("Thumbnail Card", query.url, [
      query.url,
    ]);

    const result = {
      attachmentLayout: "list",
      type: "result",
      attachments: [attachment],
    };

    const response = {
      composeExtension: result,
    };
    return response;
  }
}

function createCardCommand(context, action) {
  // The user has chosen to create a card by choosing the 'Create Card' context menu command.
  const data = action.data;
  const heroCard = CardFactory.heroCard(data.title, data.text);
  heroCard.content.subtitle = data.subTitle;
  const attachment = {
    contentType: heroCard.contentType,
    content: heroCard.content,
    preview: heroCard,
  };

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: [attachment],
    },
  };
}

function shareMessageCommand(context, action) {
  // The user has chosen to share a message by choosing the 'Share Message' context menu command.
  let userName = "unknown";
  if (
    action.messagePayload &&
    action.messagePayload.from &&
    action.messagePayload.from.user &&
    action.messagePayload.from.user.displayName
  ) {
    userName = action.messagePayload.from.user.displayName;
  }

  // This Messaging Extension example allows the user to check a box to include an image with the
  // shared message.  This demonstrates sending custom parameters along with the message payload.
  let images = [];
  const includeImage = action.data.includeImage;
  if (includeImage === "true") {
    images = [
      "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcQtB3AwMUeNoq4gUBGe6Ocj8kyh3bXa9ZbV7u1fVKQoyKFHdkqU",
    ];
  }
  const heroCard = CardFactory.heroCard(
    `${userName} originally sent this message:`,
    action.messagePayload.body.content,
    images
  );

  if (
    action.messagePayload &&
    action.messagePayload.attachment &&
    action.messagePayload.attachments.length > 0
  ) {
    // This sample does not add the MessagePayload Attachments.  This is left as an
    // exercise for the user.
    heroCard.content.subtitle = `(${action.messagePayload.attachments.length} Attachments not included)`;
  }

  const attachment = {
    contentType: heroCard.contentType,
    content: heroCard.content,
    preview: heroCard,
  };

  return {
    composeExtension: {
      type: "result",
      attachmentLayout: "list",
      attachments: [attachment],
    },
  };
}

module.exports.TeamsBot = TeamsBot;

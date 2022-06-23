// index.js is used to setup and configure your bot

// Import required packages
const restify = require("restify");
var bodyParser = require("body-parser");

// Import required bot services.
// See https://aka.ms/bot-services to learn more about the different parts of a bot.
const { BotFrameworkAdapter } = require("botbuilder");
const { TeamsBot } = require("./teamsBot");
const FeedbackUtil = require("./scrum-feedback-util");

// This bot's main dialog.
// const { ProactiveBot } = require("./bots/proactiveBot");
// Create the main dialog.
const conversationReferences = {};
// const bot = new ProactiveBot(conversationReferences);
// Create adapter.
// See https://aka.ms/about-bot-adapter to learn more about adapters.
const adapter = new BotFrameworkAdapter({
  appId: process.env.BOT_ID,
  appPassword: process.env.BOT_PASSWORD,
});

adapter.onTurnError = async (context, error) => {
  // This check writes out errors to console log .vs. app insights.
  // NOTE: In production environment, you should consider logging this to Azure
  //       application insights. See https://aka.ms/bottelemetry for telemetry
  //       configuration instructions.
  console.error(`\n [onTurnError] unhandled error: ${error}`);

  // Send a message to the user
  await context.sendActivity(
    `The bot encountered an unhandled error:\n ${error.message}`
  );
  await context.sendActivity(
    "To continue to run this bot, please fix the bot source code."
  );
};

// Create the bot that will handle incoming messages.
const bot = new TeamsBot(conversationReferences);
// Create the main dialog.

// Create HTTP server.
const server = restify.createServer();
server.use(restify.plugins.queryParser());
server.use(restify.plugins.bodyParser());

server.listen(process.env.port || process.env.PORT || 3978, function () {
  console.log(`\nBot started, ${server.name} listening to ${server.url}`);
});

// Listen for incoming requests.
server.post("/api/messages", async (req, res) => {
  console.log("inside messages");

  await adapter.processActivity(req, res, async (context) => {
    await bot.run(context);
  });
});

// Listen for incoming requests.
server.post("/jira-event-listener", async (req, res) => {
  console.log("inside webhook", req.body);
  var usersData = req.body;
  var params = req.params;
  console.log("params", params, req.query);
  const newConversation = await FeedbackUtil.connectToBot(usersData, params);
  res.json({ status: "success" });
});

// Listen for incoming notifications and send proactive messages to users.
server.get("/api/notify", async (req, res) => {
  console.log("inside notify");

  for (const conversationReference of Object.values(conversationReferences)) {
    await adapter.continueConversationAsync(
      process.env.BOT_ID,
      conversationReference,
      async (context) => {
        await context.sendActivity("proactive hello");
      }
    );
  }
  res.setHeader("Content-Type", "text/html");
  res.writeHead(200);
  res.write(
    "<html><body><h1>Proactive messages have been sent.</h1></body></html>"
  );
  res.end();
});
// Gracefully shutdown HTTP server
[
  "exit",
  "uncaughtException",
  "SIGINT",
  "SIGTERM",
  "SIGUSR1",
  "SIGUSR2",
].forEach((event) => {
  process.on(event, () => {
    server.close();
  });
});

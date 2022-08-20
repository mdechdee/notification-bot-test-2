const notificationTemplate = require("./adaptiveCards/notification-default.json");
const { AdaptiveCards } = require("@microsoft/adaptivecards-tools");
const { bot } = require("./internal/initialize");
const { getPriorityMessagesClient } = require("./getPriorityMessages");

// HTTP trigger to send notification. You need to add authentication / authorization for this API. Refer https://aka.ms/teamsfx-notification for more details.
module.exports = async function (context, req) {
  let responseBody = "HELL WORLD";
  try{
    const response = await getPriorityMessagesClient.get("HttpExample");
    responseBody = response.body;
  }catch(err){
    responseBody = "Failed to call getPriorityMessages endpoint: " + err;
  }

  for (const target of await bot.notification.installations()) {
    await target.sendAdaptiveCard(
      AdaptiveCards.declare(notificationTemplate).render({
        title: "Check out this message you might have missed!",
        appName: "Noti Application",
        description: `THIS IS SAMPLE MESSAGE AND A MOCKUP URL: ${responseBody}`,
        notificationUrl: "https://teams.microsoft.com/l/message/19:4ae96d30c78d4b1381e123cb16786709@thread.tacv2/1661006752435?tenantId=4b0a13ca-95ed-42b4-a17b-914261b8c920&groupId=dcc10f25-9096-4a23-9ae3-632fd490f539&parentMessageId=1661006752435&teamName=Sales%20and%20Marketing&channelName=General&createdTime=1661006752435&allowXTenantAccess=false",
      })
    );
  }

  /****** To distinguish different target types ******/
  /** "Channel" means this bot is installed to a Team (default to notify General channel)
  if (target.type === "Channel") {
    // Directly notify the Team (to the default General channel)
    await target.sendAdaptiveCard(...);
    // List all channels in the Team then notify each channel
    const channels = await target.channels();
    for (const channel of channels) {
      await channel.sendAdaptiveCard(...);
    }
    // List all members in the Team then notify each member
    const members = await target.members();
    for (const member of members) {
      await member.sendAdaptiveCard(...);
    }
  }
  **/

  /** "Group" means this bot is installed to a Group Chat
  if (target.type === "Group") {
    // Directly notify the Group Chat
    await target.sendAdaptiveCard(...);
    // List all members in the Group Chat then notify each member
    const members = await target.members();
    for (const member of members) {
      await member.sendAdaptiveCard(...);
    }
  }
  **/

  /** "Person" means this bot is installed as a Personal app
  if (target.type === "Person") {
    // Directly notify the individual person
    await target.sendAdaptiveCard(...);
  }
  **/

  context.res = {};
};

// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

const {
    TurnContext,
    MessageFactory,
    TeamsActivityHandler,
    CardFactory,
    ActionTypes,
    ActivityTypes
} = require('botbuilder');

const WelcomeCard = require("./WelcomeCard.json");

class BotActivityHandler extends TeamsActivityHandler {
    constructor() {
        super();
        /* Conversation Bot */
        /*  Teams bots are Microsoft Bot Framework bots.
            If a bot receives a message activity, the turn handler sees that incoming activity
            and sends it to the onMessage activity handler.
            Learn more: https://aka.ms/teams-bot-basics.

            NOTE:   Ensure the bot endpoint that services incoming conversational bot queries is
                    registered with Bot Framework.
                    Learn more: https://aka.ms/teams-register-bot. 
        */
        // Registers an activity event handler for the message event, emitted for every incoming message activity.
        this.onMessage(async (context, next) => {
            TurnContext.removeRecipientMention(context.activity);

            //#region  new code for conversational bot --11/10/2020 by Kanakarajulu Thota
            switch (context.activity.type) {
                case ActivityTypes.Message:
                    let text = TurnContext.removeRecipientMention(context.activity);
                    // text = text.toLocaleLowerCase();

                    let chatroomParticipants = "";
                    let RecipientsInformation = "";

                    text = text.toLowerCase().toString();
                    if (text.startsWith("hello")) {
                        await this.mentionActivityAsync(context);
                        return;
                    } else if (text.startsWith("chatroom")) {
                        chatroomParticipants = text.split("chatroom")[1].trim();
                        RecipientsInformation = "controlbot@microsoftO365competency.onmicrosoft.com," + chatroomParticipants;
                        
                        if(RecipientsInformation.includes("@")) {
                            // await context.sendActivity(RecipientsInformation);
                            let OpenURL = "https://teams.microsoft.com/l/chat/0/0?users=" + RecipientsInformation + "&topicName=Call%20from%20Siemens%20Service%20Desk";
                            let userInfo = context.activity.from.name;
                            WelcomeCard.body[1].text = "Hello " + userInfo + ", nice to meet you!";
                            WelcomeCard.actions[0].url = OpenURL;
                            const welcomeCard = CardFactory.adaptiveCard(WelcomeCard);
                            
                            // welcomeCard.contentUrl = "https://teams.microsoft.com/l/chat/0/0?users=suman.sarkar@microsoftO365competency.onmicrosoft.com,sourav.mandal@microsoftO365competency.onmicrosoft.com,sayan.sinha@microsoftO365competency.onmicrosoft.com&topicName=Call%20from%20Siemens%20Service%20Desk";
                            
                            await context.sendActivity({
                                attachments: [welcomeCard]
                            });
                            return;
                        }
                        else 
                        {
                            await context.sendActivity("Send participants list by comma seperated !");
                            return;
                        }                        
                    } else {
                        await context.sendActivity(`I\'m terribly sorry, but my master hasn\'t trained me to do anything yet...`);
                        return;
                    }
                    break;
                default:
                    break;
            }
            await next();
            //#endregion

            //#region old code
            // switch (context.activity.text.trim()) {
            // case 'Hello':
            //     await this.mentionActivityAsync(context);
            //     break;
            // case 'Chatroom':
            //     await this.sendActivityAsync(context);
            //     break;
            // default:
            //     // By default for unknown activity sent by user show
            //     // a card with the available actions.
            //     const value = { count: 0 };
            //     const card = CardFactory.heroCard(
            //         'Lets talk...',
            //         null,
            //         [{
            //             type: ActionTypes.MessageBack,
            //             title: 'Say Hello',
            //             value: value,
            //             text: 'Hello'
            //         }]);
            //     await context.sendActivity({ attachments: [card] });
            //     break;
            // }
            // await next();
            //#endregion

        });
        /* Conversation Bot */
    }

    /* Conversation Bot */
    /**
     * Say hello and @ mention the current user.
     */
    async mentionActivityAsync(context) {
        const TextEncoder = require('html-entities').XmlEntities;

        const mention = {
            mentioned: context.activity.from,
            text: `<at>${ new TextEncoder().encode(context.activity.from.name) }</at>`,
            type: 'mention'
        };

        const replyActivity = MessageFactory.text(`Hi ${ mention.text }`);
        replyActivity.entities = [mention];
        
        await context.sendActivity(replyActivity);
    }
    /* Conversation Bot */

    /* Conversation Bot */
    /**
     * Say hello and @ mention the current user.
     */
    async sendActivityAsync(context) {
        const TextEncoder = require('html-entities').XmlEntities;

        const userName = {
            mentioned: context.text,
            text: context.text,
            type: 'message'
        };

        const mention = {
            mentioned: context.activity.from,
            text: `<at>${ new TextEncoder().encode(context.activity.from.name) }</at>`,
            type: 'mention'
        };

        const replyActivity = MessageFactory.text(`Hi ${ userName.text }`);
        replyActivity.entities = [userName];
        
        await context.sendActivity(replyActivity);
    }
    /* Conversation Bot */

}

module.exports.BotActivityHandler = BotActivityHandler;


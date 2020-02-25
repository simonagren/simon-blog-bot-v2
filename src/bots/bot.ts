// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

import { Activity, ActivityHandler, ActivityTypes, Mention, TurnContext} from 'botbuilder';

export class SimonBot extends ActivityHandler {
    constructor() {
        super();
        // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
        this.onMessage(async (context, next) => {

            if (context.activity.channelId === 'msteams') {
                await this._messageWithMention(context);
            } else {
                await context.sendActivity(`You said '${ context.activity.text }'`);
            }
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            for (const member of membersAdded) {
                if (member.id !== context.activity.recipient.id) {
                    await context.sendActivity('Hello and welcome!');
                }
            }
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }

    private async _messageWithMention(context: TurnContext): Promise<void> {
        // Create mention object
        const mention: Mention = {
            mentioned: context.activity.from,
            text: `<at>${context.activity.from.name}</at>`,
            type: 'mention'
        };

        // Construct message to send
        const message: Partial<Activity> = {
            entities: [mention],
            text: `${mention.text} You said '${ context.activity.text }'`,
            type: ActivityTypes.Message
        };

        await context.sendActivity(message);
    }

}

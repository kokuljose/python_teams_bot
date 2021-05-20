# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.
import time
import sys
import traceback

from datetime import datetime,date
from http import HTTPStatus
from typing import Dict, List
from aiohttp import web
from aiohttp.web import Request, Response, json_response
from botbuilder.core import (
    BotFrameworkAdapterSettings,
    TurnContext,
    BotFrameworkAdapter,
    MessageFactory,
    CardFactory
)
from botbuilder.core.integration import aiohttp_error_middleware
from botbuilder.schema import (
    Activity,
    ActivityTypes,
    ConversationReference,
    CardAction,
    ActionTypes,
    HeroCard,
    SuggestedActions
)

from bots import TeamsFileUploadBot
from config import DefaultConfig

CONFIG = DefaultConfig()

# Create adapter.
# See https://aka.ms/about-bot-adapter to learn more about how bots work.
SETTINGS = BotFrameworkAdapterSettings(CONFIG.APP_ID, CONFIG.APP_PASSWORD)
ADAPTER = BotFrameworkAdapter(SETTINGS)


# Catch-all for errors.
async def on_error(context: TurnContext, error: Exception):
    # This check writes out errors to console log .vs. app insights.
    # NOTE: In production environment, you should consider logging this to Azure
    #       application insights.
    print(f"\n [on_turn_error] unhandled error: {error}", file=sys.stderr)
    traceback.print_exc()

    # Send a message to the user
    await context.send_activity("The bot encountered an error or bug.")
    await context.send_activity(
        "To continue to run this bot, please fix the bot source code."
    )
    # Send a trace activity if we're talking to the Bot Framework Emulator
    if context.activity.channel_id == "emulator":
        # Create a trace activity that contains the error object
        trace_activity = Activity(
            label="TurnError",
            name="on_turn_error Trace",
            timestamp=datetime.utcnow(),
            type=ActivityTypes.trace,
            value=f"{error}",
            value_type="https://www.botframework.com/schemas/error",
        )
        # Send a trace activity, which will be displayed in Bot Framework Emulator
        await context.send_activity(trace_activity)


ADAPTER.on_turn_error = on_error
# Create a shared dictionary.  The Bot will add conversation references when users
# join the conversation and send messages.
CONVERSATION_REFERENCES: Dict[str, ConversationReference] = dict()

# If the channel is the Emulator, and authentication is not in use, the AppId will be null.
# We generate a random AppId for this case only. This is not required for production, since
# the AppId will have a value.
APP_ID = SETTINGS.app_id if SETTINGS.app_id else uuid.uuid4()
# Create the Bot
BOT = TeamsFileUploadBot(CONFIG.APP_ID, CONFIG.APP_PASSWORD,CONVERSATION_REFERENCES)


# Listen for incoming requests on /api/messages.s
async def messages(req: Request) -> Response:
    # Main bot message handler.
    if "application/json" in req.headers["Content-Type"]:
        body = await req.json()
    else:
        return Response(status=HTTPStatus.UNSUPPORTED_MEDIA_TYPE)

    activity = Activity().deserialize(body)
    auth_header = req.headers["Authorization"] if "Authorization" in req.headers else ""

    invoke_response = await ADAPTER.process_activity(
        activity, auth_header, BOT.on_turn
    )
    if invoke_response:
        return json_response(
            data=invoke_response.body, status=invoke_response.status
        )
    return Response(status=HTTPStatus.OK)

# Listen for requests on /api/notify, and send a messages to all conversation members.
async def notify(req: Request) -> Response:  # pylint: disable=unused-argument
    await _send_proactive_message()
    return Response(status=HTTPStatus.OK, text="Messages have been sent")

async def teams_create_conversation(self, turn_context: TurnContext, teams_channel_id: str, message):
    params = ConversationParameters(
                                        is_group=True,
                                        channel_data={"channel": {"id": teams_channel_id}},
                                        activity=message
                                        )


    connector_client = await turn_context.adapter.create_connector_client(turn_context.activity.service_url)
    conversation_resource_response = await connector_client.conversations.create_conversation(params)
    conversation_reference = TurnContext.get_conversation_reference(turn_context.activity)
    conversation_reference.conversation.id = conversation_resource_response.id
    return [conversation_reference, conversation_resource_response.activity_id]

# Send a message to all conversation members.
# This uses the shared Dictionary that the Bot adds conversation references to.
async def _send_proactive_message():
    for conversation_reference in CONVERSATION_REFERENCES.values():
        reply = MessageFactory.list([])
        card = HeroCard(
            text=f"Hello, {conversation_reference.user.name} today is {date.today().strftime('%B %d, %Y')}, would you like to see the report?",
            buttons=[
                CardAction(
                    type=ActionTypes.im_back, title="Yes", value="Yes, I want to see the Report."
                ),
                CardAction(
                    type=ActionTypes.im_back, title="No", value="No, I don't want to see the Report."
                ),
            ],
        )
        reply.attachments.insert(0,CardFactory.hero_card(card))
        await ADAPTER.continue_conversation(
            conversation_reference,
            lambda turn_context: turn_context.send_activity(reply),
            APP_ID,
        )

        #reply = MessageFactory.list([])
        #reply.attachments.append(self._send_suggested_actions_yes_no(member.name))
        #return await turn_context.send_activity(reply)

    def _send_suggested_actions_yes_no(name:str) -> Attachment:
        card = HeroCard(
            text=f"Hello, {name} today is {date.today().strftime('%B %d, %Y')}, would you like to see the report?",
            buttons=[
                CardAction(
                    type=ActionTypes.im_back, title="Yes", value="Yes, I want to see the Report."
                ),
                CardAction(
                    type=ActionTypes.im_back, title="No", value="No, I don't want to see the Report."
                ),
            ],
        )

        return CardFactory.hero_card(card)
def init_func(argv):
    APP = web.Application(middlewares=[aiohttp_error_middleware])
    APP.router.add_post("/api/messages", messages)
    APP.router.add_get("/api/notify", notify)
    return APP

if __name__ == "__main__":
    APP = init_func(None)
    try:
        web.run_app(APP, host="0.0.0.0", port=CONFIG.PORT)
    except Exception as error:
        raise error

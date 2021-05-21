# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License.

from datetime import datetime,date
import os
import csv
import requests
import time
from typing import List, Dict
from botbuilder.core import TurnContext,ActivityHandler, MessageFactory, CardFactory
from botbuilder.core.teams import TeamsActivityHandler, TeamsInfo,teams_get_channel_id
from botbuilder.schema import (
    Activity,
    Attachment,
    ChannelAccount,
    ActivityTypes,
    ConversationAccount,
    ConversationParameters,
    ConversationReference,
    Attachment,
    CardAction,
    ActionTypes,
    HeroCard,
    SuggestedActions
)
from botbuilder.schema.teams import (
    FileDownloadInfo,
    FileConsentCard,
    FileConsentCardResponse,
    FileInfoCard,
    TeamInfo,
    TeamsChannelAccount
)
from botbuilder.schema.teams.additional_properties import ContentType


class TeamsFileUploadBot(TeamsActivityHandler):

    def __init__(self, app_id: str, app_password: str,conversation_references: Dict[str, ConversationReference]):
        self._app_id = app_id
        self._app_password = app_password
        self.conversation_references = conversation_references
        

    async def on_conversation_update_activity(self, turn_context: TurnContext):
        await self._add_conversation_reference(turn_context)
        
        return await super().on_conversation_update_activity(turn_context)

    async def on_teams_members_added(  # pylint: disable=unused-argument
            self,
            teams_members_added: [TeamsChannelAccount],
            team_info: TeamInfo,
            turn_context: TurnContext,
    ):
        await self._add_conversation_reference(turn_context)
        for member in teams_members_added:
            if member.id != turn_context.activity.recipient.id:
                await turn_context.send_activity(
                    f"Welcome to the team {member.given_name} {member.surname}. "
                )


    async def on_message_activity(self, turn_context: TurnContext):
        
        await self._add_conversation_reference(turn_context)
        message_with_file_download = (
            False
            if not turn_context.activity.attachments
            else turn_context.activity.attachments[0].content_type == ContentType.FILE_DOWNLOAD_INFO
        )

        if message_with_file_download:
            file = turn_context.activity.attachments[0]
            file_download = FileDownloadInfo.deserialize(file.content)
            file_path = "files/" + file.name

            response = requests.get(file_download.download_url, allow_redirects=True)
            open(file_path, "wb").write(response.content)
            with open(file_path) as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=',')
                line_count = 0
                for row in csv_reader:
                    if line_count == 0:
                        head=", ".join(row)
            reply = self._create_reply(
                turn_context.activity, f"Your parameters have been updated using the template in <b>{file.name}</b> with Headings <b>{head}</b>", "xml"
            )
            await turn_context.send_activity(reply)
        elif turn_context.activity.text != None:
            text = turn_context.activity.text
            filename = "report.csv"
            file_path = "files/" + filename
            file_size = os.path.getsize(file_path)
            await self._process_input(turn_context,text,filename, file_size)
        else:
            filename = "report.csv"
            reply = self._create_reply(
                turn_context.activity, f"Hello, Rod today is {date.today().strftime('%B %d, %Y')}, would you like to see the report?", "xml"
            )
            await turn_context.send_activity(reply)
            reply = MessageFactory.list([])
            reply.attachments.append(self._send_suggested_actions_yes_no("Rod"))
            await turn_context.send_activity(reply)

    def _send_suggested_actions_yes_no(self,name:str) -> Attachment:
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

    async def _process_input(self,turn_context: TurnContext, text: str, filename: str, file_size: int):

        if text.find("hello")!=-1:
            reply = MessageFactory.list([])
            reply.attachments.append(self._send_suggested_actions_yes_no(member.name))
            await turn_context.send_activity(reply)

        elif text.find("MessageAllMembers")!=-1:

            await self._message_all_members(turn_context)
            reply = self._create_reply(
                turn_context.activity,
                f"Successfully Messaged All Members", "xml"
            )
            await turn_context.send_activity(reply)

        elif text.find("Yes, I want to see the Report.")!=-1 or text.find("report")!=-1:
            await self._send_file_card(turn_context, filename, file_size)
            reply = self._create_reply(
                turn_context.activity,
                f"Please type 'settings' to update report settings", "xml"
            )
            await turn_context.send_activity(reply)

        elif text.find("No, I don't want to see the Report.")!=-1:
            reply = self._create_reply(
                turn_context.activity,
                f"ThankYou. Get back to me when you need it. I'm here to serve you!", "xml"
            )
            await turn_context.send_activity(reply)


        elif text.find("settings")!=-1:
            reply = self._create_reply(
                turn_context.activity,
                f"Would you like to update report parameters or the options for this report?", "xml"
            )
            await turn_context.send_activity(reply)
            reply=MessageFactory.list([])
            reply.attachments.append(self._send_suggested_actions_reportparameters_options())
            await turn_context.send_activity(reply)

        elif text.find("Update Report Parameters for Report")!=-1:
            await self._send_file_card(turn_context, filename, file_size)
            reply = self._create_reply(
                turn_context.activity,
                f"Please update the template and upload.", "xml"
            )
            await turn_context.send_activity(reply)

        elif text.find("Update Options for Report")!=-1:
            reply = self._create_reply(
                turn_context.activity,
                f"What threshold would you like to set for this report?", "xml"
            )
            await turn_context.send_activity(reply)

        elif all([xi in '1234567890' for xi in text.lstrip('-')]):
            reply = self._create_reply(
                turn_context.activity,
                f"Thanks your new threshold is {text}", "xml"
            )
            await turn_context.send_activity(reply)
        else:
            reply = self._create_reply(
                turn_context.activity,
                f"Sorry, I couldn't Understand. Please Enter a valid Value.", "xml"
            )
            await turn_context.send_activity(reply)

    def _send_suggested_actions_reportparameters_options(self) -> Attachment:
        card = HeroCard(
            text="Would you like to update report parameters or the options for this report?",
            buttons=[
                CardAction(
                    type=ActionTypes.im_back, title="Report Parameters", value="Update Report Parameters for Report"
                ),
                CardAction(
                    type=ActionTypes.im_back, title="Options", value="Update Options for Report"
                ),
            ],
        )

        return CardFactory.hero_card(card)

    async def _send_file_card(
            self, turn_context: TurnContext, filename: str, file_size: int
    ):
        """
        Send a FileConsentCard to get permission from the user to upload a file.
        """

        consent_context = {"filename": filename}

        file_card = FileConsentCard(
            description="This is the file I want to send you",
            size_in_bytes=file_size,
            accept_context=consent_context,
            decline_context=consent_context
        )

        as_attachment = Attachment(
            content=file_card.serialize(), content_type=ContentType.FILE_CONSENT_CARD, name=filename
        )

        reply_activity = self._create_reply(turn_context.activity)
        reply_activity.attachments = [as_attachment]
        await turn_context.send_activity(reply_activity)

    async def on_teams_file_consent_accept(
            self,
            turn_context: TurnContext,
            file_consent_card_response: FileConsentCardResponse
    ):
        """
        The user accepted the file upload request.  Do the actual upload now.
        """

        file_path = "files/" + file_consent_card_response.context["filename"]
        file_size = os.path.getsize(file_path)

        headers = {
            "Content-Length": f"\"{file_size}\"",
            "Content-Range": f"bytes 0-{file_size-1}/{file_size}"
        }
        response = requests.put(
            file_consent_card_response.upload_info.upload_url, open(file_path, "rb"), headers=headers
        )

        if response.status_code != 201:
            print(f"Failed to upload, status {response.status_code}, file_path={file_path}")
            await self._file_upload_failed(turn_context, "Unable to upload file.")
        else:
            await self._file_upload_complete(turn_context, file_consent_card_response)

    async def on_teams_file_consent_decline(
            self,
            turn_context: TurnContext,
            file_consent_card_response: FileConsentCardResponse
    ):
        """
        The user declined the file upload.
        """

        context = file_consent_card_response.context

        reply = self._create_reply(
            turn_context.activity,
            f"Declined. We won't upload file <b>{context['filename']}</b>.",
            "xml"
        )
        await turn_context.send_activity(reply)

    async def _file_upload_complete(
            self,
            turn_context: TurnContext,
            file_consent_card_response: FileConsentCardResponse
    ):
        """
        The file was uploaded, so display a FileInfoCard so the user can view the
        file in Teams.
        """

        name = file_consent_card_response.upload_info.name

        download_card = FileInfoCard(
            unique_id=file_consent_card_response.upload_info.unique_id,
            file_type=file_consent_card_response.upload_info.file_type
        )

        as_attachment = Attachment(
            content=download_card.serialize(),
            content_type=ContentType.FILE_INFO_CARD,
            name=name,
            content_url=file_consent_card_response.upload_info.content_url
        )

        reply = self._create_reply(
            turn_context.activity,
            f"<b>Report uploaded to your OneDrive.</b> Your report <b>{name}</b> is ready to download",
            "xml"
        )
        reply.attachments = [as_attachment]

        await turn_context.send_activity(reply)
        reply = self._create_reply(
                turn_context.activity, f"Would you like to update report parameters or the options for this report? Then Please update the template and upload...", "xml"
            )
        await turn_context.send_activity(reply)

    async def _file_upload_failed(self, turn_context: TurnContext, error: str):
        reply = self._create_reply(
            turn_context.activity,
            f"<b>File upload failed.</b> Error: <pre>{error}</pre>",
            "xml"
        )
        await turn_context.send_activity(reply)

    def _create_reply(self, activity, text=None, text_format=None):
        return Activity(
            type=ActivityTypes.message,
            timestamp=datetime.utcnow(),
            from_property=ChannelAccount(
                id=activity.recipient.id, name=activity.recipient.name
            ),
            recipient=ChannelAccount(
                id=activity.from_property.id, name=activity.from_property.name
            ),
            reply_to_id=activity.id,
            service_url=activity.service_url,
            channel_id=activity.channel_id,
            conversation=ConversationAccount(
                is_group=activity.conversation.is_group,
                id=activity.conversation.id,
                name=activity.conversation.name,
            ),
            text=text or "",
            text_format=text_format or None,
            locale=activity.locale,
        )
        
    async def _message_all_members(self, turn_context: TurnContext):
        team_members = await self._get_paged_members(turn_context)

        for member in team_members:
            conversation_reference = TurnContext.get_conversation_reference(
                turn_context.activity
            )
            #await self._add_conversation_reference(turn_context)

            conversation_parameters = ConversationParameters(
                is_group=False,
                bot=turn_context.activity.recipient,
                members=[member],
                tenant_id=turn_context.activity.conversation.tenant_id,
            )

            async def get_ref(tc1):
                conversation_reference_inner = TurnContext.get_conversation_reference(
                    tc1.activity
                )
                return await tc1.adapter.continue_conversation(
                    conversation_reference_inner, send_message, self._app_id
                )

            async def send_message(tc2: TurnContext):
                reply = MessageFactory.list([])
                reply.attachments.append(self._send_suggested_actions_yes_no(member.name))

                return await tc2.send_activity(reply)
                

            await turn_context.adapter.create_conversation(
                conversation_reference, get_ref, conversation_parameters
            )
            #await self._add_conversation_reference(turn_context)
        await turn_context.send_activity(
            MessageFactory.text("All messages have been sent")
        )

    async def _get_paged_members(
        self, turn_context: TurnContext
    ) -> List[TeamsChannelAccount]:
        paged_members = []
        continuation_token = None

        while True:
            current_page = await TeamsInfo.get_paged_members(
                turn_context, continuation_token, 100
            )
            continuation_token = current_page.continuation_token
            paged_members.extend(current_page.members)

            if continuation_token is None:
                break

        return paged_members

    async def _add_conversation_reference(self, turn_context: TurnContext):
        conversation_reference = TurnContext.get_conversation_reference(turn_context.activity)
        if conversation_reference.conversation.conversation_type =='personal':
            self.conversation_references[
                conversation_reference.user.id
            ] = conversation_reference
        else:
            team_members = await self._get_paged_members(turn_context)
            for member in team_members:
                conversation_reference = TurnContext.get_conversation_reference(
                    turn_context.activity
                )
                # await self._add_conversation_reference(turn_context)

                conversation_parameters = ConversationParameters(
                    is_group=False,
                    bot=turn_context.activity.recipient,
                    members=[member],
                    tenant_id=turn_context.activity.conversation.tenant_id,
                )

                async def get_ref(tc1):
                    conversation_reference_inner = TurnContext.get_conversation_reference(
                        tc1.activity
                    )
                    return await tc1.adapter.continue_conversation(
                        conversation_reference_inner, send_message, self._app_id
                    )

                async def send_message(tc2: TurnContext):
                    reply = MessageFactory.list([])
                    reply.attachments.append(self._send_suggested_actions_yes_no(member.name))
                    new_conversation_reference = TurnContext.get_conversation_reference(tc2.activity)
                    new_conversation_reference.user=member
                    self.conversation_references[
                        new_conversation_reference.user.id
                    ] = new_conversation_reference
                    return await tc2.send_activity(reply)

                await turn_context.adapter.create_conversation(
                    conversation_reference, get_ref, conversation_parameters
                )

                # await self._add_conversation_reference(turn_context)
            await turn_context.send_activity(
                MessageFactory.text("All messages have been sent")
            )


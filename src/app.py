import base64
import os
from datetime import datetime, timezone, timedelta
from logging import Logger
from typing import List

import hydra
from loguru import logger

import requests
from msal import PublicClientApplication
from omegaconf import DictConfig

from src.models.attachment import Attachment
from src.models.message import Message, Body, EmailAddress
from src.inventory_generator import generate_inventory

ACCESS_TOKEN = ""
REFRESH_TOKEN = ""
OUTPUT_DIRECTORY = ""
SYSTEM_LOGGER: Logger = None
CONFIG: DictConfig = None


@hydra.main(version_base=None, config_path="./", config_name="config")
def load_config(config: DictConfig):
    global CONFIG
    CONFIG = config


def log_binder(name):
    def filter(record):
        return record["extra"].get("name") == name

    return filter


def configure_logging():
    logger.remove()
    logger.add("./logs/logs.log", level="DEBUG", filter=log_binder("system"))

    global SYSTEM_LOGGER
    SYSTEM_LOGGER = logger.bind(name="system")


def set_output_directory():
    global OUTPUT_DIRECTORY
    OUTPUT_DIRECTORY = os.path.join(os.getcwd(), "output")
    if not os.path.exists(OUTPUT_DIRECTORY):
        os.makedirs(OUTPUT_DIRECTORY)

    run_instance_directory = os.path.join(
        OUTPUT_DIRECTORY, datetime.now().strftime("%d-%m-%Y %H-%M-%S")
    )
    if not os.path.exists(run_instance_directory):
        os.makedirs(run_instance_directory)

    OUTPUT_DIRECTORY = run_instance_directory


def set_tokens(access_token, refresh_token):
    global ACCESS_TOKEN
    global REFRESH_TOKEN
    ACCESS_TOKEN = access_token
    REFRESH_TOKEN = refresh_token


def get_tokens():
    return ACCESS_TOKEN, REFRESH_TOKEN


def create_temp_directory():
    if not os.path.exists("./temp"):
        os.makedirs("./temp")
    else:
        li = os.listdir(os.path.join("temp"))
        for item in li:
            try:
                os.remove(os.path.join("temp", item))
            except:
                pass


def process_messages():
    messages: List[Message] = []
    before = timedelta(days=CONFIG.app_config.days)
    current_datetime = datetime.now(timezone.utc)
    received_after_datetime = (
        (current_datetime - before).isoformat().replace("+00:00", "Z")
    )
    list_messages_api = f"https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$filter=receivedDateTime ge {received_after_datetime}&$select=id,bodyPreview,body,hasAttachments,sender,from,toRecipients,receivedDateTime&$top=100"

    print(
        f"⏳ Downloading messages received on or after {(current_datetime - before).strftime('%d/%m/%Y %H:%M:%S')} (UTC)"
    )
    response = api_request(
        list_messages_api, {"Prefer": "outlook.body-content-type='text'"}
    )
    response = append_messages_or_retry(response, messages, list_messages_api)

    while "@odata.nextLink" in response:
        nextLink = response["@odata.nextLink"]
        response = api_request(nextLink, {"Prefer": "outlook.body-content-type='text'"})
        response = append_messages_or_retry(response, messages, nextLink)

    print(
        f"✅ Downloaded {len(messages)} {'messages' if len(messages) > 1 else 'message'}"
    )

    number_of_messages_with_attachments = 0
    if len(messages) > 0:
        messages_with_attachments = [
            message for message in messages if message.has_attachments
        ]
        number_of_messages_with_attachments = len(messages_with_attachments)
        if number_of_messages_with_attachments > 0:
            print(
                f"[INFO] {number_of_messages_with_attachments} {'messages' if number_of_messages_with_attachments > 1 else 'message'} contains attachments."
            )
            print("⏳ Downloading attachments")

    message_counter = 0
    for message in messages:
        message_counter += 1

        if message.has_attachments:
            list_attachments_api = (
                "https://graph.microsoft.com/v1.0/me/messages/"
                + message.id
                + "/attachments"
            )
            response = api_request(list_attachments_api)
            append_attachments_or_retry(response, message, list_attachments_api)

    if number_of_messages_with_attachments:
        print("✅ Attachments downloaded")

        number_of_messages_with_excel_files = 0
        for message in messages:
            if message.has_attachments:
                for attachment in message.attachments:
                    if attachment.name.endswith(".xls") or attachment.name.endswith(
                        ".xlsx"
                    ):
                        number_of_messages_with_excel_files += 1
                        message.has_excel_files = True
                        break

        print(
            f"[INFO] {number_of_messages_with_excel_files} {'messages' if number_of_messages_with_excel_files > 1 else 'message'} contains Excel files in attachments."
        )

    return messages


def append_messages(response, messages: List[Message]):
    response_messages = response["value"]
    if len(response_messages) > 0:
        for message in response_messages:
            body = Body(message["body"]["content"], message["body"]["contentType"])
            sender = EmailAddress(
                message["sender"]["emailAddress"]["name"],
                message["sender"]["emailAddress"]["address"],
            )
            from_sender = EmailAddress(
                message["from"]["emailAddress"]["name"],
                message["from"]["emailAddress"]["address"],
            )
            to_recipients: List[EmailAddress] = []
            for recipient in message["toRecipients"]:
                to_recipients.append(
                    EmailAddress(
                        recipient["emailAddress"]["name"],
                        recipient["emailAddress"]["address"],
                    )
                )

            received_at = str(message["receivedDateTime"])
            received_at_date = received_at.split("T")[0]
            received_at_time = (
                received_at.split("T")[1].replace("Z", "").replace("z", "")
            )

            messages.append(
                Message(
                    message["id"],
                    message["bodyPreview"],
                    body,
                    sender,
                    from_sender,
                    to_recipients,
                    message["hasAttachments"],
                    f"{received_at_date} {received_at_time}",
                )
            )


def append_messages_or_retry(response, messages: List[Message], url: str):
    try:
        append_messages(response, messages)
    except:
        if (
            "error" in response
            and "code" in response["error"]
            and response["error"]["code"] == "InvalidAuthenticationToken"
        ):
            result = app.acquire_token_silent(
                scopes=SCOPES, account=app.get_accounts()[0]
            )
            set_tokens(result["access_token"], result["refresh_token"])
            response = api_request(url, {"Prefer": "outlook.body-content-type='text'"})
            append_messages(response, messages)
        else:
            raise
    return response


def append_attachments(response, message: Message):
    message.attachments = []
    for attachment in response["value"]:
        if attachment["@odata.type"] == "#microsoft.graph.fileAttachment":
            content = Attachment(
                attachment["name"], base64.b64decode(attachment["contentBytes"])
            )
            message.attachments.append(content)


def append_attachments_or_retry(response, message: Message, url: str):
    try:
        append_attachments(response, message)
    except:
        if (
            "error" in response
            and "code" in response["error"]
            and response["error"]["code"] == "InvalidAuthenticationToken"
        ):
            result = app.acquire_token_silent(
                scopes=SCOPES, account=app.get_accounts()[0]
            )
            set_tokens(result["access_token"], result["refresh_token"])
            response = api_request(url)
            append_attachments(response, message)
        else:
            raise


def api_request(url: str, header_list: dict[str, str] = dict()):
    headers = {"Authorization": "Bearer " + ACCESS_TOKEN}
    for header in header_list.keys():
        headers[header] = header_list.get(header)
    response = requests.get(url=url, headers=headers)
    response_json = response.json()
    return response_json


def save_messages(messages: List[Message]):
    if not os.path.exists(os.path.join(OUTPUT_DIRECTORY, "Messages")):
        os.makedirs(os.path.join(OUTPUT_DIRECTORY, "Messages"))

    message_counter = 0
    for message in messages:
        if message.has_excel_files:
            message_counter += 1
            message_directory = os.path.join(
                OUTPUT_DIRECTORY, "Messages", "Message " + str(message_counter)
            )
            if not os.path.exists(message_directory):
                os.makedirs(message_directory)
            file_path = os.path.join(
                message_directory,
                "message" + (".html" if message.body.contentType == "html" else ".txt"),
            )
            if os.path.exists(file_path):
                os.remove(file_path)
            file = open(file_path, "w", encoding="utf-8")
            file.write(str(message.body.content))
            file.close()

            if message.has_attachments:
                attachments_directory = os.path.join(message_directory, "Attachments")
                if not os.path.exists(attachments_directory):
                    os.makedirs(attachments_directory)

                for attachment in message.attachments:
                    if attachment.name.lower().endswith(
                        ".xlsx"
                    ) or attachment.name.lower().endswith(".xls"):
                        attachment_path = os.path.join(
                            attachments_directory, attachment.name
                        )
                        if os.path.exists(attachment_path):
                            os.remove(attachment_path)
                        file = open(attachment_path, "w+b")
                        file.write(attachment.content)
                        file.close()


if __name__ == "__main__":
    try:
        load_config()
        configure_logging()
        set_output_directory()

        if "secrets" not in CONFIG or (
            "client_id" not in CONFIG.secrets
            or "openai_api_key" not in CONFIG.secrets
            or "price_runner_token" not in CONFIG.secrets
        ):
            if "secrets" not in CONFIG or "client_id" not in CONFIG.secrets:
                print("⚠️ Secret 'client_id' not found")
            if "secrets" not in CONFIG or "openai_api_key" not in CONFIG.secrets:
                print("⚠️ Secret 'openai_api_key' not found")
            if "secrets" not in CONFIG or "price_runner_token" not in CONFIG.secrets:
                print("⚠️ Secret 'price_runner_token' not found")
            print(
                "\nPlease make sure you have a file named 'config.yaml' in the root directory of this application and have placed your secrets as:\n\n"
                "config.yaml\n"
                "-----------\n"
                "client_id: <YOUR_CLIENT_ID>\n"
                "openai_api_key: <YOUR_OPENAI_API_KEY>\n"
                "price_runner_token: <YOUR_PRICE_RUNNER_TOKEN>\n\n"
                "INFO:\n"
                "client_id: Your Microsoft Entra registered app's Client ID\n"
                "openai_api_key: Your OpenAI's API Key\n"
                "price_runner_token: Your PriceRunner's Product Token\n"
            )
        else:
            app = PublicClientApplication(
                CONFIG.secrets.client_id,
                authority="https://login.microsoftonline.com/common/",
            )
            SCOPES = ["Mail.Read"]
            flow = app.initiate_device_flow(scopes=SCOPES)
            print(f"🚀 {flow['message']}")
            result = app.acquire_token_by_device_flow(flow)

            if "access_token" in result:
                set_tokens(result["access_token"], result["refresh_token"])
                create_temp_directory()

                messages = process_messages()
                if len(messages) > 0:
                    save_messages(messages)
                    generate_inventory(
                        OUTPUT_DIRECTORY,
                        [message for message in messages if message.has_excel_files],
                        SYSTEM_LOGGER,
                        CONFIG,
                    )
                print("✅ Done")
            else:
                print(
                    "⚠️ Microsoft communication failure. Please try re-running the application."
                )
            input()
    except Exception as ex:
        SYSTEM_LOGGER.debug(ex)
        print("Exception occurred. Check logs for details.")
        input()

import base64
import datetime
import os.path
from typing import List

from src.app import (
    set_tokens,
    get_tokens,
    api_request,
    append_messages,
    append_attachments,
    create_temp_directory,
    configure_logging,
)
from src.models.message import Message, Body, EmailAddress


def test_set_tokens():
    set_tokens("access_token", "refresh_token")
    assert get_tokens() == ("access_token", "refresh_token")


def test_get_tokens():
    assert get_tokens() == ("access_token", "refresh_token")


def test_create_temp_directory():
    configure_logging()
    create_temp_directory()
    assert os.path.exists("./temp")


def test_api_request():
    configure_logging()
    code = "InvalidAuthenticationToken"
    response = api_request("https://graph.microsoft.com/v1.0/me/messages")
    assert response["error"]["code"] == code


def test_append_messages():
    configure_logging()
    response = {
        "value": [
            {
                "id": "1",
                "bodyPreview": "Test preview",
                "body": {"content": "<body>Test</body>", "contentType": "html"},
                "sender": {
                    "emailAddress": {"name": "Tester", "address": "tester@gmail.com"}
                },
                "from": {
                    "emailAddress": {"name": "Tester", "address": "tester@gmail.com"}
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "name": "Recipient",
                            "address": "recipient@gmail.com",
                        }
                    }
                ],
                "hasAttachments": False,
                "receivedDateTime": datetime.datetime.now().isoformat(),
            }
        ]
    }
    messages: List[Message] = []
    append_messages(response, messages)
    assert len(messages) == 1


def test_append_attachments():
    configure_logging()
    response = {
        "value": [
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": "test.xlsx",
                "contentBytes": base64.b64encode("Test Content".encode("utf-8")),
            }
        ]
    }
    message: Message = Message(
        "1",
        "Test",
        Body(
            "Athlone Dental items comes under sale with the launch of new tooth brushes.",
            "html",
        ),
        EmailAddress("Tester", "test@gmail.com"),
        EmailAddress("Tester", "test@gmail.com"),
        [EmailAddress("Tester", "test@gmail.com")],
        True,
        datetime.datetime.now().isoformat(),
    )
    append_attachments(response, message)
    assert len(message.attachments) == 1


# def test_process_messages():
#     set_tokens(
#         "EwB4A8l6BAAUbDba3x2OMJElkF7gJ4z/VbCPEz0AAXdihoLYfO/gFTFSEZxvdMDAZZhDqD1Mci2S6IL6DUBjP+dPE8yC057Gdz8d6WScTz2KNR7RHzjqWnGlDnO6MfWa+1+KyuR0yBS0qMHYe+BjOzk8mwzWFq6QMQ2LpLcJF9SZh20aSL3DQku30Xa4WAyH/WzmWIsox6yuxxlSvZngTHrKD+n+bBOXVO9H2Bv94PL6CzySHiixWWqT3Ck0eGAi8QED/IctiYQkf1u/L5aKovm111N+8z9FPTkHmw8Ka/INtUrGKqsIYhkDXLTNd/K3NLGFC4c42headZHqU0+SteTd0P3b9LyUTSOSYlgZJV7+71GKkLqbKbtFfM1RcE4DZgAACKO0fXaDZgWhSAJ5BFYXafjUCCPfpSExBSus23G1Z+m6VTPfpwphZKeZfgw15deqI8w043b6oP7D3nD1Z4vfYFD+KSoEj4SrUGzTyzKqhbGWLzV98vWeg8GWLeWvaiPHgmnG9blr28iS/G8aG+YNMfdPe/nzNGpmgly520A0uUDKpJV5vaGEjT8ySbyyUBnRfD2yoJKxKJBsTeqxZytBepxPVLD9u3BfNe88VOhyOi57m4QvG0r/Ro9mpo1xoKK5BDL7koHxf4FfgIwNMxNxX5cHnju0R8flwdp28tXtZ8p2qEQLeAywyY8ckX4JQgnnbc7RcQqTi/lSqV6wstv8Mjjy/7S9fR5/hteWLAjVbHUbjlRpqUMtevxWZ0bwUk/tDjdEaqo1GhNUPoFoiS5dx5jkMBbRlvgvoQaKtaYs9YPsX9xLtJnKiM/25Yir3gZUinOXVEsP49MK1kVDsoeDZZ3vEroSpseJ95r/k8mv+zR1/jRQa+N35sfffJatwl57wov2REYFUy783aosTYpgXVncG2MEicyutbqjfy0+nm+hTrkT5W2fFlShUhw4cP/vF/S0D7zRTlI5/C4jrS6wmJ3WTvI7ePe5nKXMP0xT7KfUcTdgDB80aTYcX99gugWiHjej+TgIG0ceyrHCwqVtLyZXkVyFnnAmEoxMb92PVeBeOViiP7CK79J2HO+aGZZnOtqi3nWTQDWGmF59HYI7n8CBdF1S8fuXEB9PWsSh8Zqems9MyDAoqXqwT8rfJAJ0BnXvqu97nwXn3cMVechDE7kqyIsC",
#         "M.C508_BAY.0.U.-CvhZ9vnGYfOJApm5DuZdG9kNkRC1DeIxY68HlOgT03OtEcdQft!whEpHcPbHm9dd0oT98UDRma4CGK!i7iejmlGZLEQ6Myj78NfZ1vxmyxJSk5GC6P*5U*UaoKboqeUWnNROFnaSlc6k*39To!*AJfs8bu9uQRrw64Ex!KStY9jcLsMvsyG6DDRIN8ql3l0s9IyfVhO0gYL31*Ojqag87exiNp0l4r5Ofj9m4q9houBtdpQnuKyJiLWa6nB7suSeOMUjczbbWN4e56BSYrZgcDZmhpq!zPhOOHi7YakqsRjkEOFf*vvl6PXt1PbrwgYze9I7hsxFviTGa7eG3T!nittqORFKQt7WNRDHjnn1sZJc",
#     )
#     assert len(process_messages()) > 0


# def test_generate_inventory_model():
#     logger.remove()
#     logger.add(lambda msg: None, level="DEBUG")
#     logger.add(f"./logs/log.log", rotation="50 MB", level="DEBUG")
#
#     message: Message = Message(
#         "1",
#         "Test",
#         Body(
#             "Athlone Dental items comes under sale with the launch of new tooth brushes.",
#             "html",
#         ),
#         EmailAddress("Tester", "test@gmail.com"),
#         EmailAddress("Tester", "test@gmail.com"),
#         [EmailAddress("Tester", "test@gmail.com")],
#     )
#     attachment: Attachment = Attachment("test-attachment.xlsx", None)
#     attachment.path = "./test-attachment.xlsx"
#     message.attachments = [attachment]
#
#     if not os.path.exists("./test-outputs"):
#         os.makedirs(os.path.join("test-outputs"))
#
#     generate_inventory_model("./test-outputs", [message], logger)
#
#     assert os.path.exists("./test-outputs/Report.xlsx")
#
#
# def test_generate_excel():
#     item: Item = Item("Pepsodent", 10, 12.75, 127.5)
#     inventory: Inventory = Inventory("Dental", "Tooth brushes")
#     inventory.sender = "Tester - tester@gmail.com"
#     inventory.items = [item]
#     inventory_list: List[Inventory] = [inventory]
#
#     logger.remove()
#     logger.add(lambda msg: None, level="DEBUG")
#     logger.add(f"./logs/log.log", rotation="50 MB", level="DEBUG")
#
#     if not os.path.exists("./test-outputs"):
#         os.makedirs(os.path.join("test-outputs"))
#
#     generate_excel(inventory_list, "./test-outputs", logger)
#
#     assert os.path.exists("./test-outputs/Report.xlsx")

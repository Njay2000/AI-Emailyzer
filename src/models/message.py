from typing import List

from src.models.attachment import Attachment


class Body:
    def __init__(self, content, content_type):
        self.contentType = content_type
        self.content = content


class Sender:
    def __init__(self, email_address):
        self.emailAddress = email_address


class From:
    def __init__(self, email_address):
        self.emailAddress = email_address


class EmailAddress:
    def __init__(self, name, address):
        self.name = name
        self.address = address


class ToRecipients:
    def __init__(self, email_address: List[EmailAddress]):
        self.emailAddress = email_address


class Message:
    attachments: List[Attachment]
    has_excel_files: bool = False

    def __init__(
        self,
        id: str,
        body_preview: str,
        body: Body,
        sender: EmailAddress,
        from_sender: EmailAddress,
        to_recipients: List[EmailAddress],
        has_attachments: bool,
        received_at: str,
    ):
        self.id = id
        self.bodyPreview = body_preview
        self.body: Body = body
        self.sender: EmailAddress = sender
        self.fromSender: EmailAddress = from_sender
        self.toRecipients: List[EmailAddress] = to_recipients
        self.has_attachments = has_attachments
        self.attachments = []
        self.received_at = received_at

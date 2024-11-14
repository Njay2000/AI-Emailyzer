class ReportSheetItem:

    def __init__(
        self,
        sender: str,
        received_at: str,
        sheet_name: str,
        status: str,
        comments: str,
        file_name: str,
        file_path: str,
    ):
        self.sender: str = sender
        self.received_at: str = received_at
        self.sheet_name: str = sheet_name
        self.status: str = status
        self.comments: str = comments
        self.file_name: str = file_name
        self.file_path: str = file_path

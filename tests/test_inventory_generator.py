# import os
# from datetime import datetime
# from typing import List
#
# from hydra import initialize, compose
# from loguru import logger
# from omegaconf import DictConfig
#
# from src.app import append_messages
# from src.inventory_generator import generate_inventory_model
# from src.models.message import Message
#
# CONFIG: DictConfig = None
#
#
# def log_binder(name):
#     def filter(record):
#         return record["extra"].get("name") == name
#
#     return filter
#
#
# def test_generate_inventory_model():
#     with initialize(version_base=None, config_path="../"):
#         config = compose(config_name="config")
#         output_directory = "./test-output"
#
#         logger.remove()
#         logger.add(
#             "./tests/test-logs/system.log", level="DEBUG", filter=log_binder("system")
#         )
#         logger.add(
#             "./tests/test-logs/info.log", level="INFO", filter=log_binder("info")
#         )
#
#         system_logger = logger.bind(name="system")
#         user_logger = logger.bind(name="info")
#
#         response = {
#             "value": [
#                 {
#                     "id": "1",
#                     "bodyPreview": "Test preview",
#                     "body": {"content": "<body>Test</body>", "contentType": "html"},
#                     "sender": {
#                         "emailAddress": {
#                             "name": "Tester",
#                             "address": "tester@gmail.com",
#                         }
#                     },
#                     "from": {
#                         "emailAddress": {
#                             "name": "Tester",
#                             "address": "tester@gmail.com",
#                         }
#                     },
#                     "toRecipients": [
#                         {
#                             "emailAddress": {
#                                 "name": "Recipient",
#                                 "address": "recipient@gmail.com",
#                             }
#                         }
#                     ],
#                     "hasAttachments": False,
#                     "receivedDateTime": datetime.now().isoformat(),
#                 }
#             ]
#         }
#         messages: List[Message] = []
#         append_messages(response, messages)
#
#         generate_inventory_model(
#             output_directory,
#             messages,
#             config,
#             system_logger,
#             user_logger,
#         )
#
#         resources = os.listdir("./tests/test-output")
#         assert len(resources) >= 0

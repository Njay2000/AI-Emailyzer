# from loguru import logger
#
# from src.report_generator import generate_excel
# from src.models.inventory import Inventory
# from src.models.item import Item
# from src.models.real_time_pricing_details import RealTimePricingDetails
#
#
# def log_binder(name):
#     def filter(record):
#         return record["extra"].get("name") == name
#
#     return filter
#
#
# def test_generate_excel():
#     inventory: Inventory = Inventory("Centra")
#     item: Item = Item("00000123456789", "Sample Product", 10, 10.5)
#     item.real_time_pricing_details = RealTimePricingDetails()
#     item.real_time_pricing_details.median = 10.5
#     item.real_time_pricing_details.lowest_price = 2.0
#     item.real_time_pricing_details.highest_price = 20
#     item.real_time_pricing_details.lowest_price_retailer = "Lowest price retailer"
#     item.real_time_pricing_details.highest_price_retailer = "Highest price retailer"
#     item.real_time_pricing_details.average_price = 5.5
#     items: list[Item] = [item]
#     inventory.items = items
#     inventory.product_category = "Sample Category"
#     inventory.sender = "Tester - tester@gmail.com"
#     inventory_list: list[Inventory] = [inventory]
#
#     logger.remove()
#     logger.add(
#         "./tests/test-logs/system.log", level="DEBUG", filter=log_binder("system")
#     )
#
#     system_logger = logger.bind(name="system")
#
#     generate_excel(inventory_list, "./tests/test-output", system_logger)

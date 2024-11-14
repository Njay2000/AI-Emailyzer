from src.models.real_time_pricing_details import RealTimePricingDetails


class ConsolidatedSheetItem:

    def __init__(
        self, sender: str, barcode, quantity, product_description: str, unit_price
    ):
        self.sender: str = sender
        self.barcode = barcode
        self.quantity = quantity
        self.product_description: str = product_description
        self.unit_price = unit_price
        self.price_runner_details: RealTimePricingDetails = RealTimePricingDetails()

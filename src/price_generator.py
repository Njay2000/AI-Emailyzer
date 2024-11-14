import json
from typing import List

import requests
import numpy as ny
from omegaconf import DictConfig

from src.models.consolidated_sheet import ConsolidatedSheetItem


def fetch_prices(
    items: List[ConsolidatedSheetItem],
    price_runner_token: str,
    system_logger: "loguru.system_logger",
):
    params = []
    try:
        for item in items:
            barcode = str(item.barcode).strip().replace("-", "")
            params.append(
                (
                    "gtin14s",
                    (
                        (((14 - len(barcode)) * "0") + barcode)
                        if len(barcode) < 14
                        else barcode
                    ),
                )
            )

        response = requests.get(
            url="https://api.pricerunner.com/public/v2/product/offers/UK/gtin14s",
            params=params,
            headers={"tokenId": price_runner_token},
        )

        if response.status_code == 200:
            try:
                listings = json.loads(response.content)["productListings"]

                passed_id_list = [param[1] for param in params]
                for listing in listings:
                    prices: list[float] = []
                    lowest_price = 9999999999
                    lowest_price_retailer = ""
                    highest_price = -1.0
                    highest_price_retailer = ""
                    for offer in listing["offers"]:
                        price = 0.0
                        price += float(offer["price"]["value"])
                        if offer["shippingCost"] is not None:
                            price += float(offer["shippingCost"]["value"])
                        prices.append(price)
                        if price < lowest_price:
                            lowest_price = price
                            lowest_price_retailer = offer["merchantName"]
                        if price > highest_price:
                            highest_price = price
                            highest_price_retailer = offer["merchantName"]

                    item: ConsolidatedSheetItem = None
                    for passed_id in passed_id_list:
                        if passed_id in listing["productListingProduct"]["gtin14s"]:
                            item = items[passed_id_list.index(passed_id)]

                            if item is not None:
                                if len(prices) > 0:
                                    prices.sort()
                                    item.price_runner_details.median = round(
                                        float(ny.median(prices)), 2
                                    )
                                    item.price_runner_details.lowest_price = (
                                        lowest_price
                                    )
                                    item.price_runner_details.lowest_price_retailer = (
                                        lowest_price_retailer
                                    )
                                    item.price_runner_details.highest_price = (
                                        highest_price
                                    )
                                    item.price_runner_details.highest_price_retailer = (
                                        highest_price_retailer
                                    )
                                    item.price_runner_details.average_price = round(
                                        float(ny.mean(prices)), 2
                                    )

            except Exception as ex:
                system_logger.debug(ex)
        else:
            try:
                system_logger.debug(
                    str(response.status_code) + " - " + str(response.content)
                )
            except:
                pass
    except Exception as ex:
        system_logger.debug(ex)


def get_price_results(
    items: List[ConsolidatedSheetItem],
    config: DictConfig,
    system_logger: "loguru.system_logger",
):
    print("⏳ Fetching PriceRunner details")
    skip = 0
    take = 100
    sliced_results = items[skip : skip + take]
    fetch_prices(sliced_results, str(config.secrets.price_runner_token), system_logger)

    while len(items) > len(items[: skip + take]):
        skip += 100
        sliced_results = items[skip : skip + take]
        fetch_prices(
            sliced_results, str(config.secrets.price_runner_token), system_logger
        )

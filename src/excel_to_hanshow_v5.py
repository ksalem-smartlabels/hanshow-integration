#!/usr/bin/env python3

import requests
import time
import base64
import pandas as pd
import sys
import os
import shutil
import logging
from logging.handlers import TimedRotatingFileHandler

# ======================================================
# CONFIG
# ======================================================

WATCH_FOLDER = r"C:\wamp64\www\csv\INTERFACE_FILE"
DONE_FOLDER = os.path.join(WATCH_FOLDER, "done")

HANSHOW_BASE_URL = "https://eastwest-integration.slscanada.ca"
CUSTOMER_CODE = "eastwest"
STORE_CODE = "teststore"

HANSHOW_BASIC_USER = "ot8mQ)Voz)T9puEzaQjcy8"
HANSHOW_BASIC_PASSWORD = "(zi8wJdFxXSpcQ54YSRE9X"

REQUIRED_COLUMNS = [
    "ItemID",
    "ItemName",
    "ItemNumber",
    "PrimaryUpc",
    "UnitPrice"
]

LOG_FILE = "hanshow_integration.log"

BATCH_SIZE = 1000
SLEEP_SECONDS = 300  # 5 minutes
RETENTION_DAYS = 30

# ======================================================
# LOGGING SETUP
# ======================================================

logger = logging.getLogger("HanshowIntegration")
logger.setLevel(logging.INFO)

file_handler = TimedRotatingFileHandler(
    LOG_FILE,
    when="midnight",
    interval=1,
    backupCount=RETENTION_DAYS
)

formatter = logging.Formatter(
    "%(asctime)s - %(levelname)s - %(message)s"
)

file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

console_handler = logging.StreamHandler()
console_handler.setFormatter(formatter)
logger.addHandler(console_handler)

# ======================================================
# READ EXCEL
# ======================================================

def read_excel_file(file_path):
    logger.info(f"Reading Excel file: {file_path}")

    df = pd.read_excel(file_path)

    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            raise Exception(f"Missing required column: {col}")

    logger.info(f"Excel loaded. Rows found: {len(df)}")
    return df


# ======================================================
# BUILD PAYLOAD
# ======================================================

def build_hanshow_items(df):
    items = []

    for _, row in df.iterrows():

        if pd.isna(row["ItemID"]):
            continue

        sku = str(row["ItemID"]).strip()
        name = str(row["ItemName"]).strip()

        item_number = row["ItemNumber"]
        primary_upc = row["PrimaryUpc"]

        item_number_str = str(item_number).strip() if pd.notna(item_number) else ""
        primary_upc_str = str(primary_upc).strip() if pd.notna(primary_upc) else ""

        if item_number_str and primary_upc_str:
            ean_value = f"{item_number_str},{primary_upc_str}"
        elif item_number_str:
            ean_value = item_number_str
        elif primary_upc_str:
            ean_value = primary_upc_str
        else:
            ean_value = ""

        item = {
            "IIS_COMMAND": "COMPLETE_UPDATE",
            "sku": sku,
            "itemName": name,
            "ean": ean_value,
            "price1": float(row["UnitPrice"]) if not pd.isna(row["UnitPrice"]) else 0.0
        }

        items.append(item)

    logger.info(f"Items prepared: {len(items)}")
    return items


# ======================================================
# AUTH
# ======================================================

def hanshow_get_token():
    basic = f"{HANSHOW_BASIC_USER}:{HANSHOW_BASIC_PASSWORD}"
    basic_b64 = base64.b64encode(basic.encode()).decode()

    headers = {
        "Authorization": f"Basic {basic_b64}"
    }

    resp = requests.post(
        f"{HANSHOW_BASE_URL}/proxy/token",
        headers=headers,
        timeout=15
    )

    resp.raise_for_status()
    logger.info("Token obtained successfully")

    return resp.json()["access_token"]


# ======================================================
# PUSH BATCH
# ======================================================

def push_batch(token, batch, batch_number):

    payload = {
        "customerStoreCode": CUSTOMER_CODE,
        "storeCode": STORE_CODE,
        "batchNo": time.strftime("%Y%m%d%H%M%S"),
        "items": batch
    }

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    logger.info(f"Sending batch {batch_number} ({len(batch)} items)")

    resp = requests.post(
        f"{HANSHOW_BASE_URL}/proxy/integration/{CUSTOMER_CODE}/{STORE_CODE}",
        json=payload,
        headers=headers,
        timeout=60
    )

    logger.info(f"Batch {batch_number} HTTP Status: {resp.status_code}")

    if resp.status_code != 200:
        logger.error(f"Batch {batch_number} failed: {resp.text}")

    resp.raise_for_status()


# ======================================================
# CLEANUP OLD FILES
# ======================================================

def cleanup_old_done_files():

    if not os.path.exists(DONE_FOLDER):
        return

    now = time.time()
    cutoff = now - (RETENTION_DAYS * 24 * 60 * 60)

    for filename in os.listdir(DONE_FOLDER):

        file_path = os.path.join(DONE_FOLDER, filename)

        if not os.path.isfile(file_path):
            continue

        file_mtime = os.path.getmtime(file_path)

        if file_mtime < cutoff:
            try:
                os.remove(file_path)
                logger.info(f"Deleted old processed file: {filename}")
            except Exception as e:
                logger.error(f"Failed to delete old file {filename}: {e}")


# ======================================================
# PROCESS FILE
# ======================================================

def process_file(file_path):

    df = read_excel_file(file_path)
    items = build_hanshow_items(df)

    token = hanshow_get_token()

    total = len(items)
    logger.info(f"Total items to send: {total}")

    batch_number = 1

    for i in range(0, total, BATCH_SIZE):
        batch = items[i:i + BATCH_SIZE]
        push_batch(token, batch, batch_number)
        batch_number += 1

    os.makedirs(DONE_FOLDER, exist_ok=True)

    destination = os.path.join(DONE_FOLDER, os.path.basename(file_path))
    shutil.move(file_path, destination)

    logger.info(f"File moved to: {destination}")


# ======================================================
# WATCH LOOP
# ======================================================

def watch_folder():

    logger.info(f"Watching folder: {WATCH_FOLDER}")

    os.makedirs(WATCH_FOLDER, exist_ok=True)

    while True:

        files = [
            f for f in os.listdir(WATCH_FOLDER)
            if f.lower().endswith(".xlsx")
        ]

        for file in files:
            full_path = os.path.join(WATCH_FOLDER, file)

            try:
                process_file(full_path)
            except Exception as e:
                logger.error(f"Error processing file {file}: {e}")

        cleanup_old_done_files()

        logger.info("Sleeping 5 minutes...")
        time.sleep(SLEEP_SECONDS)


# ======================================================
# ENTRY
# ======================================================

if __name__ == "__main__":
    watch_folder()

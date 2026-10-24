"""
WhatsApp Mess Menu Bot — Strictly for Your "Mess Menu" WhatsApp Group (v3)
==========================================================================
Built for BITSoM - Kamba Weekly Mess Menu:
• Pre-loaded with this week's menu (19/10/26 to 25/10/26) so it works immediately!
• Connects to your WhatsApp via QR Code ("Linked Devices") — 100% Free.
• STRICTLY & EXCLUSIVELY operates inside the WhatsApp group named "Mess Menu"
  (never messages your personal "You" chat or any other chat/group).
• Automatically posts to the "Mess Menu" group in IST:
    - 06:00 AM : Breakfast
    - 11:00 AM : Lunch (including Bowled Over & Food Trotters)
    - 03:00 PM : Evening Snacks (including Jain Snack & Hot Beverage)
    - 07:00 PM : Dinner (including Perfect Jodi / Bowled Over)
    - 09:00 AM (Mondays) : Reminder to drop the new week's Excel (.xlsx) file
• On Mondays, just drop the new `.xlsx` file into the "Mess Menu" group to update!
• Anyone in the "Mess Menu" group can text anytime:
    - "menu" or "today" -> Full day's menu (all 4 meals)
    - "breakfast", "lunch", "snacks", or "dinner" -> Specific meal for today
    - "next" -> Next upcoming meal
    - "tomorrow" or any weekday (e.g. "thursday", "friday lunch", "sunday dinner")
"""

import argparse
import base64
import html
import json
import logging
import os
import shutil
import subprocess
import threading
import time
from datetime import datetime, date
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from zoneinfo import ZoneInfo

import openpyxl

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
TIMEZONE = ZoneInfo("Asia/Kolkata")
MENU_FILE_PATH = os.environ.get("MENU_FILE_PATH", "current_menu.xlsx")
PORT = int(os.environ.get("PORT", 10000))
TARGET_GROUP_NAME = "Mess Menu"

# Daily notification schedule (24-hour format in IST)
MEAL_SCHEDULE = {
    "breakfast": "06:00",  # 6:00 AM IST
    "lunch":     "11:00",  # 11:00 AM IST
    "snacks":    "15:00",  # 3:00 PM IST
    "dinner":    "19:00",  # 7:00 PM IST
}

# Monday reminder time (24-hour format in IST) if the new week's Excel hasn't been uploaded
MONDAY_REMINDER_TIME = "09:00"

# Include all sections & dietary items
INCLUDE_NON_VEG = True
INCLUDE_JAIN = True
INCLUDE_FOOD_TROTTERS_IN_LUNCH = True
INCLUDE_PERFECT_JODI_IN_DINNER = True
INCLUDE_BOWLED_OVER_IN_LUNCH = True

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")
logger = logging.getLogger(__name__)

DAYS_ORDER = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"]
DAY_ALIASES = {
    "mon": "MONDAY", "monday": "MONDAY",
    "tue": "TUESDAY", "tues": "TUESDAY", "tuesday": "TUESDAY",
    "wed": "WEDNESDAY", "wednesday": "WEDNESDAY",
    "thu": "THURSDAY", "thur": "THURSDAY", "thurs": "THURSDAY", "thursday": "THURSDAY",
    "fri": "FRIDAY", "friday": "FRIDAY",
    "sat": "SATURDAY", "saturday": "SATURDAY",
    "sun": "SUNDAY", "sunday": "SUNDAY",
}

MEAL_ALIASES = {
    "breakfast": "breakfast", "bf": "breakfast", "morning": "breakfast",
    "lunch": "lunch",
    "snacks": "snacks", "snack": "snacks", "evening": "snacks",
    "dinner": "dinner", "night": "dinner",
}

SECTION_HEADERS = {
    "breakfast": "breakfast",
    "lunch": "lunch",
    "bowled over": "lunch_bowled_over",
    "food trotters": "lunch_food_trotters",
    "snacks": "snacks",
    "dinner": "dinner",
    "perfect jodi / bowled over": "dinner_perfect_jodi",
}

# Pre-loaded menu from your attached BITSoM - Kamba (19/10/26 to 25/10/26) Excel file
DEFAULT_MENU_DATA = json.loads(r"""{
  "dates": {
    "MONDAY": "19/10/26",
    "TUESDAY": "20/10/26",
    "WEDNESDAY": "21/10/26",
    "THURSDAY": "22/10/26",
    "FRIDAY": "23/10/26",
    "SATURDAY": "24/10/26",
    "SUNDAY": "25/10/26"
  },
  "day_labels": {
    "MONDAY": "MONDAY",
    "TUESDAY": "TUESDAY",
    "WEDNESDAY": "WEDNESDAY",
    "THURSDAY": "THURSDAY(gujrati day)",
    "FRIDAY": "FRIDAY",
    "SATURDAY": "SATURDAY",
    "SUNDAY": "SUNDAY"
  },
  "meals": {
    "MONDAY": {
      "breakfast": [
        [
          "Vegetarian Mains",
          "Steamed Idli with Sambhar"
        ],
        [
          "Fruits",
          "Cut Fruits"
        ],
        [
          "Cereals",
          "Wheat Flakes with Milk"
        ],
        [
          "Egg Preparation (2 Nos Max)",
          "Masala omelette"
        ],
        [
          "Hot Beverage",
          "Tea And Coffee"
        ]
      ],
      "lunch": [
        [
          "Salad",
          "Dill Cucumber salad"
        ],
        [
          "Dry Veg",
          "Tawa veg"
        ],
        [
          "Semi Dry Veg",
          "Amritsari Chole"
        ],
        [
          "Dal & Kadhi",
          "punjabi kadhi"
        ],
        [
          "Rice",
          "Steamed Rice"
        ],
        [
          "Non-Veg",
          "Chicken Lababdar"
        ],
        [
          "Indian Breads",
          "Phulka"
        ],
        [
          "Ancillaries",
          "Papad/Pickle/curd"
        ],
        [
          "Dessert",
          "rice coconut kheer"
        ],
        [
          "Bowled Over Lunch",
          "Bowled Over 1"
        ]
      ],
      "food_trotters": [
        "veg penne primavera",
        "Garlic buns",
        "Chicken Penne Primavera"
      ],
      "snacks": [
        [
          "Evening Snacks",
          "Ragda Pattice"
        ],
        [
          "Jain Snack",
          "Ragda Pattice"
        ],
        [
          "Hot Beverage",
          "Hot Beverage"
        ]
      ],
      "dinner": [
        [
          "Salad",
          "Greek Salad"
        ],
        [
          "Semi - Dry",
          "Paneer Mutter Masala"
        ],
        [
          "Dry Prep",
          "Aloo jeera"
        ],
        [
          "Rice",
          "Steamed Rice"
        ],
        [
          "Dal (Lentils)",
          "Dal Tadka"
        ],
        [
          "Non Veg",
          "Egg Punjabi Curry"
        ],
        [
          "Indian Breads",
          "Phulka"
        ],
        [
          "Ancillaries",
          "Papad/Pickle/mocktail"
        ],
        [
          "Dessert",
          "Chocolate brownie"
        ]
      ],
      "perfect_jodi": [
        "veg tawa pulao",
        "cucumber raita",
        "chicken tawa pulao"
      ]
    },
    "TUESDAY": {
      "breakfast": [
        [
          "Vegetarian Mains",
          "black chana sukha masal with puri and suji halwa"
        ],
        [
          "Fruits",
          "Cut Muskmelon + Watermelon"
        ],
        [
          "Cereals",
          "Muesili with milk"
        ],
        [
          "Egg Preparation (2 Nos Max)",
          "Pancake with maple syrup"
        ],
        [
          "Hot Beverage",
          "Tea And Coffee"
        ]
      ],
      "lunch": [
        [
          "Salad",
          "American corn salad with dressing"
        ],
        [
          "Dry Veg",
          "tandoori gobhi"
        ],
        [
          "Semi Dry Veg",
          "Veg kofta curry"
        ],
        [
          "Dal & Kadhi",
          "herb Rice"
        ],
        [
          "Rice",
          "Dal Rajasthani"
        ],
        [
          "Non-Veg",
          "Andhra Anda Masala"
        ],
        [
          "Indian Breads",
          "Chapati"
        ],
        [
          "Ancillaries",
          "Fry Papad/Pickle/Butter Milk"
        ],
        [
          "Dessert",
          "Orange slice cake"
        ],
        [
          "Bowled Over Lunch",
          "Bowled Over 2"
        ]
      ],
      "food_trotters": [
        "Manchurian chilli in garlic sauce",
        "Combination rice",
        "Chicken In Chilli Garlic Gravy"
      ],
      "snacks": [
        [
          "Evening Snacks",
          "Frankie"
        ],
        [
          "Jain Snack",
          "Frankie"
        ],
        [
          "Hot Beverage",
          "Hot Beverage"
        ]
      ],
      "dinner": [
        [
          "Salad",
          "Sprout salad"
        ],
        [
          "Semi - Dry",
          "Paneer Lazeez"
        ],
        [
          "Dry Prep",
          "Bhindi Do Payaza"
        ],
        [
          "Rice",
          "Jeera Rice"
        ],
        [
          "Dal (Lentils)",
          "Yellow Moong Dal"
        ],
        [
          "Non Veg",
          "Butter chicken"
        ],
        [
          "Indian Breads",
          "Chapatti"
        ],
        [
          "Ancillaries",
          "Papad/Pickle/Curd"
        ],
        [
          "Dessert",
          "Kesar sevyian kheer"
        ]
      ],
      "perfect_jodi": [
        "veg quesadilla",
        "Ranch",
        "Chicken Quesadilla"
      ]
    },
    "WEDNESDAY": {
      "breakfast": [
        [
          "Vegetarian Mains",
          "Dal pakwan"
        ],
        [
          "Fruits",
          "Banana"
        ],
        [
          "Cereals",
          "Oats With Milk"
        ],
        [
          "Egg Preparation (2 Nos Max)",
          "Punjabi Omelette"
        ],
        [
          "Hot Beverage",
          "Tea And Coffee"
        ]
      ],
      "lunch": [
        [
          "Salad",
          "Boondi Raita"
        ],
        [
          "Dry Veg",
          "Aloo Mutter Rassilla"
        ],
        [
          "Dal & Kadhi",
          "Paneer and soya Biryani"
        ],
        [
          "Rice",
          "Mirch ka Salan"
        ],
        [
          "Non-Veg",
          "Dum Chicken Biryani"
        ],
        [
          "Indian Breads",
          "Phulka"
        ],
        [
          "Ancillaries",
          "Papad/Pickle/virgin mojito"
        ],
        [
          "Dessert",
          "Shahi tukda"
        ],
        [
          "Bowled Over Lunch",
          "Bowled Over 3"
        ]
      ],
      "food_trotters": [
        "veg hakka noodles",
        "burnt galic vegetable sauce",
        "chicken hakka noodles"
      ],
      "snacks": [
        [
          "Evening Snacks",
          "Mumbai Bhel"
        ],
        [
          "Jain Snack",
          "Mumbai Bhel"
        ],
        [
          "Hot Beverage",
          "Hot Beverage"
        ]
      ],
      "dinner": [
        [
          "Salad",
          "Spring salad"
        ],
        [
          "Semi - Dry",
          "mix veg poriyal"
        ],
        [
          "Dry Prep",
          "Makhane ki sabji"
        ],
        [
          "Rice",
          "Lemon rice"
        ],
        [
          "Dal (Lentils)",
          "Dal Triveni"
        ],
        [
          "Non Veg",
          "Chicken Korma"
        ],
        [
          "Indian Breads",
          "Phulka"
        ],
        [
          "Ancillaries",
          "Papad/Pickle/Lemon Ice Tea"
        ],
        [
          "Dessert",
          "Tiramisu pastry"
        ]
      ],
      "perfect_jodi": [
        "methi corn rice",
        "cheese sauce (white sauce pasta sauce)",
        "Chicken rice"
      ]
    },
    "THURSDAY": {
      "breakfast": [
        [
          "Vegetarian Mains",
          "Fafda jalebi with kadhi and fried chilli"
        ],
        [
          "Fruits",
          "Orange + Watermelon"
        ],
        [
          "Cereals",
          "Corn Flakes with Milk"
        ],
        [
          "Egg Preparation (2 Nos Max)",
          "Bellpepper omelette"
        ],
        [
          "Hot Beverage",
          "Tea And Coffee"
        ]
      ],
      "lunch": [
        [
          "Salad",
          "Kachumbar salad"
        ],
        [
          "Dry Veg",
          "sev tamatar"
        ],
        [
          "Semi Dry Veg",
          "Undhiyu"
        ],
        [
          "Dal & Kadhi",
          "Gujrati dal"
        ],
        [
          "Rice",
          "Ghee rice"
        ],
        [
          "Non-Veg",
          "Chicken sukka"
        ],
        [
          "Indian Breads",
          "Methi roti"
        ],
        [
          "Ancillaries",
          "Papad/Pickle/Chaas"
        ],
        [
          "Dessert",
          "Mango shreekhand"
        ],
        [
          "Bowled Over Lunch",
          "Bowled Over 4"
        ]
      ],
      "food_trotters": [],
      "snacks": [
        [
          "Evening Snacks",
          "locho with sev and a drizzle of oil"
        ],
        [
          "Jain Snack",
          "locho with sev and a drizzle of oil"
        ],
        [
          "Hot Beverage",
          "Hot Beverage"
        ]
      ],
      "dinner": [
        [
          "Salad",
          "Cabbage lettuce Coleslaw"
        ],
        [
          "Semi - Dry",
          "Paneer Handi"
        ],
        [
          "Dry Prep",
          "ringan batata nu shaak"
        ],
        [
          "Rice",
          "Steamed Rice"
        ],
        [
          "Dal (Lentils)",
          "moong dal"
        ],
        [
          "Non Veg",
          "Dhaba Style Egg Curry"
        ],
        [
          "Indian Breads",
          "thepla"
        ],
        [
          "Ancillaries",
          "Papad/Pickle/Curd"
        ],
        [
          "Dessert",
          "Chocolate mohanthal"
        ]
      ],
      "perfect_jodi": [
        "messi roti",
        "punjabi kadhi pakoda",
        "chicken pakoda"
      ]
    },
    "FRIDAY": {
      "breakfast": [
        [
          "Vegetarian Mains",
          "rava upma with green chutney"
        ],
        [
          "Fruits",
          "Banana"
        ],
        [
          "Cereals",
          "Wheat Flakes with Milk"
        ],
        [
          "Egg Preparation (2 Nos Max)",
          "egg bytes but put masala"
        ],
        [
          "Hot Beverage",
          "Tea And Coffee"
        ]
      ],
      "lunch": [
        [
          "Salad",
          "Tossed salad"
        ],
        [
          "Dry Veg",
          "paneer ghassi"
        ],
        [
          "Semi Dry Veg",
          "kootu curry"
        ],
        [
          "Dal & Kadhi",
          "Sambhar"
        ],
        [
          "Rice",
          "Tamarind rice"
        ],
        [
          "Non-Veg",
          "Egg Kadai"
        ],
        [
          "Indian Breads",
          "Chapati"
        ],
        [
          "Ancillaries",
          "Papad / Pickle /Curd"
        ],
        [
          "Dessert",
          "red velvet pastry"
        ],
        [
          "Bowled Over Lunch",
          "Bowled Over 5"
        ]
      ],
      "food_trotters": [
        "Mexican Burrito Wrap",
        "sour ranch",
        "Mexican chicken burrito"
      ],
      "snacks": [
        [
          "Evening Snacks",
          "Pulugulu wth Peanut Chutney"
        ],
        [
          "Jain Snack",
          "Pulugulu wth Peanut Chutney"
        ],
        [
          "Hot Beverage",
          "Hot Beverage"
        ]
      ],
      "dinner": [
        [
          "Salad",
          "Green Salad"
        ],
        [
          "Semi - Dry",
          "Veg Hangama"
        ],
        [
          "Dry Prep",
          "Hara pyaaz aloo"
        ],
        [
          "Rice",
          "Jeera rice"
        ],
        [
          "Dal (Lentils)",
          "Dal Maharani"
        ],
        [
          "Non Veg",
          "Batti Da Murg"
        ],
        [
          "Indian Breads",
          "Phulka"
        ],
        [
          "Ancillaries",
          "Papad/Pickle/Chaas"
        ],
        [
          "Dessert",
          "Fruit Custard"
        ]
      ],
      "perfect_jodi": [
        "Pav bhaji",
        "Pav",
        "chicken keema"
      ]
    },
    "SATURDAY": {
      "breakfast": [
        [
          "Vegetarian Mains",
          "French toast maple syrup"
        ],
        [
          "Fruits",
          "Cut Muskmelon + Watermelon"
        ],
        [
          "Cereals",
          "Corn Flakes with Milk"
        ],
        [
          "Egg Preparation (2 Nos Max)",
          "French toast with egg"
        ],
        [
          "Hot Beverage",
          "Tea And Coffee"
        ]
      ],
      "lunch": [
        [
          "Salad",
          "Koshimbir"
        ],
        [
          "Dry Veg",
          "achari bhindi"
        ],
        [
          "Semi Dry Veg",
          "Rajasthani Sigdi ka paneer"
        ],
        [
          "Dal & Kadhi",
          "Rajma raseela"
        ],
        [
          "Rice",
          "Peas pulav"
        ],
        [
          "Non-Veg",
          "Chicken do Payaza"
        ],
        [
          "Indian Breads",
          "Palak Phulka"
        ],
        [
          "Ancillaries",
          "Papad/Pickle/ Kala Khatta"
        ],
        [
          "Dessert",
          "Black forest pastry"
        ],
        [
          "Bowled Over Lunch",
          "Bowled Over 6"
        ]
      ],
      "food_trotters": [
        "vegetable in burnt garlic sauce",
        "Malaysian Rice",
        "Chicken In Burnt Garlic Sauce Sauce"
      ],
      "snacks": [
        [
          "Evening Snacks",
          "Hara bhara Kabab"
        ],
        [
          "Jain Snack",
          "Hara bhara Kabab"
        ],
        [
          "Hot Beverage",
          "Hot Beverage"
        ]
      ],
      "dinner": [
        [
          "Salad",
          "Chimichurri Salad"
        ],
        [
          "Semi - Dry",
          "Pindi chole"
        ],
        [
          "Dry Prep",
          "Khada Bhaji"
        ],
        [
          "Rice",
          "Dum Rice"
        ],
        [
          "Dal (Lentils)",
          "Surti Dal"
        ],
        [
          "Non Veg",
          "Homestyle egg curry"
        ],
        [
          "Indian Breads",
          "Chapatti"
        ],
        [
          "Ancillaries",
          "Papad/Pickle/Curd"
        ],
        [
          "Dessert",
          "Rasgulla"
        ]
      ],
      "perfect_jodi": [
        "Dal makhani",
        "paneer masal kulcha",
        "chicken masala"
      ]
    },
    "SUNDAY": {
      "breakfast": [
        [
          "Vegetarian Mains",
          "Luchi gugni"
        ],
        [
          "Fruits",
          "Banana"
        ],
        [
          "Cereals",
          "Muesili with milk"
        ],
        [
          "Egg Preparation (2 Nos Max)",
          "Anda Bhurji With Brown bread"
        ],
        [
          "Hot Beverage",
          "Tea And Coffee"
        ]
      ],
      "lunch": [
        [
          "Salad",
          "Lemony Carrot Beetroot Salad"
        ],
        [
          "Dry Veg",
          "Soya corn masala"
        ],
        [
          "Semi Dry Veg",
          "Veg jalfrezi"
        ],
        [
          "Dal & Kadhi",
          "Maa ki dal"
        ],
        [
          "Rice",
          "Kari patta pulav"
        ],
        [
          "Non-Veg",
          "Egg kosha"
        ],
        [
          "Indian Breads",
          "Ajwain Phulka"
        ],
        [
          "Ancillaries",
          "Papad/Pickle/Curd"
        ],
        [
          "Dessert",
          "Rose Phirni"
        ],
        [
          "Bowled Over Lunch",
          "Bowled Over 7"
        ]
      ],
      "food_trotters": [
        "Mini Burgers",
        "Chicken mini burgers"
      ],
      "snacks": [
        [
          "Evening Snacks",
          "Gobhi manchurian"
        ],
        [
          "Jain Snack",
          "Gobhi manchurian"
        ],
        [
          "Hot Beverage",
          "Hot Beverage"
        ]
      ],
      "dinner": [
        [
          "Salad",
          "caesar salad with dressing"
        ],
        [
          "Semi - Dry",
          "Veg korma"
        ],
        [
          "Dry Prep",
          "Lassoni aloo methi dry"
        ],
        [
          "Rice",
          "Steamed Rice"
        ],
        [
          "Dal (Lentils)",
          "Pepper rasam"
        ],
        [
          "Non Veg",
          "Dum ka chicken"
        ],
        [
          "Indian Breads",
          "Pulka"
        ],
        [
          "Ancillaries",
          "Papad/Pickle/Chaas"
        ],
        [
          "Dessert",
          "Pineapple Pastry"
        ]
      ],
      "perfect_jodi": [
        "malai panner kofta",
        "malabar paratha",
        "malai chicken"
      ]
    }
  }
}""")

# Shared state between Python Scheduler/Web Server and the WhatsApp QR Bridge
BRIDGE_STATE = {
    "connected": False,
    "qr_data_url": None,
    "self_jid": None,
    "group_jid": None,
    "group_name": None,
    "outbox": [],  # Queued outbound messages to send exclusively to "Mess Menu"
    "recent_logs": [],
}
STATE_LOCK = threading.Lock()


def add_bridge_log(entry: str):
    ts = datetime.now(TIMEZONE).strftime("%H:%M:%S")
    with STATE_LOCK:
        BRIDGE_STATE["recent_logs"].append(f"[{ts}] {entry}")
        if len(BRIDGE_STATE["recent_logs"]) > 25:
            BRIDGE_STATE["recent_logs"] = BRIDGE_STATE["recent_logs"][-25:]


# ---------------------------------------------------------------------------
# Excel Menu Parser (BITSoM - Kamba Layout)
# ---------------------------------------------------------------------------
def normalize_day_name(raw_cell: str) -> str:
    if not raw_cell:
        return ""
    upper = str(raw_cell).strip().upper()
    for day in DAYS_ORDER:
        if upper.startswith(day):
            return day
    return upper


def is_non_veg_entry(category: str, item_text: str) -> bool:
    cat_lower = category.lower()
    item_lower = item_text.lower()
    keywords = (
        "non-veg", "non veg", "egg", "chicken", "mutton",
        "fish", "prawn", "anda", "murg", "omelette", "bhurji",
    )
    return any(k in cat_lower or k in item_lower for k in keywords)


def parse_date_string(date_str: str) -> date | None:
    if not date_str:
        return None
    for fmt in ("%d/%m/%y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%y", "%d-%m-%Y"):
        try:
            return datetime.strptime(date_str.strip(), fmt).date()
        except ValueError:
            continue
    return None


def is_menu_outdated(menu_data: dict, today_date: date) -> bool:
    parsed_dates = [
        d for d in (parse_date_string(v) for v in menu_data.get("dates", {}).values())
        if d is not None
    ]
    if not parsed_dates:
        return False
    return max(parsed_dates) < today_date


def parse_menu_excel(file_path: str) -> dict:
    """Parses the BITSoM - Kamba weekly Excel file (.xlsx) into a structured dict."""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    sheet = wb.active

    rows = list(sheet.iter_rows(values_only=True))
    if not rows:
        raise ValueError("The uploaded Excel file is empty.")

    col_to_day = {}
    dates_by_day = {}
    day_labels = {}

    date_row = rows[0] if len(rows) > 0 else []
    for r in rows[:5]:
        if r and any(normalize_day_name(c) in DAYS_ORDER for c in r[1:8] if c):
            for col_idx in range(1, min(8, len(r))):
                raw_day = r[col_idx]
                day_key = normalize_day_name(raw_day)
                if day_key in DAYS_ORDER:
                    col_to_day[col_idx] = day_key
                    day_labels[day_key] = str(raw_day).strip()
                    if col_idx < len(date_row) and date_row[col_idx] is not None:
                        val = date_row[col_idx]
                        if isinstance(val, datetime):
                            dates_by_day[day_key] = val.strftime("%d/%m/%y")
                        else:
                            dates_by_day[day_key] = str(val).strip()
            break

    if not col_to_day:
        for idx, day_key in enumerate(DAYS_ORDER, start=1):
            col_to_day[idx] = day_key
            day_labels[day_key] = day_key

    meals = {
        day: {
            "breakfast": [],
            "lunch": [],
            "food_trotters": [],
            "snacks": [],
            "dinner": [],
            "perfect_jodi": [],
        }
        for day in DAYS_ORDER
    }

    current_section = None
    sub_row_counter = 0

    for row in rows[1:]:
        if not row or all(cell is None or str(cell).strip() == "" for cell in row[:8]):
            continue

        first_col = str(row[0]).strip() if row[0] is not None else ""
        first_col_lower = first_col.lower()

        second_col_day = normalize_day_name(row[1]) if len(row) > 1 and row[1] else ""
        if first_col_lower in SECTION_HEADERS and second_col_day in DAYS_ORDER:
            current_section = SECTION_HEADERS[first_col_lower]
            sub_row_counter = 0
            continue

        if not current_section:
            continue

        sub_row_counter += 1

        for col_idx, day_key in col_to_day.items():
            if col_idx >= len(row):
                continue
            val = row[col_idx]
            if val is None or str(val).strip() == "":
                continue
            item_text = " ".join(str(val).strip().split())

            if current_section == "breakfast":
                category = first_col or f"Item {sub_row_counter}"
                meals[day_key]["breakfast"].append((category, item_text))

            elif current_section == "lunch":
                category = first_col or f"Item {sub_row_counter}"
                meals[day_key]["lunch"].append((category, item_text))

            elif current_section == "lunch_bowled_over" and INCLUDE_BOWLED_OVER_IN_LUNCH:
                category = first_col or "Bowled Over Lunch"
                meals[day_key]["lunch"].append((category, item_text))

            elif current_section == "lunch_food_trotters":
                meals[day_key]["food_trotters"].append(item_text)

            elif current_section == "snacks":
                category = first_col or f"Snack {sub_row_counter}"
                if not INCLUDE_JAIN and "jain" in category.lower():
                    continue
                meals[day_key]["snacks"].append((category, item_text))

            elif current_section == "dinner":
                category = first_col or f"Item {sub_row_counter}"
                meals[day_key]["dinner"].append((category, item_text))

            elif current_section == "dinner_perfect_jodi":
                meals[day_key]["perfect_jodi"].append(item_text)

    return {
        "dates": dates_by_day,
        "day_labels": day_labels,
        "meals": meals,
    }


def get_active_menu_data() -> dict:
    """Loads the latest uploaded current_menu.xlsx, or falls back to the pre-loaded Oct 19-25 menu."""
    if os.path.exists(MENU_FILE_PATH):
        try:
            return parse_menu_excel(MENU_FILE_PATH)
        except Exception as exc:
            logger.warning("Could not parse %s, using pre-loaded menu: %s", MENU_FILE_PATH, exc)
    return DEFAULT_MENU_DATA


# ---------------------------------------------------------------------------
# WhatsApp Message Formatter
# ---------------------------------------------------------------------------
def format_meal_message(menu_data: dict, day_key: str, meal_type: str) -> str:
    day_key = day_key.upper()
    meal_type = meal_type.lower()

    meal_titles = {
        "breakfast": "🌅 *BREAKFAST MENU (6:00 AM)*",
        "lunch":     "☀️ *LUNCH MENU (11:00 AM)*",
        "snacks":    "☕ *SNACKS MENU (3:00 PM)*",
        "dinner":    "🌙 *DINNER MENU (7:00 PM)*",
    }

    header = meal_titles.get(meal_type, f"🍽️ *{meal_type.upper()} MENU*")
    day_label = menu_data.get("day_labels", {}).get(day_key, day_key)
    date_str = menu_data.get("dates", {}).get(day_key, "")
    date_suffix = f" ({date_str})" if date_str else ""

    day_meals = menu_data.get("meals", {}).get(day_key, {})
    entries = day_meals.get(meal_type, [])

    lines = [f"{header}", f"📅 *{day_label}{date_suffix}*", ""]

    if not entries:
        lines.append("_No main items listed for this meal._")
    else:
        grouped = {}
        order = []
        for category, item in entries:
            if not INCLUDE_NON_VEG and is_non_veg_entry(category, item):
                continue
            if category not in grouped:
                grouped[category] = []
                order.append(category)
            if item not in grouped[category]:
                grouped[category].append(item)

        for category in order:
            items_str = " | ".join(grouped[category])
            lines.append(f"• *{category}:* {items_str}")

    # Append Food Trotters to Lunch
    if meal_type == "lunch" and INCLUDE_FOOD_TROTTERS_IN_LUNCH:
        ft_items = [
            item for item in day_meals.get("food_trotters", [])
            if INCLUDE_NON_VEG or not is_non_veg_entry("Food Trotters", item)
        ]
        lines.append("")
        lines.append("🍴 *FOOD TROTTERS:*")
        if ft_items:
            for item in ft_items:
                lines.append(f"  ◦ {item}")
        else:
            lines.append("  _None today_")

    # Append Perfect Jodi / Bowled Over to Dinner
    if meal_type == "dinner" and INCLUDE_PERFECT_JODI_IN_DINNER:
        pj_items = [
            item for item in day_meals.get("perfect_jodi", [])
            if INCLUDE_NON_VEG or not is_non_veg_entry("Perfect Jodi", item)
        ]
        lines.append("")
        lines.append("🍛 *PERFECT JODI / BOWLED OVER:*")
        if pj_items:
            for item in pj_items:
                lines.append(f"  ◦ {item}")
        else:
            lines.append("  _None today_")

    return "\n".join(lines)


def format_full_day_menu(menu_data: dict, day_key: str) -> str:
    parts = [
        format_meal_message(menu_data, day_key, meal)
        for meal in ("breakfast", "lunch", "snacks", "dinner")
    ]
    return ("\n" + "─" * 24 + "\n").join(parts)


def get_next_upcoming_meal(now_ist: datetime) -> tuple[str, str]:
    current_hm = now_ist.strftime("%H:%M")
    today_idx = now_ist.weekday()
    today_key = DAYS_ORDER[today_idx]

    for meal in ("breakfast", "lunch", "snacks", "dinner"):
        if current_hm <= MEAL_SCHEDULE[meal]:
            return today_key, meal

    tomorrow_key = DAYS_ORDER[(today_idx + 1) % 7]
    return tomorrow_key, "breakfast"


# ---------------------------------------------------------------------------
# On-Demand Query Resolver (Strictly for the "Mess Menu" Group)
# ---------------------------------------------------------------------------
def resolve_user_query(incoming_text: str, menu_data: dict, now_ist: datetime) -> str | None:
    cleaned = incoming_text.strip().lower()
    if cleaned.startswith("!"):
        cleaned = cleaned[1:].strip()
    for ch in ("?", "!", ".", ",", "'s", "’s"):
        cleaned = cleaned.replace(ch, "")
    tokens = [
        t for t in cleaned.split()
        if t not in ("please", "pls", "send", "show", "get", "what", "is", "the", "for", "of", "on")
    ]

    today_idx = now_ist.weekday()
    today_key = DAYS_ORDER[today_idx]
    tomorrow_key = DAYS_ORDER[(today_idx + 1) % 7]

    if cleaned in ("menu", "today", "todays menu", "today menu", "all", "full", "full menu"):
        return format_full_day_menu(menu_data, today_key)

    if cleaned in ("next", "now", "upcoming", "next meal"):
        d_key, m_type = get_next_upcoming_meal(now_ist)
        return format_meal_message(menu_data, d_key, m_type)

    if "food trotter" in cleaned or "trotter" in cleaned:
        return format_meal_message(menu_data, today_key, "lunch")

    if "perfect jodi" in cleaned or "jodi" in cleaned:
        return format_meal_message(menu_data, today_key, "dinner")

    if cleaned in ("help", "commands", "bot"):
        return (
            "👋 *BITSoM Mess Menu Bot (\"Mess Menu\" Group)*\n"
            "Text any of these keywords in this group:\n"
            "• *menu* or *today* — Full menu for today (all 4 meals)\n"
            "• *breakfast*, *lunch*, *snacks*, or *dinner* — Specific meal for today\n"
            "• *next* — The next upcoming meal\n"
            "• *tomorrow* (or *tomorrow lunch*) — Tomorrow's menu\n"
            "• *monday*, *thursday dinner*, etc. — Any day or meal of the week\n"
            "• Drop a `.xlsx` file here on Monday to update the week's menu!"
        )

    target_day = None
    target_meal = None

    for tok in tokens:
        if tok in ("today", "todays"):
            target_day = today_key
        elif tok in ("tomorrow", "tmrw", "tom", "tomorrows"):
            target_day = tomorrow_key
        elif tok in DAY_ALIASES:
            target_day = DAY_ALIASES[tok]
        elif tok in MEAL_ALIASES:
            target_meal = MEAL_ALIASES[tok]

    if target_day and target_meal and len(tokens) <= 4:
        return format_meal_message(menu_data, target_day, target_meal)
    if target_meal and len(tokens) <= 3:
        return format_meal_message(menu_data, today_key, target_meal)
    if target_day and len(tokens) <= 3:
        return format_full_day_menu(menu_data, target_day)

    return None


# ---------------------------------------------------------------------------
# Queue Outbound WhatsApp Message to the "Mess Menu" Group
# ---------------------------------------------------------------------------
def queue_whatsapp_message(text: str):
    with STATE_LOCK:
        BRIDGE_STATE["outbox"].append(text)
        connected = BRIDGE_STATE["connected"]
        group_jid = BRIDGE_STATE["group_jid"]
    status_note = f"target={group_jid}" if (connected and group_jid) else "waiting for 'Mess Menu' group discovery"
    add_bridge_log(f"Queued outbound message ({len(text)} chars) [{status_note}]")
    logger.info("Queued WhatsApp message (%d chars)", len(text))


# ---------------------------------------------------------------------------
# Background Daily Meal Scheduler (6 AM, 11 AM, 3 PM, 7 PM IST + Monday Reminder)
# ---------------------------------------------------------------------------
def scheduler_loop():
    sent_log = set()
    logger.info("Scheduler active in %s: %s", TIMEZONE, MEAL_SCHEDULE)

    while True:
        now_ist = datetime.now(TIMEZONE)
        current_hm = now_ist.strftime("%H:%M")
        today_str = now_ist.strftime("%Y-%m-%d")
        day_key = DAYS_ORDER[now_ist.weekday()]

        # 1. Monday 9:00 AM Upload Reminder if the active menu is outdated
        if day_key == "MONDAY" and current_hm == MONDAY_REMINDER_TIME:
            rem_key = (today_str, "monday_reminder")
            if rem_key not in sent_log:
                sent_log.add(rem_key)
                menu_data = get_active_menu_data()
                if is_menu_outdated(menu_data, now_ist.date()):
                    queue_whatsapp_message(
                        "⏰ *Monday Menu Reminder!*\n"
                        "Please drop this week's new Mess Menu Excel (`.xlsx`) file here in the "
                        "*Mess Menu* group so I can update your daily alerts!"
                    )

        # 2. Scheduled Meal Texts (06:00, 11:00, 15:00, 19:00 IST)
        for meal_type, target_hm in MEAL_SCHEDULE.items():
            key = (today_str, meal_type)
            if current_hm == target_hm and key not in sent_log:
                sent_log.add(key)
                try:
                    menu_data = get_active_menu_data()
                    msg = format_meal_message(menu_data, day_key, meal_type)
                    queue_whatsapp_message(msg)
                except Exception as exc:
                    logger.exception("Error formatting scheduled %s menu: %s", meal_type, exc)

        time.sleep(20)


# ---------------------------------------------------------------------------
# Auto-Generated Node.js Baileys QR-Code Bridge (v3 - Strictly "Mess Menu" Group)
# ---------------------------------------------------------------------------
NODE_BRIDGE_SCRIPT = r"""
const {
    default: makeWASocket,
    useMultiFileAuthState,
    DisconnectReason,
    downloadMediaMessage,
    fetchLatestBaileysVersion,
    makeCacheableSignalKeyStore
} = require('@whiskeysockets/baileys');
const QRCode = require('qrcode');
const pino = require('pino');
const http = require('http');

const PY_PORT = process.env.PORT || 10000;
const TARGET_GROUP_NAME = 'mess menu';
const processedMsgIds = new Set();
const groupSubjectCache = new Map();

let messMenuGroupJid = null;
let messMenuGroupTitle = null;
let activeSock = null;
let isConnected = false;
let isSendingOutbox = false;

function postToPython(path, payload) {
    return new Promise((resolve) => {
        const data = JSON.stringify(payload);
        const req = http.request({
            hostname: '127.0.0.1',
            port: PY_PORT,
            path: path,
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Content-Length': Buffer.byteLength(data) }
        }, (res) => {
            let body = '';
            res.on('data', chunk => body += chunk);
            res.on('end', () => {
                try { resolve(JSON.parse(body)); } catch (e) { resolve({}); }
            });
        });
        req.on('error', () => resolve({}));
        req.write(data);
        req.end();
    });
}

function logToPython(msg) {
    console.log(msg);
    postToPython('/internal/log', { message: msg }).catch(() => {});
}

function withTimeout(promise, ms, label) {
    return Promise.race([
        promise,
        new Promise((_, reject) =>
            setTimeout(() => reject(new Error(label + ' timed out after ' + (ms / 1000) + 's')), ms)
        )
    ]);
}

process.on('uncaughtException', (err) => {
    logToPython('⚠️ Bridge caught exception: ' + (err?.message || String(err)));
});
process.on('unhandledRejection', (reason) => {
    logToPython('⚠️ Bridge caught rejection: ' + (reason?.message || String(reason)));
});

function unwrapMessage(m) {
    if (!m) return null;
    if (m.ephemeralMessage?.message) return unwrapMessage(m.ephemeralMessage.message);
    if (m.deviceSentMessage?.message) return unwrapMessage(m.deviceSentMessage.message);
    if (m.viewOnceMessage?.message) return unwrapMessage(m.viewOnceMessage.message);
    if (m.viewOnceMessageV2?.message) return unwrapMessage(m.viewOnceMessageV2.message);
    if (m.documentWithCaptionMessage?.message) return unwrapMessage(m.documentWithCaptionMessage.message);
    return m;
}

function stripJidUser(jid) {
    if (!jid) return '';
    return String(jid).split('@')[0].split(':')[0];
}

function extractTimestampSeconds(ts) {
    if (!ts) return 0;
    if (typeof ts === 'number') return ts;
    if (typeof ts === 'object' && typeof ts.low === 'number') return ts.low;
    const parsed = Number(ts);
    return Number.isNaN(parsed) ? 0 : parsed;
}

function isTargetGroupName(subject) {
    if (!subject) return false;
    return String(subject).trim().toLowerCase() === TARGET_GROUP_NAME;
}

// Scans all WhatsApp groups the user is in to find the one named "Mess Menu"
async function discoverMessMenuGroup(sock) {
    try {
        const groups = await withTimeout(sock.groupFetchAllParticipating(), 15000, 'groupFetchAllParticipating');
        const allGroups = Object.values(groups || {});
        for (const g of allGroups) {
            if (g && g.id && g.subject) {
                groupSubjectCache.set(g.id, g.subject);
                if (isTargetGroupName(g.subject)) {
                    messMenuGroupJid = g.id;
                    messMenuGroupTitle = g.subject;
                    await postToPython('/internal/status', {
                        group_jid: messMenuGroupJid,
                        group_name: messMenuGroupTitle,
                    });
                    logToPython('🎯 Found target group "' + g.subject + '" (' + g.id + ')');
                    return g.id;
                }
            }
        }
        logToPython('⚠️ No WhatsApp group named "Mess Menu" found yet (checked ' + allGroups.length + ' groups).');
    } catch (err) {
        logToPython('⚠️ Group scan warning: ' + (err?.message || String(err)));
    }
    return messMenuGroupJid;
}

// Checks whether a given group JID is the "Mess Menu" group
async function isMessMenuGroupJid(sock, remoteJid) {
    if (!remoteJid || !remoteJid.endsWith('@g.us')) return false;
    if (messMenuGroupJid && remoteJid === messMenuGroupJid) return true;

    let subject = groupSubjectCache.get(remoteJid);
    if (!subject) {
        try {
            const meta = await withTimeout(sock.groupMetadata(remoteJid), 8000, 'groupMetadata');
            if (meta && meta.subject) {
                subject = meta.subject;
                groupSubjectCache.set(remoteJid, subject);
            }
        } catch (e) {
            return false;
        }
    }

    if (isTargetGroupName(subject)) {
        messMenuGroupJid = remoteJid;
        messMenuGroupTitle = subject;
        await postToPython('/internal/status', {
            group_jid: messMenuGroupJid,
            group_name: messMenuGroupTitle,
        });
        logToPython('🎯 Locked onto target group "' + subject + '" (' + remoteJid + ')');
        return true;
    }
    return false;
}

// Sends a message strictly and exclusively to the "Mess Menu" group
async function sendToMessMenuGroup(text) {
    if (!activeSock || !isConnected) {
        throw new Error('WhatsApp socket is not connected right now');
    }

    if (!messMenuGroupJid) {
        await discoverMessMenuGroup(activeSock);
    }

    if (!messMenuGroupJid) {
        throw new Error('Group "Mess Menu" not found yet — please create or send a message in the "Mess Menu" group');
    }

    logToPython('⏳ Sending message (' + text.length + ' chars) to "' + (messMenuGroupTitle || 'Mess Menu') + '" (' + messMenuGroupJid + ')...');
    await withTimeout(
        activeSock.sendMessage(messMenuGroupJid, { text: '\u200B' + text }),
        15000,
        'sendMessage(Mess Menu)'
    );
    return (messMenuGroupTitle || 'Mess Menu') + ' (' + messMenuGroupJid + ')';
}

// Single global outbox poller (never duplicated across reconnects)
setInterval(async () => {
    if (!isConnected || !activeSock || isSendingOutbox) return;
    isSendingOutbox = true;
    try {
        const res = await postToPython('/internal/outbox_peek', {});
        if (res.messages && res.messages.length > 0) {
            for (const m of res.messages) {
                const usedTarget = await sendToMessMenuGroup(m);
                await postToPython('/internal/outbox_pop', {
                    delivered_to: usedTarget,
                    chars: m.length,
                });
            }
        }
    } catch (err) {
        logToPython('❌ Outbox delivery waiting: ' + (err?.message || String(err)));
    } finally {
        isSendingOutbox = false;
    }
}, 4000);

async function startWhatsApp() {
    const { state, saveCreds } = await useMultiFileAuthState('./wa_auth_session');
    const silentLogger = pino({ level: 'silent' });

    let waVersion = undefined;
    try {
        const vInfo = await withTimeout(fetchLatestBaileysVersion(), 6000, 'fetchLatestBaileysVersion');
        waVersion = vInfo.version;
        logToPython('ℹ️ Using WhatsApp Web version: ' + waVersion.join('.'));
    } catch (e) {
        logToPython('ℹ️ Using default bundled WhatsApp Web version');
    }

    const authConfig = {
        creds: state.creds,
        keys: typeof makeCacheableSignalKeyStore === 'function'
            ? makeCacheableSignalKeyStore(state.keys, silentLogger)
            : state.keys,
    };

    const sock = makeWASocket({
        ...(waVersion ? { version: waVersion } : {}),
        auth: authConfig,
        logger: silentLogger,
        printQRInTerminal: true,
        syncFullHistory: false,
        keepAliveIntervalMs: 20000,
        defaultQueryTimeoutMs: 15000,
        getMessage: async () => ({ conversation: 'BITSoM Mess Menu' }),
    });
    activeSock = sock;

    sock.ev.on('creds.update', saveCreds);

    sock.ev.on('connection.update', async (update) => {
        const { connection, lastDisconnect, qr } = update;
        if (qr) {
            isConnected = false;
            const qrUrl = await QRCode.toDataURL(qr);
            await postToPython('/internal/status', { connected: false, qr_data_url: qrUrl });
        }
        if (connection === 'open') {
            const currentPhone = stripJidUser(sock.user?.id || state.creds?.me?.id);
            const myJid = currentPhone + '@s.whatsapp.net';
            isConnected = true;
            console.log('✅ Connected to WhatsApp:', myJid);
            await postToPython('/internal/status', {
                connected: true,
                qr_data_url: null,
                self_jid: myJid,
            });
            // Automatically scan groups for "Mess Menu"
            await discoverMessMenuGroup(sock);
        }
        if (connection === 'close') {
            isConnected = false;
            const code = lastDisconnect?.error?.output?.statusCode;
            await postToPython('/internal/status', { connected: false });
            if (code !== DisconnectReason.loggedOut) {
                setTimeout(startWhatsApp, 4000);
            } else {
                logToPython('⚠️ WhatsApp session logged out. Please re-scan QR code.');
            }
        }
    });

    // Detect if user creates or renames a group to "Mess Menu"
    sock.ev.on('groups.upsert', async (newGroups) => {
        for (const g of (newGroups || [])) {
            if (g && g.id && g.subject) {
                groupSubjectCache.set(g.id, g.subject);
                if (isTargetGroupName(g.subject)) {
                    messMenuGroupJid = g.id;
                    messMenuGroupTitle = g.subject;
                    await postToPython('/internal/status', {
                        group_jid: messMenuGroupJid,
                        group_name: messMenuGroupTitle,
                    });
                    logToPython('🎯 Detected new target group "' + g.subject + '" (' + g.id + ')');
                }
            }
        }
    });

    sock.ev.on('groups.update', async (updates) => {
        for (const u of (updates || [])) {
            if (u && u.id && u.subject) {
                groupSubjectCache.set(u.id, u.subject);
                if (isTargetGroupName(u.subject)) {
                    messMenuGroupJid = u.id;
                    messMenuGroupTitle = u.subject;
                    await postToPython('/internal/status', {
                        group_jid: messMenuGroupJid,
                        group_name: messMenuGroupTitle,
                    });
                    logToPython('🎯 Updated target group "' + u.subject + '" (' + u.id + ')');
                }
            }
        }
    });

    // Listen ONLY for messages & Excel files inside the "Mess Menu" group
    sock.ev.on('messages.upsert', async ({ messages, type }) => {
        if (type !== 'notify' && type !== 'append') return;

        for (const msg of messages) {
            try {
                if (!msg.message || !msg.key) continue;
                const msgId = msg.key.id;
                if (!msgId || processedMsgIds.has(msgId)) continue;

                // Ignore old history messages (> 120 seconds old)
                const msgTs = extractTimestampSeconds(msg.messageTimestamp);
                const nowSec = Math.floor(Date.now() / 1000);
                if (msgTs > 0 && (nowSec - msgTs) > 120) continue;

                const remoteJid = msg.key.remoteJid || '';
                // STRICT FILTER: Ignore all non-group chats (never reply to personal/self chats)
                if (!remoteJid.endsWith('@g.us')) continue;

                // STRICT FILTER: Only respond if the group is named "Mess Menu"
                const isTargetGroup = await isMessMenuGroupJid(sock, remoteJid);
                if (!isTargetGroup) continue;

                const content = unwrapMessage(msg.message);
                if (!content) continue;

                const text = (
                    content.conversation ||
                    content.extendedTextMessage?.text ||
                    content.imageMessage?.caption ||
                    content.documentMessage?.caption ||
                    ''
                ).trim();

                // Ignore the bot's own replies (prefixed with zero-width space \u200B)
                if (text.startsWith('\u200B')) {
                    processedMsgIds.add(msgId);
                    continue;
                }

                processedMsgIds.add(msgId);

                // 1. Check if an Excel (.xlsx) document was dropped into the "Mess Menu" group
                const doc = content.documentMessage;
                if (doc && (doc.fileName?.endsWith('.xlsx') || doc.mimetype?.includes('spreadsheet') || doc.mimetype?.includes('excel'))) {
                    logToPython('📥 Downloading uploaded Excel file from "Mess Menu" group...');
                    const buffer = await downloadMediaMessage(msg, 'buffer', {});
                    const res = await postToPython('/internal/upload_xlsx', {
                        filename: doc.fileName || 'menu.xlsx',
                        b64: buffer.toString('base64'),
                        remote_jid: remoteJid,
                    });
                    if (res.reply) {
                        await withTimeout(
                            sock.sendMessage(remoteJid, { text: '\u200B' + res.reply }),
                            15000,
                            'sendMessage(excel_reply)'
                        );
                        logToPython('✅ Sent Excel update confirmation to "Mess Menu" group');
                    }
                    continue;
                }

                // 2. Handle text query inside the "Mess Menu" group
                if (text) {
                    const res = await postToPython('/internal/query', {
                        text: text,
                        remote_jid: remoteJid,
                    });
                    if (res.reply) {
                        await withTimeout(
                            sock.sendMessage(remoteJid, { text: '\u200B' + res.reply }),
                            15000,
                            'sendMessage(query_reply)'
                        );
                        logToPython('✅ Delivered reply for "' + text + '" to "Mess Menu" group');
                    }
                }
            } catch (err) {
                logToPython('❌ Group message handler error: ' + (err?.message || String(err)));
            }
        }
    });
}

startWhatsApp();
"""


def start_node_qr_bridge():
    """Writes wa_bridge.js, installs Baileys if needed, and supervises the WhatsApp QR bridge."""
    with open("wa_bridge.js", "w", encoding="utf-8") as f:
        f.write(NODE_BRIDGE_SCRIPT.strip())

    if not os.path.exists("node_modules/@whiskeysockets/baileys"):
        logger.info("Installing WhatsApp QR bridge packages (@whiskeysockets/baileys, qrcode, pino)...")
        subprocess.run(
            ["npm", "install", "--no-audit", "--no-fund", "@whiskeysockets/baileys", "qrcode", "pino"],
            check=False,
        )

    while True:
        logger.info("Starting WhatsApp QR-Code Bridge process...")
        add_bridge_log("🚀 Starting WhatsApp QR bridge process (v3 — 'Mess Menu' Group Only)...")
        proc = subprocess.Popen(["node", "wa_bridge.js"])
        exit_code = proc.wait()
        with STATE_LOCK:
            BRIDGE_STATE["connected"] = False
        add_bridge_log(f"⚠️ WhatsApp bridge exited (code {exit_code}), restarting in 3s...")
        time.sleep(3)


# ---------------------------------------------------------------------------
# Web Dashboard & Internal Bridge Server
# ---------------------------------------------------------------------------
class BotWebServer(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        return  # Keep console logs clean

    def do_GET(self):
        if self.path == "/health":
            self.send_response(200)
            self.send_header("Content-Type", "text/plain")
            self.end_headers()
            self.wfile.write(b"OK")
            return

        if self.path == "/send-test":
            now_ist = datetime.now(TIMEZONE)
            d_key, m_type = get_next_upcoming_meal(now_ist)
            menu_data = get_active_menu_data()
            queue_whatsapp_message(
                "✅ *Test Alert from Your BITSoM Mess Bot!*\n\n" +
                format_meal_message(menu_data, d_key, m_type)
            )
            self.send_response(302)
            self.send_header("Location", "/?sent=1")
            self.end_headers()
            return

        now_ist = datetime.now(TIMEZONE)
        today_key = DAYS_ORDER[now_ist.weekday()]
        menu_data = get_active_menu_data()
        today_preview = html.escape(format_full_day_menu(menu_data, today_key))

        with STATE_LOCK:
            connected = BRIDGE_STATE["connected"]
            qr_url = BRIDGE_STATE["qr_data_url"]
            self_jid = BRIDGE_STATE["self_jid"] or ""
            group_jid = BRIDGE_STATE["group_jid"] or ""
            group_name = BRIDGE_STATE["group_name"] or ""
            logs_list = list(BRIDGE_STATE["recent_logs"])
            outbox_count = len(BRIDGE_STATE["outbox"])

        logs_html = html.escape("\n".join(logs_list)) if logs_list else "No messages processed yet."
        sent_banner = (
            '<div style="background:#dbeafe;color:#1e40af;padding:12px;border-radius:8px;margin-bottom:16px;">'
            '📨 <b>Test message queued for "Mess Menu" group!</b> Watch the Activity Log below (refresh in 5 seconds) to see delivery confirmation.'
            '</div>'
            if "sent=1" in self.path else ""
        )

        if connected:
            group_badge = (
                f'🎯 <b>Target Group Locked:</b> "{html.escape(group_name)}" (<code>{html.escape(group_jid)}</code>)'
                if group_jid
                else '⚠️ <b>Target Group Not Found Yet:</b> Make sure you have a WhatsApp group named exactly <b>Mess Menu</b> and send any message in it!'
            )
            status_html = f"""
            <div style="background:#dcfce7;color:#166534;padding:16px;border-radius:10px;margin-bottom:20px;">
                <strong>🟢 Connected to WhatsApp! (v3 — "Mess Menu" Group Only)</strong> ({html.escape(self_jid)})<br>
                {group_badge}<br>
                <small><b>Pending in Outbox: {outbox_count}</b></small><br><br>
                Open your <b>"Mess Menu"</b> WhatsApp group and text <code>menu</code>, <code>breakfast</code>, <code>lunch</code>, <code>snacks</code>, <code>dinner</code>, or <code>next</code>!<br><br>
                <a href="/send-test" style="display:inline-block;background:#16a34a;color:#fff;text-decoration:none;padding:10px 16px;border-radius:8px;font-weight:600;">📲 Send Test Menu to "Mess Menu" Group Now</a>
            </div>
            """
        elif qr_url:
            status_html = f"""
            <div style="background:#fef9c3;color:#854d0e;padding:20px;border-radius:10px;margin-bottom:20px;text-align:center;">
                <strong>📱 Scan this QR Code with your Phone to Link Your WhatsApp:</strong><br>
                <small>Open WhatsApp → Settings / Menu → <b>Linked Devices</b> → <b>Link a Device</b></small><br><br>
                <img src="{qr_url}" style="width:260px;height:260px;border:4px solid #fff;border-radius:12px;box-shadow:0 4px 12px rgba(0,0,0,0.1);" />
                <p><small>Page refreshes automatically every 8 seconds...</small></p>
            </div>
            <script>setTimeout(() => location.reload(), 8000);</script>
            """
        else:
            status_html = """
            <div style="background:#e0f2fe;color:#075985;padding:16px;border-radius:10px;margin-bottom:20px;">
                ⏳ Reconnecting to WhatsApp / Generating QR code... (Page refreshes every 5 seconds)
            </div>
            <script>setTimeout(() => location.reload(), 5000);</script>
            """

        page = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>BITSoM Mess Menu WhatsApp Bot (v3)</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
        body {{ font-family: system-ui, -apple-system, sans-serif; max-width: 680px; margin: 30px auto; padding: 0 16px; background: #f8fafc; color: #0f172a; }}
        .card {{ background: #fff; padding: 24px; border-radius: 14px; box-shadow: 0 2px 10px rgba(0,0,0,0.06); margin-bottom: 20px; }}
        pre {{ background: #f1f5f9; padding: 16px; border-radius: 10px; white-space: pre-wrap; font-size: 14px; line-height: 1.5; }}
        code {{ background: #e2e8f0; padding: 2px 6px; border-radius: 4px; }}
    </style>
</head>
<body>
    <div class="card">
        <h2 style="margin-top:0;">🍽️ BITSoM Mess Menu Bot (v3 — "Mess Menu" Group Only)</h2>
        {sent_banner}
        {status_html}
        <p><b>Daily Schedule (IST):</b> Breakfast 06:00 AM • Lunch 11:00 AM • Snacks 03:00 PM • Dinner 07:00 PM • Monday Upload Reminder 09:00 AM</p>
    </div>
    <div class="card">
        <h3 style="margin-top:0;">🔍 Recent Bot Activity Log</h3>
        <pre>{logs_html}</pre>
    </div>
    <div class="card">
        <h3 style="margin-top:0;">📅 Today's Loaded Menu Preview ({today_key})</h3>
        <pre>{today_preview}</pre>
    </div>
</body>
</html>"""
        body = page.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length).decode("utf-8") if length else "{}"
        data = json.loads(raw) if raw else {}
        response_payload = {}

        if self.path == "/internal/status":
            with STATE_LOCK:
                if "connected" in data:
                    BRIDGE_STATE["connected"] = bool(data["connected"])
                if "qr_data_url" in data:
                    BRIDGE_STATE["qr_data_url"] = data["qr_data_url"]
                if "self_jid" in data and data["self_jid"]:
                    BRIDGE_STATE["self_jid"] = data["self_jid"]
                if "group_jid" in data and data["group_jid"]:
                    BRIDGE_STATE["group_jid"] = data["group_jid"]
                if "group_name" in data and data["group_name"]:
                    BRIDGE_STATE["group_name"] = data["group_name"]
            if data.get("connected"):
                add_bridge_log(f"Connected to WhatsApp ({data.get('self_jid')})")
            elif "connected" in data and not data.get("connected") and not data.get("qr_data_url"):
                add_bridge_log("WhatsApp connection closed, reconnecting...")

        elif self.path == "/internal/log":
            msg = data.get("message", "")
            if msg:
                add_bridge_log(msg)

        elif self.path == "/internal/outbox_peek":
            with STATE_LOCK:
                msgs = list(BRIDGE_STATE["outbox"][:1])
            response_payload = {"messages": msgs}

        elif self.path == "/internal/outbox_pop":
            with STATE_LOCK:
                if BRIDGE_STATE["outbox"]:
                    BRIDGE_STATE["outbox"].pop(0)
            delivered_to = data.get("delivered_to", "Mess Menu")
            chars = data.get("chars", 0)
            add_bridge_log(f"✅ Delivered outbound message ({chars} chars) -> {delivered_to}")

        elif self.path == "/internal/query":
            text = data.get("text", "")
            r_jid = data.get("remote_jid", "")
            menu_data = get_active_menu_data()
            reply = resolve_user_query(text, menu_data, datetime.now(TIMEZONE))
            if reply:
                add_bridge_log(f"Received query '{text}' in 'Mess Menu' ({r_jid}) -> Replied ({len(reply)} chars)")
            response_payload = {"reply": reply}

        elif self.path == "/internal/upload_xlsx":
            try:
                b64_content = data.get("b64", "")
                with open(MENU_FILE_PATH, "wb") as f:
                    f.write(base64.b64decode(b64_content))
                menu_data = parse_menu_excel(MENU_FILE_PATH)
                dates_info = ", ".join(
                    f"{d[:3].title()}: {menu_data['dates'].get(d, 'N/A')}"
                    for d in DAYS_ORDER if d in menu_data.get("dates", {})
                )
                add_bridge_log(f"Updated weekly Excel menu from 'Mess Menu' group ({dates_info})")
                response_payload = {
                    "reply": (
                        f"✅ *Weekly Menu Updated!*\n"
                        f"📅 Dates loaded: {dates_info or 'Mon–Sun'}\n\n"
                        f"Daily alerts set for 6:00 AM, 11:00 AM, 3:00 PM, and 7:00 PM IST in *Mess Menu*."
                    )
                }
            except Exception as exc:
                add_bridge_log(f"Excel upload error: {exc}")
                response_payload = {"reply": f"❌ Could not read Excel file: {exc}"}

        out = json.dumps(response_payload).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(out)))
        self.end_headers()
        self.wfile.write(out)


# ---------------------------------------------------------------------------
# CLI Entrypoint
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="BITSoM Mess Menu WhatsApp Bot (Mess Menu Group Only)")
    parser.add_argument("--file", help="Optional path to a .xlsx menu file to load")
    parser.add_argument("--preview", choices=DAYS_ORDER + ["ALL"], help="Preview formatted WhatsApp messages")
    parser.add_argument("--query", help="Test an on-demand query (e.g. 'today', 'thursday dinner')")
    args = parser.parse_args()

    if args.file and os.path.exists(args.file):
        if os.path.abspath(args.file) != os.path.abspath(MENU_FILE_PATH):
            shutil.copyfile(args.file, MENU_FILE_PATH)

    if args.preview:
        menu_data = get_active_menu_data()
        days_to_show = DAYS_ORDER if args.preview == "ALL" else [args.preview]
        for d in days_to_show:
            print(format_full_day_menu(menu_data, d))
            print("\n" + "=" * 40 + "\n")
        return

    if args.query:
        menu_data = get_active_menu_data()
        print(resolve_user_query(args.query, menu_data, datetime.now(TIMEZONE)))
        return

    # Start daily IST scheduler thread
    threading.Thread(target=scheduler_loop, daemon=True).start()

    # Start Node.js Baileys QR-Code bridge in background
    threading.Thread(target=start_node_qr_bridge, daemon=True).start()

    logger.info("Web Dashboard & QR page running on http://0.0.0.0:%d", PORT)
    ThreadingHTTPServer(("0.0.0.0", PORT), BotWebServer).serve_forever()


if __name__ == "__main__":
    main()

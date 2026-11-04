"""
Central configuration for the Hospitality Procurement Intelligence Dashboard.

All business assumptions used to generate the synthetic dataset live here so they
can be reviewed and tuned in one place. Every name below (hotel group, hotels,
suppliers) is fictional. The structure is modelled on how a multi-property Irish
hotel group buys food, beverage and operating supplies.
"""

from pathlib import Path

# --------------------------------------------------------------------------- #
# Paths
# --------------------------------------------------------------------------- #
PROJECT_ROOT = Path(__file__).resolve().parents[1]
RAW_DIR = PROJECT_ROOT / "data" / "raw"
PROCESSED_DIR = PROJECT_ROOT / "data" / "processed"
OUTPUT_DIR = PROJECT_ROOT / "output"
IMAGES_DIR = PROJECT_ROOT / "docs" / "images"

DASHBOARD_FILE = OUTPUT_DIR / "Procurement_Intelligence_Dashboard.xlsx"

# --------------------------------------------------------------------------- #
# Simulation window and reproducibility
# --------------------------------------------------------------------------- #
RANDOM_SEED = 42
START_DATE = "2024-10-01"
END_DATE = "2026-09-30"          # 24 full months

GROUP_NAME = "Harbourline Hotel Group (fictional)"

# --------------------------------------------------------------------------- #
# Hotels: id, name, city, rooms, compliance discipline (0-1, higher = better)
# --------------------------------------------------------------------------- #
HOTELS = [
    ("H01", "Harbourline Dublin City",    "Dublin",   320, 0.95),
    ("H02", "Harbourline Dublin West",    "Dublin",   245, 0.92),
    ("H03", "Harbourline Dublin Airport", "Dublin",   280, 0.90),
    ("H04", "Harbourline Cork",           "Cork",     210, 0.86),
    ("H05", "Harbourline Galway",         "Galway",   190, 0.80),
    ("H06", "Harbourline Limerick",       "Limerick", 160, 0.74),
]

# --------------------------------------------------------------------------- #
# Categories -> subcategories
# --------------------------------------------------------------------------- #
CATEGORIES = {
    "Food": ["Fresh Produce", "Meat & Poultry", "Seafood", "Dairy & Eggs",
             "Bakery", "Dry Goods"],
    "Beverage": ["Beer & Cider", "Wine", "Spirits", "Soft Drinks & Mixers",
                 "Tea & Coffee"],
    "Operating Supplies": ["Cleaning Chemicals", "Guest Amenities",
                           "Linen & Laundry", "Packaging & Disposables"],
}

# Annual price inflation applied to market prices, by category
CATEGORY_INFLATION = {
    "Food": 0.055,
    "Beverage": 0.040,
    "Operating Supplies": 0.030,
}

# --------------------------------------------------------------------------- #
# Suppliers
#   id, name, subcategories served, contracted?, lead time (days),
#   delivery reliability (prob. on time), invoice accuracy (prob. clean),
#   price discipline (avg. premium charged over contract on contract lines)
# --------------------------------------------------------------------------- #
SUPPLIERS = [
    ("S01", "Emerald Fresh Produce",     ["Fresh Produce"],                  True, 1, 0.94, 0.97, 0.000),
    ("S02", "Midlands Meat Co.",         ["Meat & Poultry"],                 True, 2, 0.91, 0.95, 0.012),
    ("S03", "Atlantic Seafood Traders",  ["Seafood"],                        True, 1, 0.86, 0.93, 0.025),
    ("S04", "Celtic Dairy Distributors", ["Dairy & Eggs"],                   True, 1, 0.96, 0.98, 0.000),
    ("S05", "Oakfield Craft Bakery",     ["Bakery"],                         True, 1, 0.95, 0.97, 0.004),
    ("S06", "Shannon Dry Goods",         ["Dry Goods"],                      True, 3, 0.93, 0.96, 0.008),
    ("S07", "Liffey Beverage Wholesale", ["Beer & Cider", "Soft Drinks & Mixers"], True, 2, 0.92, 0.96, 0.006),
    ("S08", "Tara Wines & Spirits",      ["Wine", "Spirits"],                True, 3, 0.89, 0.94, 0.018),
    ("S09", "Roastworks Coffee",         ["Tea & Coffee"],                   True, 4, 0.97, 0.99, 0.000),
    ("S10", "BrightClean Hygiene",       ["Cleaning Chemicals"],             True, 4, 0.90, 0.95, 0.010),
    ("S11", "Linden Guest Supplies",     ["Guest Amenities"],                True, 5, 0.88, 0.96, 0.005),
    ("S12", "LinenCare Ireland",         ["Linen & Laundry"],                True, 3, 0.85, 0.92, 0.020),
    ("S13", "Coastline Packaging",       ["Packaging & Disposables"],        True, 4, 0.92, 0.97, 0.003),
    # Off-contract / spot suppliers: used when sites buy outside agreements
    ("S14", "QuickServe Cash & Carry",   ["Fresh Produce", "Dairy & Eggs", "Dry Goods",
                                          "Soft Drinks & Mixers", "Packaging & Disposables"],
                                                                             False, 0, 0.98, 0.90, 0.000),
    ("S15", "Westside Local Butchers",   ["Meat & Poultry"],                 False, 1, 0.90, 0.88, 0.000),
    ("S16", "Corrib Fish Market",        ["Seafood"],                        False, 1, 0.84, 0.86, 0.000),
    ("S17", "Metro Drinks Direct",       ["Beer & Cider", "Wine", "Spirits"], False, 1, 0.87, 0.89, 0.000),
    ("S18", "Budget Janitorial Supply",  ["Cleaning Chemicals", "Guest Amenities"],
                                                                             False, 2, 0.82, 0.87, 0.000),
]

# How much more off-contract suppliers charge than contract price (mean, sd)
OFF_CONTRACT_PREMIUM = (0.14, 0.05)

# --------------------------------------------------------------------------- #
# Products: name, subcategory, unit, base contract price (EUR),
#           base weekly qty per 100 rooms
# --------------------------------------------------------------------------- #
PRODUCTS = [
    # Fresh Produce
    ("Potatoes (Rooster) 10kg",            "Fresh Produce", "sack", 9.80, 6.0),
    ("Mixed Salad Leaves 1kg",             "Fresh Produce", "bag", 6.40, 8.0),
    ("Tomatoes Vine 5kg",                  "Fresh Produce", "box", 14.20, 3.0),
    ("Onions Brown 10kg",                  "Fresh Produce", "sack", 8.60, 2.5),
    ("Seasonal Fruit Platter Mix 5kg",     "Fresh Produce", "box", 21.50, 3.5),
    ("Lemons (case of 80)",                "Fresh Produce", "case", 18.90, 1.2),
    # Meat & Poultry
    ("Irish Beef Sirloin 1kg",             "Meat & Poultry", "kg", 24.50, 9.0),
    ("Chicken Fillets 5kg",                "Meat & Poultry", "box", 38.00, 5.0),
    ("Back Rashers 2.27kg",                "Meat & Poultry", "pack", 17.60, 7.0),
    ("Pork Sausages (60)",                 "Meat & Poultry", "pack", 14.80, 6.5),
    ("Lamb Shank (each)",                  "Meat & Poultry", "each", 6.90, 10.0),
    # Seafood
    ("Salmon Fillet Side 1.5kg",           "Seafood", "side", 27.40, 3.0),
    ("Cod Loin 1kg",                       "Seafood", "kg", 22.80, 3.5),
    ("Smoked Salmon Sliced 1kg",           "Seafood", "pack", 31.50, 1.5),
    # Dairy & Eggs
    ("Free Range Eggs (tray of 30)",       "Dairy & Eggs", "tray", 7.40, 14.0),
    ("Fresh Milk 2L",                      "Dairy & Eggs", "bottle", 2.15, 40.0),
    ("Irish Butter Portions (100)",        "Dairy & Eggs", "box", 12.60, 5.0),
    ("Cream 1L",                           "Dairy & Eggs", "carton", 4.30, 6.0),
    ("Cheddar Block 2.5kg",                "Dairy & Eggs", "block", 21.90, 1.5),
    # Bakery
    ("Brown Soda Bread",                   "Bakery", "loaf", 2.60, 18.0),
    ("Croissants (pack of 24)",            "Bakery", "pack", 13.40, 6.0),
    ("Sliced Pan White 800g",              "Bakery", "loaf", 1.95, 12.0),
    # Dry Goods
    ("Porridge Oats 3kg",                  "Dry Goods", "bag", 6.20, 2.0),
    ("Basmati Rice 10kg",                  "Dry Goods", "sack", 19.80, 0.8),
    ("Penne Pasta 3kg",                    "Dry Goods", "bag", 6.90, 1.2),
    ("Rapeseed Oil 10L",                   "Dry Goods", "drum", 24.50, 1.0),
    ("Breakfast Cereal Variety (case)",    "Dry Goods", "case", 28.60, 1.5),
    # Beer & Cider
    ("Stout Keg 50L",                      "Beer & Cider", "keg", 225.00, 0.9),
    ("Lager Keg 50L",                      "Beer & Cider", "keg", 210.00, 1.1),
    ("Craft IPA Bottles (24)",             "Beer & Cider", "case", 42.00, 1.0),
    ("Irish Cider Bottles (24)",           "Beer & Cider", "case", 36.50, 0.8),
    # Wine
    ("House Red Wine (case of 6)",         "Wine", "case", 48.00, 3.0),
    ("House White Wine (case of 6)",       "Wine", "case", 46.50, 3.4),
    ("Prosecco (case of 6)",               "Wine", "case", 57.00, 1.6),
    ("Premium Rioja (case of 6)",          "Wine", "case", 84.00, 0.5),
    # Spirits
    ("Irish Whiskey 70cl",                 "Spirits", "bottle", 24.90, 2.5),
    ("Premium Gin 70cl",                   "Spirits", "bottle", 27.50, 2.2),
    ("Vodka 70cl",                         "Spirits", "bottle", 19.40, 1.8),
    ("Dark Rum 70cl",                      "Spirits", "bottle", 21.20, 0.9),
    # Soft Drinks & Mixers
    ("Cola Postmix 10L",                   "Soft Drinks & Mixers", "box", 62.00, 1.0),
    ("Tonic Water 200ml (24)",             "Soft Drinks & Mixers", "case", 13.80, 2.4),
    ("Still Water 500ml (24)",             "Soft Drinks & Mixers", "case", 8.90, 4.0),
    ("Orange Juice 1L (12)",               "Soft Drinks & Mixers", "case", 19.20, 2.6),
    # Tea & Coffee
    ("Coffee Beans House Blend 1kg",       "Tea & Coffee", "bag", 17.50, 4.0),
    ("Breakfast Tea Bags (1100)",          "Tea & Coffee", "box", 22.40, 0.7),
    # Cleaning Chemicals
    ("Multi-Surface Cleaner 5L",           "Cleaning Chemicals", "drum", 11.80, 2.0),
    ("Dishwasher Detergent 20L",           "Cleaning Chemicals", "drum", 46.00, 0.8),
    ("Washroom Sanitiser 5L",              "Cleaning Chemicals", "drum", 13.50, 1.5),
    # Guest Amenities
    ("Shampoo 30ml (case of 300)",         "Guest Amenities", "case", 54.00, 0.9),
    ("Wrapped Soap Bars (case of 500)",    "Guest Amenities", "case", 49.00, 0.6),
    ("Toilet Rolls (case of 36)",          "Guest Amenities", "case", 24.60, 2.8),
    # Linen & Laundry
    ("Bath Towel (each)",                  "Linen & Laundry", "each", 6.20, 6.0),
    ("King Duvet Cover (each)",            "Linen & Laundry", "each", 18.40, 1.8),
    ("Laundry Service (per 100kg)",        "Linen & Laundry", "100kg", 82.00, 3.2),
    # Packaging & Disposables
    ("Takeaway Coffee Cups 12oz (1000)",   "Packaging & Disposables", "case", 64.00, 0.4),
    ("Napkins 2-ply (case of 2000)",       "Packaging & Disposables", "case", 27.80, 0.9),
    ("Cling Film 450mm",                   "Packaging & Disposables", "roll", 12.40, 0.8),
]

# Subcategories whose stock is counted monthly in the bar store (stock variance)
STOCK_COUNTED_SUBCATEGORIES = ["Beer & Cider", "Wine", "Spirits"]

# Monthly demand seasonality for an Irish city hotel (Jan..Dec)
SEASONALITY = [0.80, 0.82, 0.90, 0.97, 1.02, 1.10,
               1.18, 1.20, 1.05, 0.98, 0.96, 1.22]

# Year-2 scenario: a group compliance push cuts off-contract buying at most sites,
# while one property slips. Multiplier on each hotel's off-contract probability
# from the start of contract year 2 (October 2025).
YEAR2_OFF_CONTRACT_MULTIPLIER = {
    "H01": 0.65, "H02": 0.70, "H03": 0.70, "H04": 0.80, "H05": 0.90, "H06": 1.15,
}

# Multiplier on all order quantities (calibrates total spend to a realistic level)
VOLUME_SCALE = 2.2

# Probability that a contracted supplier sneaks a price above contract
PRICE_CREEP_PROB = 0.03

# Probability an invoice line has a quantity short-delivery
SHORT_DELIVERY_PROB = 0.035

# Buyers placing orders (fictional)
BUYERS = ["Procurement Desk", "Head Chef", "Bar Manager", "Accommodation Manager",
          "Duty Manager"]

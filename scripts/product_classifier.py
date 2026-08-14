"""
product_classifier.py
Golden Opportunity Product Classification Engine
"""

import re


PRODUCT_TYPE_RULES = [
    {
        "category": "Kitchen Appliances",
        "subcategory": "Electric Burners",
        "phrases": [
            "electric double burner",
            "double burner",
            "electric burner",
            "countertop burner",
            "buffet burner",
            "hot plate",
        ],
    },
    {
        "category": "Kitchen Appliances",
        "subcategory": "Slow Cookers",
        "phrases": [
            "slow cooker",
            "crock pot",
        ],
    },
    {
        "category": "Kitchen Appliances",
        "subcategory": "Rice Cookers",
        "phrases": [
            "rice cooker",
        ],
    },
    {
        "category": "Kitchen Appliances",
        "subcategory": "Pressure Cookers",
        "phrases": [
            "pressure cooker",
            "pressure canner",
        ],
    },
    {
        "category": "Kitchen Appliances",
        "subcategory": "Coffee Makers",
        "phrases": [
            "coffee maker",
        ],
    },
    {
        "category": "Kitchen Appliances",
        "subcategory": "Espresso Makers",
        "phrases": [
            "espresso coffee maker",
            "espresso maker",
            "espresso machine",
            "moka",
        ],
    },
    {
        "category": "Kitchen Appliances",
        "subcategory": "Blenders",
        "phrases": [
            "personal blender",
            "blender",
        ],
    },
    {
        "category": "Electronics",
        "subcategory": "TV Mounts",
        "phrases": [
            "full motion wall mount",
            "tv wall mount",
            "wall mount",
        ],
    },
    {
        "category": "Electronics",
        "subcategory": "LED Lighting",
        "phrases": [
            "led strip lights",
            "led strip light",
            "led light strip",
        ],
    },
    {
        "category": "Home & Laundry",
        "subcategory": "Steam Irons",
        "phrases": [
            "steam/dry iron",
            "steam dry iron",
            "steam iron",
        ],
    },
    {
        "category": "Home & Laundry",
        "subcategory": "Garment Steamers",
        "phrases": [
            "garment steamer",
            "clothes steamer",
        ],
    },
]


def normalize(text):
    if not text:
        return ""

    text = text.lower()
    text = re.sub(r"[^a-z0-9 ]", " ", text)
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def classify_product(description):
    text = normalize(description)

    # Longest phrases first
    rules = sorted(
        PRODUCT_TYPE_RULES,
        key=lambda r: max(len(p) for p in r["phrases"]),
        reverse=True,
    )

    for rule in rules:
        for phrase in sorted(rule["phrases"], key=len, reverse=True):
            if normalize(phrase) in text:
                return {
                    "parent": rule["category"],
                    "sub": rule["subcategory"],
                }

    return None
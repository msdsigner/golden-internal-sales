import json
import xmlrpc.client
from pathlib import Path

ROOT = Path(__file__).resolve().parent

with open(ROOT / "config_odoo.json", "r", encoding="utf-8") as f:
    CONFIG = json.load(f)

ODOO_URL = CONFIG["odoo"]["url"]
DB = CONFIG["odoo"]["database"]
USERNAME = CONFIG["odoo"]["username"]
API_KEY = CONFIG["odoo"]["api_key"]

common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common")

uid = common.authenticate(DB, USERNAME, API_KEY, {})

models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object")

tags = models.execute_kw(
    DB,
    uid,
    API_KEY,
    "product.tag",
    "search_read",
    [[]],
    {
        "fields": ["id", "name"],
        "order": "id",
    },
)

print("\nProduct Tags:\n")

for tag in tags:
    print(tag)
import json
import os

BASE_DIR = os.path.dirname(__file__)
DATA_DIR = os.path.abspath(os.path.join(BASE_DIR, "../data"))

def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def search_melc(grade, query):
    file_map = {
        "7": "melc_g7.json",
        "8": "melc_g8.json",
        "9": "melc_g9.json",
        "10": "melc_g10.json",
        "11-12": "melc_g11_12.json"
    }

    filename = file_map.get(grade)
    if not filename:
        return []

    path = os.path.join(DATA_DIR, "melc", filename)
    data = load_json(path)

    results = []
    for item in data:
        if query.lower() in item["competency"].lower():
            results.append(item["competency"])

    return results[:5]

def search_matatag(query):
    path = os.path.join(DATA_DIR, "matatag", "matatag_all_raw.json")
    data = load_json(path)

    results = []
    for item in data:
        if query.lower() in item["competency"].lower():
            results.append(item["competency"])

    return results[:5]
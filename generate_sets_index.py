import os
import json

CATEGORY_ORDER = [
    "Starter Deck",
    "Booster Set",
    "Community Collection",
    "Exclusive",
    "Selection",
    "PRE0",
    "PromoCard"
]

def category_priority(path):
    for i, cat in enumerate(CATEGORY_ORDER):
        if cat.lower() in path.lower():
            return i
    return len(CATEGORY_ORDER)

def generate_sets_index(base_path="."):
    sets = []

    for root, dirs, files in os.walk(base_path):
        if "info.json" in files:
            info_path = os.path.join(root, "info.json")
            rel_path = os.path.relpath(info_path, base_path).replace("\\", "/")
            folder_name = os.path.basename(root)
            sets.append({
                "name": folder_name,
                "path": rel_path
            })

    sets.sort(key=lambda x: (category_priority(x["path"]), x["name"].lower()))

    sets_index = {"sets": sets}
    output_path = os.path.join(base_path, "sets_index.json")
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(sets_index, f, ensure_ascii=False, indent=2)

    print(f"✅ สร้างไฟล์ sets_index.json สำเร็จ ({len(sets)} ชุด)")
    print(f"📁 ตำแหน่งไฟล์: {output_path}")

if __name__ == "__main__":
    generate_sets_index()

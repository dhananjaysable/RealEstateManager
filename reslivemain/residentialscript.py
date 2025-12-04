import pandas as pd
import re
import sys
import unicodedata
from datetime import datetime
import os

# === Helper to clean description ===
def clean_description(text):
    text = str(text)
    text = re.sub(r"\b[A-Za-z]+(/[A-Za-z]+)+\b", "", text)
    text = re.sub(r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b", "", text)
    text = re.sub(r"=\s*\d+\.?\d*\s*(चौ\.?\s*फु\.?|चौ\.?\s*फूट|चौ\s*फु|चौ\s*फूट)", "", text)
    return text

# === Unicode normalization helper ===
def normalize_marathi(text: str) -> str:
    """Normalize Marathi text by removing invisible and punctuation characters."""
    if not isinstance(text, str):
        return ""
    text = unicodedata.normalize("NFKD", text)
    text = re.sub(r"[\s\u200b\u200c\u200d\u00a0\.\-]+", "", text)
    return text.strip().lower()

# === Area pattern ===
AREA_PATTERN = r"(\d+\.?\d*)\s*(चौ\.?\s*फु\.?|चौ\s*फु|चौ\.?\s*फू\.?|चौ\s*फू|चौ\.?\s*फूट|चौफुट|चौ\s*फुट|चौ\.?\s*फुटात|चौ\s*फुटात)"

# === Shared parsing logic ===
def parse_contextual_areas(description, total_from_column):
    """Parse multiple contextual areas like RCC + Parking + Open etc."""
    desc_clean = clean_description(description)
    total_area = 0.0
    raw_patterns = []
    RCC = C = E = PR = OP = 0.0

    area_matches = list(re.finditer(AREA_PATTERN, desc_clean))
    for match in area_matches:
        num = float(match.group(1))
        total_area += num
        start_idx = match.start()
        context_start = max(0, start_idx - 60)
        context = desc_clean[context_start:start_idx]

        # detect context-sensitive allocations
        if re.search(r"पार्किंग|parking", context, re.IGNORECASE):
            PR += num
        elif re.search(r"आर\s*\.?\s*सी\s*\.?\s*सी|rcc|निवासी", context):
            RCC += num
        elif re.search(r"पत्रा|पत्रा\s*शेड|सिमेंट\s*पत्रा", context):
            E += num
        elif re.search(r"कच्ची\s*पक्की|साधे\s*शेड", context):
            C += num
        elif re.search(r"मोकळी\s*जागा|ओपन\s*स्पेस", context):
            OP += num
        else:
            # default RCC if context is unknown
            RCC += num

        raw_patterns.append(match.group(0).strip())

    final_total = total_from_column if total_from_column > 0 else total_area
    assigned = RCC + C + E + PR + OP
    if final_total > assigned:
        RCC += final_total - assigned

    return ", ".join(raw_patterns) if raw_patterns else None, final_total, RCC, PR, C, E, OP


# === 2️⃣ Main logic ===
def extract_area(description, totalarea, construction_type, unmatched_types):
    description = str(description).strip()
    total_from_column = float(totalarea) if pd.notna(totalarea) else 0.0
    ctype = normalize_marathi(construction_type)

    # 🔧 Calculate all L*B patterns first
    desc_clean = clean_description(description)
    lb_matches = re.findall(r"(\d+\.?\d*)\s*[*xX]\s*(\d+\.?\d*)", desc_clean)
    total_lb_area = sum(float(l) * float(b) for l, b in lb_matches)

    # 🧩 CASE 1: मिश्र OR description contains "पार्किंग" → contextual parse
    if str(construction_type).strip() == "मिश्र" or re.search(r"पार्किंग|parking", description, re.IGNORECASE):
        return parse_contextual_areas(description, total_from_column or total_lb_area)

    # 🧩 CASE 2: Non-मिश्र → direct classification
    raw_patterns = [f"{l}*{b}" for l, b in lb_matches]
    area_matches = list(re.finditer(AREA_PATTERN, description))
    for m in area_matches:
        raw_patterns.append(m.group(0))

    # 🔧 use either Excel totalarea or calculated L×B
    final_total = total_from_column if total_from_column > 0 else total_lb_area

    RCC = C = E = PR = OP = 0.0

    # RCC
    if re.search(r"(आरसीसीकिंवालोडबेअरिंग|आरसीसीशेडकिंवाँऑफीस|आरसीसीकिंवालोडबेअरिंगफ्लटसिस्टिमइमारतवचाळ|rcc)", ctype):
        RCC = final_total
    # C
    elif "कच्चीपक्कीवीटमातीचीछतपत्र्याचेवगवताचेधाब्याचे" in ctype or "साधेशेडकिंवाँऑफीस" in ctype:
        C = final_total
    # E
    elif "पत्र्याचीटेम्पररीशेड्स" in ctype:
        E = final_total
    # PR
    elif "पार्किंगएरीया" in ctype:
        PR = final_total
    # OP
    elif "मोकळ्याजमिन" in ctype:
        OP = final_total
    else:
        unmatched_types.add(str(construction_type))
        RCC = final_total  # default RCC

    return ", ".join(raw_patterns) if raw_patterns else None, final_total, RCC, PR, C, E, OP


def process_residential_data(file_path, log_callback=None):
    def log(msg):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    log("🏗️ Starting Real Estate Data Cleaning (Final v8 with RCC + Parking Split)...")
    
    if not file_path:
        log("❌ No file provided.")
        return

    log(f"📂 Reading file: {file_path}")

    try:
        df = pd.read_excel(file_path, engine="openpyxl")
    except Exception as e:
        log(f"❌ Error reading file: {e}")
        return

    total_rows = len(df)
    log(f"📊 Total rows to process: {total_rows}")

    # === 3️⃣ Process all rows ===
    raw_texts, areas, RCCs, PRs, Cs, Es, OPs = [], [], [], [], [], [], []
    unmatched_types = set()

    for idx, row in df.iterrows():
        raw, area, rcc, pr, c, e, op = extract_area(
            row.get("description", ""),
            row.get("totalarea", 0),
            row.get("finalconstructiontype", ""),
            unmatched_types
        )
        raw_texts.append(raw)
        areas.append(area)
        RCCs.append(rcc)
        PRs.append(pr)
        Cs.append(c)
        Es.append(e)
        OPs.append(op)

        if (idx + 1) % 2000 == 0:
            log(f"✅ Processed {idx + 1}/{total_rows} rows...")

    # === 4️⃣ Add results ===
    df["Raw_Area_Text"] = raw_texts
    df["Area_R"] = areas
    df["RCC"] = RCCs
    df["PR"] = PRs
    df["C"] = Cs
    df["E"] = Es
    df["OP"] = OPs

    # === 5️⃣ Output ===
    output_dir = os.path.dirname(file_path)
    output_file = os.path.join(output_dir, f"Residential_bifurcation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    
    try:
        df.to_excel(output_file, index=False)
        log(f"🎉 Cleaning complete! (Smart Split for RCC + Parking + मिश्र + L×B)")
        log(f"📁 Output saved as: {output_file}")
    except Exception as e:
        log(f"❌ Error saving file: {e}")
        return

    # === 6️⃣ Write unmatched safely ===
    if unmatched_types:
        unmatched_clean = [str(u) for u in unmatched_types if isinstance(u, str) and u.strip()]
        unmatched_clean = sorted(list(set(unmatched_clean)))
        unmatched_file = os.path.join(output_dir, "unmatched_construction_types.txt")
        with open(unmatched_file, "w", encoding="utf-8") as f:
            f.write("\n".join(unmatched_clean))
        log(f"⚠️ {len(unmatched_clean)} unmatched construction types written to {unmatched_file}")

    return output_file

if __name__ == "__main__":
    file_path = sys.argv[1] if len(sys.argv) > 1 else "input.xlsx"
    process_residential_data(file_path)

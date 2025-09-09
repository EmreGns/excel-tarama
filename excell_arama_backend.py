
# Gerekli kütüphaneler ve test edilmiş sürümler:
# Python: 3.13.7
# pandas: 2.3.2
# openpyxl: 3.1.5
# rapidfuzz: 3.14.0

import pandas as pd
from pathlib import Path
from rapidfuzz import fuzz
import os
import re

# Kısaltma sözlüğü (Bu kısma yeni kıslatmalar ekleyebilirsiniz)
abbreviations = {
    "P-": "Pressure",
    "T-": "Temperature",
    "HI": "High",
    "LO": "Low",
    "HIHI": "High Warning",
    "LOLO": "Low Alarm",
    "FIIL": "Filter",
    "Ctrl": "Control",
    "P-Ctrl": "Pressure Control",
    "RL": "Return Line",
    "INLT": "Inlet",
    "OUTLT": "Outlet",
    "LUBE": "Lubrication",
    "GEAR": "Shaft",
    "Upstr": "Upstream",
    "WF": "Wrong feedback",
}

def normalize_text(text):
    if not isinstance(text, str):
        return ""
    for abbr, full in abbreviations.items():
        text = text.replace(abbr, full)
    text = text.lower()
    # Noktalama ve özel karakterleri boşlukla değiştir
    text = re.sub(r'[\-\.,()\[\]/]', ' ', text)
    # Fazla boşlukları tek boşluk yap
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def extract_all_text_from_excel(file_path):
    try:
        # Tüm sayfaları oku
        df_all = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
    except Exception as e:
        print(f"⛔ Dosya okunamadı: {file_path.name} → {e}")
        return []

    matched_cells = []

    for sheet_name, df in df_all.items():
        for row_idx in range(len(df)):
            row = df.iloc[row_idx]
            for col_idx in range(len(row)):
                cell_value = row.iloc[col_idx]

                if pd.isna(cell_value):
                    continue

                cell_text = str(cell_value).strip()
                norm_text = normalize_text(cell_text)

                matched_cells.append({
                    "row_num": row_idx + 2,  # Excel satır numarası (1 başlık için)
                    "col_idx": col_idx,
                    "row": row,
                    "cell_text": cell_text,
                    "normalized": norm_text,
                })

    return matched_cells

def search_excel_files(search_term, folder_path, case=1, offset=2):
    search_term_norm = normalize_text(search_term)
    results = []
    seen = set()

    excel_files = list(Path(folder_path).glob("*.xlsx"))

    for file in excel_files:
        if file.name.startswith('~$'):
            continue

        print(f"📂 Taranıyor: {file.name}")
        cells = extract_all_text_from_excel(file)

        for cell in cells:
            if len(cell["normalized"]) < 5:
                continue
            if cell["normalized"] in seen:
                continue

            similarity = fuzz.token_set_ratio(search_term_norm, cell["normalized"])

            if similarity > 60:
                row = cell["row"]
                col = cell["col_idx"]
                row_num = cell["row_num"]

                if case == 1:
                    # Tüm satırı göster
                    texts = [str(c).strip() if not pd.isna(c) else "" for c in row]
                    final_text = " | ".join(texts)

                elif case == 2:
                    # Sağa doğru offset kadar hücre
                    end = min(col + offset + 1, len(row))
                    selected = row.iloc[col:end]
                    texts = [str(c).strip() if not pd.isna(c) else "" for c in selected]
                    final_text = " | ".join(texts)

                elif case == 3:
                    # Sola doğru offset kadar hücre
                    start = max(0, col - offset)
                    selected = row.iloc[start:col + 1]
                    texts = [str(c).strip() if not pd.isna(c) else "" for c in selected]
                    final_text = " | ".join(texts)

                elif case == 4:
                    # Satırın başından offset kadar hücre
                    selected = row.iloc[:offset]
                    texts = [str(c).strip() if not pd.isna(c) else "" for c in selected]
                    final_text = " | ".join(texts)

                else:
                    final_text = str(row.iloc[col])

                results.append((similarity, file.name, f"Satır {row_num}: {final_text}"))
                seen.add(cell["normalized"])

    results.sort(reverse=True)
    return results


def main():
    print("📁 Excel Arama Aracı - Gelişmiş Fuzzy Search\n")

    folder = input("📂 Excel dosyalarının bulunduğu klasör yolu: ").strip()
    if not os.path.exists(folder):
        print("⛔ Klasör bulunamadı.")
        return

    while True:
        search = input("\n🔍 Aranacak terimi girin (Çıkmak için 'q'): ").strip()
        if search.lower() == 'q':
            print("👋 Programdan çıkılıyor.")
            break

        print("\n🎯 Gösterim türünü seçin:")
        print("1 - Tüm satırı göster")
        print("2 - Sola doğru x hücre göster")
        print("3 - Sağa doğru x hücre göster")
        print("4 - Satırın başından x hücre göster")

        try:
            case = int(input("👉 Seçiminiz (1/2/3/4): ").strip())
            if case not in [1, 2, 3, 4]:
                print("❌ Geçersiz seçim. Varsayılan: 1")
                case = 1
        except:
            case = 1

        offset = 2
        if case in [2, 3, 4]:
            try:
                offset = int(input("🔢 Kaç hücre gösterilsin?: ").strip())
            except:
                print("❌ Geçersiz sayı, varsayılan 2 kullanılacak.")

        matches = search_excel_files(search, folder, case=case, offset=offset)

        if not matches:
            print("❌ Hiçbir eşleşme bulunamadı.")
        else:
            print("\n--- 🔎 En Benzer Sonuçlar ---\n")
            for score, file, content in matches[:10]:
                print(f"[{score:.1f}%] {file} → {content}")


if __name__ == "__main__":
    main()

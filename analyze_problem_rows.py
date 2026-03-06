import docx

doc = docx.Document('uploads/ВЯЗНИКОВСКАЯ просчет.docx')
table = doc.tables[0]

print("=== ПРОБЛЕМНЫЕ СТРОКИ (10-15) ===\n")

for row_idx in [10, 11, 12, 13, 14, 15]:
    if row_idx < len(table.rows):
        print(f"--- Строка {row_idx} ---")
        row = table.rows[row_idx]
        for col_idx, cell in enumerate(row.cells):
            text = cell.text.strip()
            if len(text) > 80:
                text = text[:80] + '...'
            print(f"  Столбец {col_idx}: '{text}'")
        print()

# REPLACE_MANY – Excel VBA Toolkit

A high-performance Excel VBA solution for **bulk full-word replacement** using a mapping table. It replaces strings precisely without affecting partial matches (e.g., replaces "Cat" but ignores "Category").

## ✨ Features

* **Full-Word Matching:** Prevents accidental partial replacements.
* **Mapping-Driven:** Uses a simple "From | To" table for logic.
* **High Performance:** Array-based processing (works in memory, not cell-by-cell).
* **Flexible Scope:** Works on Selected Ranges, Active Sheets, or the Entire Workbook.
* **Two Modes:**
1. **UDF:** `=REPLACE_MANY()` for dynamic, non-destructive results.
2. **Macro:** `REPLACE_MANY_POPUP` for permanent in-place replacements.

---

## 📌 Mapping Format

| From | To |
| --- | --- |
| old | new |
| foo | bar |

**Rules:**

* **Length Priority:** Matches longer keys first to prevent "clobbering."
* **First Win:** If duplicate keys exist, the first occurrence is used.
* **Safety:** Blank keys in the mapping table are automatically ignored.

---

## 🧩 Usage

### 1. Worksheet Function (UDF)

Ideal for keeping original data intact. The result will spill into adjacent cells in Excel 365.

```excel
=REPLACE_MANY(Data_Range, Mapping_Range)

```

### 2. In-Place Macro

Ideal for bulk-cleaning existing datasets.

1. Press `ALT + F8`.
2. Run **`REPLACE_MANY_POPUP`**.
3. Follow the prompts to select your map, scope, and case-sensitivity settings.

---

## 📥 Installation

1. Open Excel and press `ALT + F11`.
2. Go to **File → Import File...**
3. Select **`REPLACE_MANY.bas`**.
4. Save your workbook as **.xlsm** or **.xlsb**.

---

## 📄 License

MIT License – Free to use and modify.

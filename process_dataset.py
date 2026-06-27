"""
Assessment 1: English-Hindi Dataset Processing and Analysis
-------------------------------------------------------------
Loads parallel eng.txt / hin.txt (line-aligned sentence pairs), computes
word and character counts, applies the required filters, and exports
a cleaned dataset to Excel.

Usage:
    python process_dataset.py
"""

import pandas as pd

ENG_FILE = "eng.txt"
HIN_FILE = "hin.txt"
OUTPUT_FILE = "cleaned_english_hindi_dataset.xlsx"

WORD_COUNT_MIN, WORD_COUNT_MAX = 3, 60
WORD_DIFF_MIN, WORD_DIFF_MAX = -10, 10


def load_parallel_lines(eng_path, hin_path):
    """Read line-aligned English and Hindi sentence files."""
    with open(eng_path, "r", encoding="utf-8") as f:
        eng_lines = [line.strip() for line in f.readlines()]
    with open(hin_path, "r", encoding="utf-8") as f:
        hin_lines = [line.strip() for line in f.readlines()]

    if len(eng_lines) != len(hin_lines):
        raise ValueError(
            f"Line count mismatch: eng={len(eng_lines)}, hin={len(hin_lines)}"
        )
    return eng_lines, hin_lines


def word_count(sentence: str) -> int:
    """Whitespace-based word count (works for English; Hindi is space-delimited too)."""
    return len(sentence.split())


def char_count(sentence: str) -> int:
    """Character count excluding leading/trailing whitespace, counting all chars including spaces within."""
    return len(sentence)


def main():
    eng_lines, hin_lines = load_parallel_lines(ENG_FILE, HIN_FILE)
    print(f"Loaded {len(eng_lines)} parallel sentence pairs.")

    df = pd.DataFrame({
        "English Sentence": eng_lines,
        "Hindi Sentence": hin_lines,
    })

    # Drop empty rows (blank lines on either side)
    before = len(df)
    df = df[(df["English Sentence"].str.len() > 0) & (df["Hindi Sentence"].str.len() > 0)]
    print(f"Dropped {before - len(df)} empty rows -> {len(df)} remaining.")

    # Step 3: Word Count Analysis
    df["Word Count (English)"] = df["English Sentence"].apply(word_count)
    df["Word Count (Hindi)"] = df["Hindi Sentence"].apply(word_count)

    mask_word_range = (
        df["Word Count (English)"].between(WORD_COUNT_MIN, WORD_COUNT_MAX)
        & df["Word Count (Hindi)"].between(WORD_COUNT_MIN, WORD_COUNT_MAX)
    )
    before = len(df)
    df = df[mask_word_range]
    print(f"Word-count range filter (3-60, both languages): {before} -> {len(df)} rows.")

    # Step 4: Word Count Difference Calculation
    df["Difference between Word Count (English) and Word Count (Hindi)"] = (
        df["Word Count (English)"] - df["Word Count (Hindi)"]
    )
    mask_diff_range = df[
        "Difference between Word Count (English) and Word Count (Hindi)"
    ].between(WORD_DIFF_MIN, WORD_DIFF_MAX)
    before = len(df)
    df = df[mask_diff_range]
    print(f"Word-count difference filter (-10 to +10): {before} -> {len(df)} rows.")

    # Character counts
    df["Character Count (English)"] = df["English Sentence"].apply(char_count)
    df["Character Count (Hindi)"] = df["Hindi Sentence"].apply(char_count)
    df["Difference between Character Count (English) and Character Count (Hindi)"] = (
        df["Character Count (English)"] - df["Character Count (Hindi)"]
    )

    # Step 5: Final Output - reorder columns as per spec
    final_columns = [
        "English Sentence",
        "Hindi Sentence",
        "Word Count (English)",
        "Word Count (Hindi)",
        "Difference between Word Count (English) and Word Count (Hindi)",
        "Character Count (English)",
        "Character Count (Hindi)",
        "Difference between Character Count (English) and Character Count (Hindi)",
    ]
    df = df[final_columns].reset_index(drop=True)

    df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
    print(f"\nFinal cleaned dataset: {len(df)} rows.")
    print(f"Saved to: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()

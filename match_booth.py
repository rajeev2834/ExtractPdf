import pandas as pd
from rapidfuzz import process, fuzz

# List of common words to ignore
STOPWORDS = {
    "UTKRAMIT", "MADHYA", "VIDYALAY", "PRATHAMIK", "PRATHMIK",
    "NAVSRIJIT", "VIDHYALAY", "URDU", "PURVI", "PASCHIMI", 
    "DAKSHINI", "UTARI", "BHAG"
}

def clean_name(name: str) -> str:
    """Remove common words and normalize the school name"""
    if not isinstance(name, str):
        return ""
    name = name.upper()
    tokens = [t for t in name.replace(",", "").split() if t not in STOPWORDS]
    return " ".join(tokens)

def fuzzy_match(source_file, target_file, output_file, threshold=85):
    # Read Excel files
    source_df = pd.read_excel(source_file)
    target_df = pd.read_excel(target_file)

    source_col = "BOOTH_LOC"   # Column in source file
    target_col = "Part Name"   # Column in target file

    # Pre-clean source names
    source_cleaned = source_df[source_col].astype(str).map(clean_name).tolist()
    source_original = source_df[source_col].astype(str).tolist()

    matches = []
    for name in target_df[target_col]:
        clean = clean_name(name)
        if not clean:
            matches.append(["", 0])
            continue

        # Match against source list instead of target list
        match, score, idx = process.extractOne(
            clean,
            source_cleaned,
            scorer=fuzz.token_sort_ratio
        )

        if score >= threshold:
            matches.append([source_original[idx], score])
        else:
            matches.append(["No Match", score])

    # Write results into target dataframe
    target_df["MatchedName"] = [m[0] for m in matches]
    target_df["Score"] = [m[1] for m in matches]

    target_df.to_excel(output_file, index=False)
    print(f"✅ Matching completed! Output saved to {output_file}")


# Example usage
fuzzy_match("BoothList_2020.xlsx", "BoothList_2025.xlsx", "matched_output.xlsx", threshold=85)
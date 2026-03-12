#!/usr/bin/env python3
"""
Excel Data Cleansing Script
Removes rows where BOTH final_value AND original_value are NULL/empty/NaN
Original file is preserved, cleansed output is written to a new file
"""

import pandas as pd
from pathlib import Path
from datetime import datetime
import os

def cleanse_excel_file(input_file: str, output_file: str = None) -> pd.DataFrame:
    """
    Cleanse Excel file by removing rows with BOTH final_value AND original_value NULL

    Args:
        input_file: Path to input Excel file
        output_file: Path to output Excel file (if None, generates based on timestamp)

    Returns:
        Cleansed DataFrame
    """
    input_path = Path(input_file)

    # Read the Excel file
    print(f"📖 Reading file: {input_path}")
    df = pd.read_excel(input_path)
    print(f"📊 Original shape: {df.shape[0]:,} rows × {df.shape[1]} columns")

    # Format datetime strings in final_value and original_value columns
    print(f"📅 Formatting datetime strings...")
    def format_datetime(value):
        """Format datetime string from YYYY-MM-DD HH:MM:SS to YYYY-MM-DD"""
        if pd.isna(value):
            return value
        val_str = str(value).strip()
        # Check if it matches datetime pattern YYYY-MM-DD HH:MM:SS
        if len(val_str) == 19 and val_str[4] == '-' and val_str[7] == '-' and val_str[10] == ' ':
            try:
                # Try to parse and format
                return val_str[:10]  # Just take the date part
            except:
                return val_str
        return val_str

    df['final_value'] = df['final_value'].apply(format_datetime)
    df['original_value'] = df['original_value'].apply(format_datetime)
    print(f"   ✅ Datetime strings formatted")

    # Classify form types based on img_path column
    print(f"📋 Classifying form types...")
    def classify_form_type(img_path):
        """Classify form type based on img_path value"""
        if pd.isna(img_path):
            return 'Unknown'
        img_str = str(img_path)
        if 'CLMHK1MEDIC1' in img_str or 'CLMHK1HOSCLM' in img_str:
            return 'Claim Form I'
        elif 'CLMHK1MEDIC2' in img_str:
            return 'Claim Form II'
        else:
            return 'Unknown'

    df['form_type'] = df['img_path'].apply(classify_form_type)

    # Print form type distribution
    form_type_counts = df['form_type'].value_counts()
    print(f"   - Claim Form I: {form_type_counts.get('Claim Form I', 0):,} rows")
    print(f"   - Claim Form II: {form_type_counts.get('Claim Form II', 0):,} rows")
    print(f"   - Unknown: {form_type_counts.get('Unknown', 0):,} rows")

    # Check NULL values in key columns
    null_final = df['final_value'].isna().sum()
    null_original = df['original_value'].isna().sum()
    print(f"⚠️  NULL values found:")
    print(f"   - final_value: {null_final:,} rows")
    print(f"   - original_value: {null_original:,} rows")

    # Remove rows where BOTH final_value AND original_value are NULL
    # This includes NaN, None, or empty string
    df_cleaned = df.dropna(subset=['final_value', 'original_value'], how='all')

    # Also remove empty string values
    df_cleaned = df_cleaned[
        (df_cleaned['final_value'].str.strip() != '') &
        (df_cleaned['original_value'].str.strip() != '')
    ]

    rows_removed = len(df) - len(df_cleaned)
    print(f"✅ Cleansed shape: {df_cleaned.shape[0]:,} rows × {df_cleaned.shape[1]} columns")
    print(f"🗑️  Rows removed: {rows_removed:,}")

    # Generate output filename if not provided
    if output_file is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"{input_path.stem}_cleansed_{timestamp}.xlsx"

    output_path = Path(output_file)

    # Save cleansed data to new file (original file remains unchanged)
    print(f"💾 Saving cleansed data to: {output_path}")
    df_cleaned.to_excel(output_path, index=False)

    print(f"🎉 Cleansing completed successfully!")
    print(f"📁 Original file preserved: {input_path}")

    return df_cleaned


def main():
    """Main function to run the cleansing process"""

    # Configuration
    # Get the folder where THIS script is saved
    base_path = Path(__file__).parent
    # Join that folder with your filename
    INPUT_FILE = str(base_path / "Sample_Comparison_data_20261103.xlsx")

    print("=" * 60)
    print("🧹 Excel Data Cleansing")
    print("=" * 60)
    print()

    # Perform cleansing
    df_cleaned = cleanse_excel_file(
        input_file=INPUT_FILE,
        output_file=None  # Auto-generate output filename
    )

    print()
    print("=" * 60)
    print("Summary:")
    print(f"  Original file: {INPUT_FILE}")
    print(f"  Cleansed file: {Path(INPUT_FILE).stem}_cleansed_*.xlsx")
    print("=" * 60)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
Excel Field Accuracy Analysis Script
Calculates accuracy rate for each field by comparing final_value vs original_value
Takes into account is_edited column - if is_edited=1, it means OCR was incorrect
Generates summary report in a separate Excel file
"""

import pandas as pd
from pathlib import Path
from datetime import datetime


def analyze_accuracy(input_file: str, output_file: str = None) -> pd.DataFrame:
    """
    Analyze accuracy rate for each field by comparing final_value vs original_value
    Considers is_edited column: if is_edited=1, OCR was incorrect (count as inaccurate)

    Args:
        input_file: Path to input Excel file (cleansed data)
        output_file: Path to output Excel file (if None, generates based on timestamp)

    Returns:
        DataFrame with accuracy summary by field
    """
    input_path = Path(input_file)

    # Read the Excel file
    print(f"📖 Reading file: {input_path}")
    df = pd.read_excel(input_file)
    print(f"📊 Total records: {len(df):,}")
    print()

    # Check is_edited distribution
    edited_count = (df['is_edited'] == 1).sum()
    not_edited_count = (df['is_edited'] == 0).sum()
    print("=== is_edited Status ===")
    print(f"Not edited (is_edited=0): {not_edited_count:,} ({not_edited_count/len(df)*100:.2f}%)")
    print(f"Edited (is_edited=1): {edited_count:,} ({edited_count/len(df)*100:.2f}%)")
    print()

    # Calculate accuracy for each row considering is_edited
    # Accurate if: is_edited == 0 AND final_value == original_value
    # Inaccurate if: is_edited == 1 OR final_value != original_value
    df['is_accurate'] = (
        (df['is_edited'] == 0) &
        (df['final_value'].astype(str).str.strip() == df['original_value'].astype(str).str.strip())
    )

    # Overall accuracy (true OCR accuracy)
    overall_accuracy = df['is_accurate'].sum() / len(df) * 100
    accurate_count = df['is_accurate'].sum()
    inaccurate_count = len(df) - accurate_count

    # Breakdown inaccuracy causes
    inaccurate_edited = ((df['is_edited'] == 1) & (~df['is_accurate'])).sum()  # is_edited=1 (OCR wrong)
    inaccurate_not_edited = ((df['is_edited'] == 0) & (~df['is_accurate'])).sum()  # is_edited=0 but values differ

    print("=== Overall Accuracy (True OCR Accuracy) ===")
    print(f"Total records: {len(df):,}")
    print(f"Accurate: {accurate_count:,} ({overall_accuracy:.2f}%)")
    print(f"Inaccurate: {inaccurate_count:,} ({100-overall_accuracy:.2f}%)")
    print(f"  - Inaccurate due to is_edited=1 (OCR failed): {inaccurate_edited:,}")
    print(f"  - Inaccurate due to value mismatch (is_edited=0): {inaccurate_not_edited:,}")
    print()

    # Group by field_name and calculate accuracy per field
    print("=== Calculating accuracy per field ===")
    field_stats = df.groupby('field_name').agg(
        total_records=('is_accurate', 'count'),
        accurate_count=('is_accurate', 'sum'),
        inaccurate_count=('is_accurate', lambda x: x.count() - x.sum()),
        edited_count=('is_edited', lambda x: (x == 1).sum()),
    ).reset_index()

    # Calculate accuracy percentage
    field_stats['accuracy_rate'] = (field_stats['accurate_count'] / field_stats['total_records'] * 100).round(2)
    field_stats['inaccuracy_rate'] = (field_stats['inaccurate_count'] / field_stats['total_records'] * 100).round(2)
    field_stats['edited_rate'] = (field_stats['edited_count'] / field_stats['total_records'] * 100).round(2)

    # Sort by accuracy rate (descending)
    field_stats = field_stats.sort_values('accuracy_rate', ascending=False)

    # Reorder columns
    field_stats = field_stats[[
        'field_name',
        'total_records',
        'accurate_count',
        'inaccurate_count',
        'edited_count',
        'accuracy_rate',
        'inaccuracy_rate',
        'edited_rate'
    ]]

    # Print summary table
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    print(field_stats.to_string(index=False))
    print()

    # Overall statistics
    avg_accuracy = field_stats['accuracy_rate'].mean()
    median_accuracy = field_stats['accuracy_rate'].median()
    avg_edited_rate = field_stats['edited_rate'].mean()

    print("=== Field Accuracy Statistics ===")
    print(f"Average accuracy across fields: {avg_accuracy:.2f}%")
    print(f"Median accuracy across fields: {median_accuracy:.2f}%")
    print(f"Average edited rate across fields: {avg_edited_rate:.2f}%")
    print(f"Highest accuracy: {field_stats['accuracy_rate'].max():.2f}% ({field_stats.iloc[0]['field_name']})")
    print(f"Lowest accuracy: {field_stats['accuracy_rate'].min():.2f}% ({field_stats.iloc[-1]['field_name']})")
    print()

    # Generate output filename if not provided
    if output_file is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"{input_path.stem}_accuracy_summary_{timestamp}.xlsx"

    output_path = Path(output_file)

    # Save summary to Excel with multiple sheets
    print(f"💾 Saving accuracy summary to: {output_path}")

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Sheet 1: Field Accuracy Summary
        field_stats.to_excel(writer, sheet_name='Field_Accuracy', index=False)

        # Sheet 2: Overall Statistics
        overall_stats_df = pd.DataFrame({
            'Metric': ['Total Records', 'Accurate Records', 'Inaccurate Records',
                       'Overall Accuracy (%)', 'Average Field Accuracy (%)', 'Median Field Accuracy (%)',
                       'Average Edited Rate (%)', 'Highest Accuracy (%)', 'Highest Accuracy Field',
                       'Lowest Accuracy (%)', 'Lowest Accuracy Field',
                       'Inaccurate due to is_edited=1', 'Inaccurate due to value mismatch'],
            'Value': [len(df), accurate_count, inaccurate_count,
                      f"{overall_accuracy:.2f}", f"{avg_accuracy:.2f}", f"{median_accuracy:.2f}",
                      f"{avg_edited_rate:.2f}", f"{field_stats['accuracy_rate'].max():.2f}", field_stats.iloc[0]['field_name'],
                      f"{field_stats['accuracy_rate'].min():.2f}", field_stats.iloc[-1]['field_name'],
                      f"{inaccurate_edited:,}", f"{inaccurate_not_edited:,}"]
        })
        overall_stats_df.to_excel(writer, sheet_name='Overall_Statistics', index=False)

        # Sheet 3: Inaccurate Records (for reference)
        inaccurate_records = df[~df['is_accurate']][['field_name', 'original_value', 'final_value', 'is_edited', 'record_type']].copy()
        inaccurate_records['inaccuracy_reason'] = inaccurate_records.apply(
            lambda row: 'is_edited=1 (OCR failed)' if row['is_edited'] == 1 else 'Value mismatch (is_edited=0)',
            axis=1
        )
        inaccurate_records.to_excel(writer, sheet_name='Inaccurate_Records', index=False)

        # Sheet 4: Edited Records Breakdown
        edited_records = df[df['is_edited'] == 1][['field_name', 'original_value', 'final_value', 'record_type']]
        edited_records.to_excel(writer, sheet_name='Edited_Records', index=False)

    print(f"🎉 Analysis completed successfully!")
    print(f"📁 Output file: {output_path}")
    print(f"   - Sheet 1: Field_Accuracy (accuracy per field with edited breakdown)")
    print(f"   - Sheet 2: Overall_Statistics (summary statistics)")
    print(f"   - Sheet 3: Inaccurate_Records (detailed inaccurate records with reasons)")
    print(f"   - Sheet 4: Edited_Records (all records where is_edited=1)")

    return field_stats


def main():
    """Main function to run the accuracy analysis"""

    # Configuration - Use latest cleansed file
    INPUT_FILE = "Sample_Comparison_data_20261103_cleansed_20260311_165145.xlsx"

    print("=" * 60)
    print("📊 Field Accuracy Analysis (Considering is_edited)")
    print("=" * 60)
    print()

    # Perform accuracy analysis
    field_stats = analyze_accuracy(
        input_file=INPUT_FILE,
        output_file=None  # Auto-generate output filename
    )

    print()
    print("=" * 60)
    print("Summary:")
    print(f"  Input file: {INPUT_FILE}")
    print(f"  Output file: {Path(INPUT_FILE).stem}_accuracy_summary_*.xlsx")
    print("=" * 60)


if __name__ == "__main__":
    main()

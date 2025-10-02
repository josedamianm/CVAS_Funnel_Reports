#!/usr/bin/env python3
"""
Category Report Generator

This script transforms category data from row-based format to column-based format,
with categories as columns and metrics as rows.

Usage:
    python generate_category_report.py <input_file> [output_file]

Example:
    python generate_category_report.py data.xlsx
    python generate_category_report.py data.xlsx output.xlsx
"""

import sys
import pandas as pd
from pathlib import Path


def read_category_data(input_file):
    """
    Read category data from Excel file's 'Export' sheet.
    
    Args:
        input_file (str): Path to input Excel file
        
    Returns:
        pd.DataFrame: Raw data from Export sheet
    """
    print(f"Reading data from: {input_file}")
    try:
        df = pd.read_excel(input_file, sheet_name='Export')
        print(f"✓ Successfully read {len(df)} rows from 'Export' sheet")
        return df
    except Exception as e:
        print(f"✗ Error reading file: {e}")
        sys.exit(1)


def transform_to_column_format(df):
    """
    Transform data from row-based (one row per category) to column-based format
    (one column per category with metrics as rows).
    
    Args:
        df (pd.DataFrame): Input dataframe with categories as rows
        
    Returns:
        pd.DataFrame: Transformed dataframe with categories as columns
    """
    print("\nTransforming data structure...")
    
    # Define the exact metric order
    metrics = [
        '[TopLine_Revenue]',
        '[Base_usuarios]',
        '[v_Activaciones_Revenue]',
        '[v__Activaciones]',
        '[v_Renovaciones_Revenue]',
        '[v_Renovaciones]',
        '[v_Rfnds]',
        '[Rfnds_U_U]',
        '[Total_Refnds]',
        '[v__Churn_from_act2]',
        '[v__Chur_prev_base]',
        '[v__Churn]',
        '[v_Auto_Ref]',
        '[Auto_Ref_UU]',
        '[Automatic_Refund_Amount]',
        '[v_Reg_Ref]',
        '[Reg_Ref_UU]',
        '[Reg_Refund_Amount]'
    ]
    
    # Define the category order
    categories = [
        'Beauty and Health',
        'Free Time',
        'Games',
        'Education',
        'Images',
        'Kids',
        'Light',
        'Music',
        'News',
        'Sports'
    ]
    
    # Create output dataframe with metrics as index
    output_df = pd.DataFrame(index=metrics)
    output_df.index.name = 'Master_CPC[TME Category]'
    
    # Use the 'Master_CPC[TME Category]' column for category names
    category_col = 'Master_CPC[TME Category]'
    
    # Process each category
    for category in categories:
        # Find the row for this category
        category_row = df[df[category_col] == category]
        
        if len(category_row) == 0:
            print(f"  ⚠ Warning: Category '{category}' not found in data")
            # Fill with zeros if category not found
            output_df[category] = 0
        else:
            # Extract values for each metric
            category_data = category_row.iloc[0]
            
            # Map metrics to their values
            for metric in metrics:
                if metric in category_data.index:
                    output_df.loc[metric, category] = category_data[metric]
                else:
                    # If metric not found, set to 0
                    output_df.loc[metric, category] = 0
            
            print(f"  ✓ Processed category: {category}")
    
    # Replace NaN with 0 first, before calculating Edu+Img
    output_df = output_df.fillna(0)
    
    # Add the special "Edu+Img" column after "Images"
    print("  Adding 'Edu+Img' calculated column...")
    if 'Education' in output_df.columns and 'Images' in output_df.columns:
        edu_img_values = output_df['Education'] + output_df['Images']
        
        # Insert after Images column
        images_idx = output_df.columns.get_loc('Images')
        output_df.insert(images_idx + 1, 'Edu+Img', edu_img_values)
    else:
        print("  ⚠ Warning: Could not create 'Edu+Img' column (Education or Images missing)")
    
    print(f"✓ Transformation complete: {len(output_df)} metrics × {len(output_df.columns)} categories")
    
    return output_df


def save_output(df, output_file):
    """
    Save transformed dataframe to Excel file.
    
    Args:
        df (pd.DataFrame): Transformed dataframe
        output_file (str): Path to output Excel file
    """
    print(f"\nSaving output to: {output_file}")
    try:
        df.to_excel(output_file)
        print(f"✓ Successfully saved output file")
    except Exception as e:
        print(f"✗ Error saving file: {e}")
        sys.exit(1)


def main():
    """Main execution function."""
    # Parse command line arguments
    if len(sys.argv) < 2:
        print("Usage: python generate_category_report.py <input_file> [output_file]")
        print("\nExample:")
        print("  python generate_category_report.py data.xlsx")
        print("  python generate_category_report.py data.xlsx output.xlsx")
        sys.exit(1)
    
    input_file = sys.argv[1]
    
    # Generate output filename if not provided
    if len(sys.argv) >= 3:
        output_file = sys.argv[2]
    else:
        input_path = Path(input_file)
        output_file = str(input_path.parent / f"{input_path.stem}_output.xlsx")
    
    print("=" * 70)
    print("CATEGORY REPORT GENERATOR")
    print("=" * 70)
    
    # Read input data
    df = read_category_data(input_file)
    
    # Transform data
    output_df = transform_to_column_format(df)
    
    # Save output
    save_output(output_df, output_file)
    
    # Display preview
    print("\n" + "=" * 70)
    print("PREVIEW - First 5 rows:")
    print("=" * 70)
    print(output_df.head())
    
    print("\n" + "=" * 70)
    print(f"✓ Process completed successfully!")
    print(f"  Input:  {input_file}")
    print(f"  Output: {output_file}")
    print("=" * 70)


if __name__ == "__main__":
    main()

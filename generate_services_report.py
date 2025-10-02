#!/usr/bin/env python3
"""
OM Services Data Transformation Script

This script transforms OM Services data from Power BI export format to a 
standardized reporting format with hardcoded structure.

Transformation:
- Input: Services as rows, metrics as columns
- Output: Metrics as rows, services as columns (in predefined order)

Usage:
    python generate_services_report.py <input_file> [output_file]
    
Example:
    python generate_services_report.py 202509_data_funnel_services.xlsx services_report.xlsx
    
Author: DeepAgent
Date: 2025-10-02
"""

import sys
import pandas as pd
import openpyxl
from pathlib import Path
from typing import List


def read_input_data(input_file: str) -> pd.DataFrame:
    """
    Read the input Excel file containing services data.
    
    Args:
        input_file: Path to the input Excel file
        
    Returns:
        DataFrame containing the input data
        
    Raises:
        FileNotFoundError: If input file doesn't exist
        ValueError: If required columns are missing
    """
    print(f"📖 Reading input file: {input_file}")
    
    if not Path(input_file).exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")
    
    # Read the Export sheet
    df = pd.read_excel(input_file, sheet_name='Export')
    
    # Verify required columns exist
    if 'Master_CPC[Service Name]' not in df.columns:
        raise ValueError("Missing required column: Master_CPC[Service Name]")
    
    print(f"   ✓ Found {len(df)} services")
    print(f"   ✓ Found {len(df.columns)} columns")
    
    return df


def get_metric_columns(df: pd.DataFrame) -> List[str]:
    """
    Extract the metric column names from the dataframe.
    
    Metric columns are those that start with '[' and end with ']'.
    
    Args:
        df: Input dataframe
        
    Returns:
        List of metric column names
    """
    metric_cols = [col for col in df.columns if col.startswith('[') and col.endswith(']')]
    print(f"   ✓ Identified {len(metric_cols)} metric columns")
    return metric_cols


def transform_data(df: pd.DataFrame, metric_cols: List[str]) -> pd.DataFrame:
    """
    Transform the data from services-as-rows to metrics-as-rows format.
    
    Args:
        df: Input dataframe with services as rows
        metric_cols: List of metric column names
        
    Returns:
        Transformed dataframe with metrics as rows and services as columns
    """
    print("🔄 Transforming data structure...")
    
    # Extract service names and metrics
    services = df['Master_CPC[Service Name]'].tolist()
    
    # Create a dictionary to hold the transformed data
    transformed_data = {'Master_CPC[Service Name]': metric_cols}
    
    # For each service, extract all metric values
    for service in services:
        service_row = df[df['Master_CPC[Service Name]'] == service].iloc[0]
        service_values = [service_row[metric] for metric in metric_cols]
        transformed_data[service] = service_values
    
    # Create the transformed dataframe
    df_transformed = pd.DataFrame(transformed_data)
    
    print(f"   ✓ Transformed to {len(df_transformed)} rows × {len(df_transformed.columns)} columns")
    
    return df_transformed


def get_hardcoded_structure() -> tuple[List[str], List[str]]:
    """
    Return the hardcoded output structure (metric and service order).
    
    Returns:
        Tuple of (metric_order, service_order)
    """
    print("📋 Using hardcoded output structure")
    
    # Hardcoded metric order (rows)
    metric_order = [
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
    
    # Hardcoded service column order
    service_order = [
        'IntimaX',
        'Rincon Prohibido',
        'The Tourist',
        'El Mundo Al Revés',
        'Noticias Emocion',
        'Deportes emocion',
        'Cuidate Mejor',
        'Sexducate con LB',
        'Yo Mujer y +',
        'Slow Life',
        'Movistar Juegos',
        'Kids Play',
        'Smile & Learn'
    ]
    
    print(f"   ✓ Structure has {len(metric_order)} metrics")
    print(f"   ✓ Structure has {len(service_order)} services")
    
    return metric_order, service_order


def apply_output_structure(df: pd.DataFrame, metric_order: List[str], 
                          service_order: List[str]) -> pd.DataFrame:
    """
    Apply the hardcoded output structure to the transformed data.
    
    This ensures:
    1. Metrics are in the correct order
    2. Services (columns) are in the correct order
    3. Only specified metrics are included
    
    Args:
        df: Transformed dataframe
        metric_order: Desired order of metrics (rows)
        service_order: Desired order of services (columns)
        
    Returns:
        DataFrame matching the output structure
    """
    print("📐 Applying output structure...")
    
    # Filter to only include metrics in the specified order
    df_filtered = df[df['Master_CPC[Service Name]'].isin(metric_order)].copy()
    
    # Create a mapping for metric order
    metric_order_map = {metric: i for i, metric in enumerate(metric_order)}
    df_filtered['_sort_order'] = df_filtered['Master_CPC[Service Name]'].map(metric_order_map)
    
    # Sort by the metric order
    df_filtered = df_filtered.sort_values('_sort_order').drop('_sort_order', axis=1)
    
    # Reorder columns to match structure (first column + services in specified order)
    column_order = ['Master_CPC[Service Name]'] + service_order
    df_final = df_filtered[column_order]
    
    print(f"   ✓ Final structure: {len(df_final)} rows × {len(df_final.columns)} columns")
    
    return df_final


def save_output(df: pd.DataFrame, output_file: str):
    """
    Save the transformed data to an Excel file.
    
    Args:
        df: Transformed dataframe to save
        output_file: Path to the output Excel file
    """
    print(f"💾 Saving output to: {output_file}")
    
    # Create the output directory if it doesn't exist
    Path(output_file).parent.mkdir(parents=True, exist_ok=True)
    
    # Save to Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Services Report', index=False)
    
    print(f"   ✓ File saved successfully")


def preview_results(df: pd.DataFrame):
    """
    Display a preview of the transformation results.
    
    Args:
        df: Transformed dataframe to preview
    """
    print("\n" + "=" * 80)
    print("📊 PREVIEW OF RESULTS")
    print("=" * 80)
    
    print(f"\nOutput shape: {df.shape[0]} rows × {df.shape[1]} columns")
    
    print("\nMetrics (rows):")
    for i, metric in enumerate(df['Master_CPC[Service Name]'].tolist()[:10], 1):
        print(f"  {i}. {metric}")
    if len(df) > 10:
        print(f"  ... and {len(df) - 10} more metrics")
    
    print("\nServices (columns):")
    services = df.columns[1:].tolist()
    for i, service in enumerate(services, 1):
        print(f"  {i}. {service}")
    
    print("\nFirst few rows of data:")
    print(df.head(10).to_string())
    
    print("\nSummary statistics:")
    # Count non-null values for each service
    for service in services[:5]:  # Show first 5 services
        non_null = df[service].notna().sum()
        print(f"  {service}: {non_null}/{len(df)} metrics have data")
    if len(services) > 5:
        print(f"  ... and {len(services) - 5} more services")


def main():
    """Main execution function."""
    print("=" * 80)
    print("OM SERVICES DATA TRANSFORMATION")
    print("=" * 80)
    print()
    
    # Parse command line arguments
    if len(sys.argv) < 2:
        print("❌ Error: Missing required argument")
        print()
        print("Usage: python generate_services_report.py <input_file> [output_file]")
        print()
        print("Arguments:")
        print("  input_file   : Path to the input Excel file (e.g., 202509_data_funnel_services.xlsx)")
        print("  output_file  : (Optional) Path to the output Excel file (default: services_report.xlsx)")
        print()
        print("Example:")
        print("  python generate_services_report.py 202509_data_funnel_services.xlsx")
        print("  python generate_services_report.py input.xlsx output.xlsx")
        print()
        print("Output Structure:")
        print("  - Metrics (rows): 18 predefined metrics in fixed order")
        print("  - Services (columns): 13 services in fixed order")
        sys.exit(1)
    
    input_file = sys.argv[1]
    input_path = Path(input_file)
    output_file = sys.argv[2] if len(sys.argv) > 2 else str(input_path.parent / f"{input_path.stem}_output.xlsx")

    try:
        # Step 1: Read input data
        df_input = read_input_data(input_file)
        
        # Step 2: Identify metric columns
        metric_cols = get_metric_columns(df_input)
        
        # Step 3: Transform data structure
        df_transformed = transform_data(df_input, metric_cols)
        
        # Step 4: Get hardcoded structure
        metric_order, service_order = get_hardcoded_structure()
        
        # Step 5: Apply output structure
        df_final = apply_output_structure(df_transformed, metric_order, service_order)
        
        # Step 6: Save output
        save_output(df_final, output_file)
        
        # Step 7: Preview results
        preview_results(df_final)
        
        print("\n" + "=" * 80)
        print("✅ TRANSFORMATION COMPLETED SUCCESSFULLY")
        print("=" * 80)
        print(f"\n📁 Output saved to: {output_file}")
        print()
        
    except Exception as e:
        print("\n" + "=" * 80)
        print("❌ ERROR OCCURRED")
        print("=" * 80)
        print(f"\n{type(e).__name__}: {e}")
        print()
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()

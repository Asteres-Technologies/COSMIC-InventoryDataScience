"""
This is a module to clean the COSMIC Technology Inventory Database Snapshots.
The snapshots are Excel files that have two sheets, "Cover_Page" and "Inventory".
The inventory sheet contains the actual data with the following columns:
- Technology Name: The name of the technology.
- Tech Producer: The producer of the technology.
- Description: A brief description of the technology.
- Existing Technology: Whether the technology is existing or not.
- Level One Category: The first level category of the technology.
- Level Two Category: The second level category of the technology.
- Level Three Category: The third level category of the technology.
- TRL: The Technology Readiness Level of the technology.
- Level One Functional Category: The first level functional category of the technology.
- Level Two Functional Category: The second level functional category of the technology.

This module provides functions to clean the data, including:
- Removing duplicate rows.
- Standardizing the Technology and Functional Category inputs.
- Handling missing values.
- Flagging missing or incomplete data with "N/A" or "Incomplete" as appropriate.
- Saving the cleaned data to a new Excel file with the same structure and formatting as the original file.
"""
from pandas import read_excel, DataFrame, ExcelWriter
import re
import os

from sqlalchemy import desc


CLEAN_DESCRIPTION = re.compile(r'[^\w\s,.\-]') # Strip out any non-alphanumeric characters except for spaces, commas, periods, and hyphens


def clean_taxonomy_labels(label: str) -> str:
    """
    If the taxonomy label doesn't have a colon, it is not a valid label.
    If it is not valid, we should try to repair it.
    We can repair it by finding the last number in the label and adding a colon after it.
    If this doesn't work, just return the label as is.
    If we clean the label, return the label split by colon.
    Args:
    - label (str): The taxonomy label to clean.
    Returns:
    - str: The cleaned taxonomy label, or None if the label is invalid.
    """
    if not label:
        return None
    
    converted_label = label.strip().title()
    if ':' not in converted_label:
        last_number = re.search(r'\d+', converted_label)
        if last_number:
            converted_label = converted_label[:last_number.end()] + ':' + converted_label[last_number.end():]

    if ':' in converted_label:
        return converted_label.split(':', 1)[0].strip()

    return converted_label


def standardize_inventory_data(file_path: str) -> DataFrame:
    """
    Cleans the COSMIC Technology Inventory Database Snapshot.
    
    Args:
    - file_path (str): The path to the Excel file containing the inventory data.
    
    Returns:
    - DataFrame: A cleaned DataFrame with standardized and processed data.
    """
    df = read_excel(file_path, sheet_name='Inventory')
    df.drop_duplicates(inplace=True)
    
    # Standardize Name and Producer fields
    df['Technology Name'] = df['Technology Name'].str.strip().str.title()
    df['Tech Producer'] = df['Tech Producer'].str.strip().str.title()

    # Make all category fields the longest value of either the level 1, 2, or 3 category
    # These columns are the worst offenders for inconsistent data Entries range from "TX04.5.5: Capture Mechanisms and Fixtures" to "1"
    # So we are just using the longest string from the 3 category columns, but first try to split by colon and check the longest part
    def longest_tech_category(row):
        """ 
        Args: 
        - row: pandas Series representing a row in the DataFrame
        Returns:
        - str: The longest category string from the three category columns.
        """
        categories = [row['Level One Category'], row['Level Two Category'], row['Level Three Category']]
        categories = [clean_taxonomy_labels(cat) for cat in categories if isinstance(cat, str)]
        if not categories:
            return "Unknown Technology Category"
        longest = max(categories, key=len)
        return longest
    
    taxonomy_values = df.apply(longest_tech_category, axis=1)
    df['Level One Category'] = taxonomy_values
    df['Level Two Category'] = taxonomy_values
    df['Level Three Category'] = taxonomy_values

    def longest_functional_category(row):
        """
        Args: 
        - row: pandas Series representing a row in the DataFrame
        Returns:
        - str: The longest functional category string from the two functional category columns.
        """
        categories = [row['Level One Functional Category'], row['Level Two Functional Category']]
        categories = [clean_taxonomy_labels(cat) for cat in categories if isinstance(cat, str)]
        if not categories:
            return "Unknown Functional Category"
        longest = max(categories, key=len)
        return longest
    
    df['Level One Functional Category'] = df.apply(longest_functional_category, axis=1)
    df['Level Two Functional Category'] = df.apply(longest_functional_category, axis=1)

    # Handle missing values for tech name and producer
    df['Technology Name'].fillna('Unknown Technology Name', inplace=True)
    df['Tech Producer'].fillna('Unknown Tech Producer', inplace=True)

    # Handle missing values for description
    df['Description'].fillna('No Description Available', inplace=True)
    df['Existing Technology'].fillna('Unknown Existing Technology', inplace=True)
    df['TRL'].fillna(0, inplace=True)  # Using 0 to indicate missing TRL
    df['Level One Category'].fillna('Unknown Technology Category', inplace=True)
    df['Level Two Category'].fillna('Unknown Technology Category', inplace=True)
    df['Level Three Category'].fillna('Unknown Technology Category', inplace=True)
    df['Level One Functional Category'].fillna('Unknown Functional Category', inplace=True)
    df['Level Two Functional Category'].fillna('Unknown Functional Category', inplace=True)

    # Validate description
    def clean_description(desc):
        """
        Args:
        - desc (str): The description as a string
        Returns:
        - str: The string with invalid characters removed.
        """ 
        if not isinstance(desc, str):
            return 'No Description Available'
        cleaned_text = CLEAN_DESCRIPTION.sub('', desc)
        return cleaned_text.strip() if cleaned_text.strip() else 'No Description Available'
    
    df['Description'] = df['Description'].apply(clean_description)

    return df
import pandas as pd
import os

def create_template_file(file_path):

    columns_messages = ["Name", "Indicative", "Number", "Message"]
    df = pd.DataFrame(columns=columns_messages)
    df.to_excel(file_path, index=False)
    return True

def load_excel_data(file_path):

    expected_columns = ["Name", "Indicative", "Number", "Message"]
    df = pd.read_excel(file_path)
    if not all(col in df.columns for col in expected_columns):
        raise ValueError(f"The file does not have the expected columns. Expected columns: {expected_columns}")
    return df

def format_time(seconds):

    hours = int(seconds // 3600)
    minutes = int((seconds % 3600) // 60)
    seconds = int(seconds % 60)
    return hours, minutes, seconds
import pandas as pd
import os

def create_sample_excel(file_path):
    data = {
        'Name': ['Alice', 'Bob', 'Charlie', 'Alice', None],
        'Age': [25, 30, 35, 25, None],
        'City': ['New York', 'Los Angeles', 'Chicago', 'New York', None],
        'Status': ['Complete', 'Pending', 'Complete', 'Complete', None]
    }
    df = pd.DataFrame(data)
    
    # Add an empty row and a duplicate row (Alice)
    # The 'None' row is already added
    
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    df.to_excel(file_path, index=False)
    print(f"Created sample Excel at {file_path}")

if __name__ == "__main__":
    create_sample_excel("input/sample_data.xlsx")

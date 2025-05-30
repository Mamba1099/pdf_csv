import pandas as pd

# Reconstruct the data from the image in structured form
data = [
    # Input your structured data here
]

# Create DataFrame
df = pd.DataFrame(data)

# Save to Excel
excel_file_path = r"path to the file name"
df.to_excel(excel_file_path, index=False)

excel_file_path

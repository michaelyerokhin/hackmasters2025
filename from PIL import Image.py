from PIL import Image
import pytesseract
import csv
import pandas as pd
import os

# Directory containing images
image_dir = '/Users/connorv-e/Desktop/hackathon'  # Replace with the folder containing your images

# Get a list of image file paths
image_files = [os.path.join(image_dir, f) for f in os.listdir(image_dir) if f.endswith(('.png', '.jpg', '.jpeg'))]

# Initialize a list to store the extracted data
data_list = []

# Loop through each image file
for image_path in image_files:
    print(f"Processing: {image_path}")
    try:
        # Load the image
        image = Image.open(image_path)

        # Perform OCR to extract text
        ocr_text = pytesseract.image_to_string(image)

        # Parse the OCR text into key-value pairs
        data = {}
        lines = ocr_text.split('\n')
        for line in lines:
            if ':' in line:  # Look for key-value pairs
                key, value = line.split(':', 1)
                last_word = key.split()[-1]
                data[last_word.strip()] = value.strip()

                if last_word.strip() == "Number":
                    data_list.append(data)  # Add the data from this file
                    print(f"Key 'Number' found in {image_path}. Skipping to next file.")
                    break  # Stop processing this file and move to the next

        # Add the parsed data to the list
        if data:  # Only add if data was successfully parsed
            data_list.append(data)

    except Exception as e:
        print(f"Error processing {image_path}: {e}")

# Convert the data list into a DataFrame
df = pd.DataFrame(data_list)

# Export to Excel and CSV
df.to_excel('employee_data_multiple.xlsx', index=False)  # Save as Excel
df.to_csv('employee_data_multiple.csv', index=False)    # Save as CSV

print("Data from multiple images has been exported successfully!")

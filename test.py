import csv

# Define the path to the uploaded CSV file
csv_file_path = 'testFile.csv'

# Read and display the contents of the CSV file
with open(csv_file_path, mode='r') as file:
    reader = csv.reader(file)
    data = list(reader)

# Display the contents of the CSV file
for row in data:
    print(row)
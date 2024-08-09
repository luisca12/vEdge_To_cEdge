[11:13 AM] Luis Alfaro
[9:51 AM] Kristel Gutiérrez
   try:
        with open(sdw03Template, "r") as inputCSV:
            # open(newSDW03Template, 'w') as outputCSV
            # authLog.info(f"Generating {site_code}-SDW-03-Template")
            # print(f"INFO: Generating {site_code}-SDW-03-Template.")
            csvReader = csv.reader(inputCSV)
            # csvWriter = csv.writer(outputCSV) 
            rows = list(csvReader)  
 
            # for rows in csvReader:
            #     rowData = []
            #     for cell in rows:
            #         cellValue = str(cell).strip()
            #         for key, value in sdw03Replacements.items():
            #             if key.lower() in cellValue.lower():
            #                 cellValue = cellValue.replace(key, value)
            #         rowData.append(cellValue)
            #     csvWriter.writerow(rowData)
        if len(rows) > 1:
            second_row = rows[1]
            modified_row = []
            for cell in second_row:
                cellValue = str(cell).strip()
                for key, value in sdw03Replacements.items():
                    if key.lower() in cellValue.lower():
                        cellValue = cellValue.replace(key, value)
                modified_row.append(cellValue)
            rows[1] = modified_row  
        with open(newSDW03Template, 'w', newline='') as outputCSV:
            csvWriter = csv.writer(outputCSV)
            csvWriter.writerows(rows)
        with open(sdw04Template, "r") as inputCSV1:
            csvReader1 = csv.reader(inputCSV1)
            # csvWriter = csv.writer(outputCSV) 
            rows1 = list(csvReader1)  
            authLog.info(f"Generating {site_code}-SDW-04-Template")
            print(f"INFO: Generating {site_code}-SDW-04-Template.")
 
        
        if len(rows1) > 1:
            second_row = rows1[1]
            modified_row = []
            for cell in second_row:
                cellValue = str(cell).strip()
                for key, value in sdw04Replacements.items():
                    if key.lower() in cellValue.lower():
                        cellValue = cellValue.replace(key, value)
                modified_row.append(cellValue)
            rows1[1] = modified_row  
 
           # csvWriter1 = csv.writer(outputCSV1)   
 
            # for rows in csvReader1:
            #     rowData1 = []
            #     for cell in rows:
            #         cellValue1 = str(cell).strip()
            #         for key, value in sdw04Replacements.items():
            #             if key.lower() in cellValue1.lower():
            #                 cellValue1 = cellValue1.replace(key, value)
            #         rowData1.append(cellValue1)
            #     csvWriter1.writerow(rowData1)
            with open(newSDW04Template, 'w', newline='') as outputCSV:
                csvWriter1 = csv.writer(outputCSV)
                csvWriter1.writerows(rows)
 
    except Exception as error:
        print(f"ERROR: {error}\n", traceback.format_exc())
        authLog.error(f"Error message: {error}\n", traceback.format_exc())

 
 
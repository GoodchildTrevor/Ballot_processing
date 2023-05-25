# Ballot_processing

This Python script provides a solution for data cleaning, transformation, and aggregation using Pandas and Openpyxl libraries. 
The script processes input data, cleanses it using the replacer function, calculates points, aggregates them, and then stores the results in an Excel file.

# How it Works
In the first part of the script, we iterate over df_original dataframe items. 
For each non-null value in each column, we populate the df_second dataframe with specific data, including user names, values, and a point value. 
The script uses the nomination list to format the column names. The count_nomination is used to keep track of the current row index in df_second.

In the second part, we iterate over the data items. We create a new dataframe df_movies with users, movies, and points. 
The top list is used to assign points. 
The count is used to keep track of the current row index in df_movies.

# Workflow

The next section applies various operations to each nomination column in df_second dataframe and the movie column in df_movies dataframe, including:

* Changing the text to lowercase.
* Removing or replacing unwanted characters using the replacer function.
* Applying the prob function.
* Applying the change function with dataframe and column name as arguments.
* Then, the script creates a new dataframe df_third that counts how many times each value appears in the respective column of df_second. 
* The result is sorted in descending order, and then processed by the results function. 
* Also the script calculates the total points for each movie in df_movies and sorts the result in descending order. 
* In the final part, this processed data is written into an Excel workbook (ws) using Openpyxl.
Additionally, it writes down nomination names and "points" into the first row of the Excel workbook. 
The final workbook is saved to the file system with the filename fn.

# Usage
Before running this script, please ensure you have the following prerequisites:

* Install necessary Python packages, Pandas and Openpyxl.
* To get the expected results, your Excel file structure should resemble the example provided. 
Please ensure the structure and layout of your Excel workbook matches the example before running the script. 
If it doesn't, you may need to modify the script or adjust your Excel file to meet the script's requirements.
* You need to define wb as an Openpyxl workbook object, and ws as a worksheet within this workbook. Define fn as the desired filename for the final Excel workbook.

# Limitations
This script is designed to work with specific data structures and functions. 
If the input data or functions are not defined as expected, the script might fail. 

# Consclusion
This Python script utilizes Pandas and Openpyxl libraries for data cleaning, transformation, and aggregation. It processes input data, assigns points based on selections, and exports results to an Excel file. The script provides clear and organized data for analysis, while relying on specific structures and functions. Further improvements can enhance flexibility and error handling. Overall, it showcases Python's capabilities for efficient ballot processing and data analysis.

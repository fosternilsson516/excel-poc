# Excel Find and Replace POC

## Overview
This repository hosts a proof of concept for an Excel find and replace tool built using Python and xlwings. This tool utilizes a template Excel file with two columns: one for "find" values and the other for "replace" values. It is designed to facilitate the cleaning of data, particularly special characters in Excel files before data uploads, and can also handle CSV files.

## Features
- **Dynamic Data Cleaning**: Connects to your working Excel or CSV file, cleans the data based on the template, and leaves the original file unchanged for auditing purposes. A new sheet with the applied changes is added.
- **Real-Time Updates**: Leverages xlwings to update data in real-time, functioning even when the Excel or CSV files are actively open by the user.
- **Additional Functionality**: Similar to the find and replace feature, another functionality involves using a template file with a single column of values to perform operations like keeping or removing matching values in the user's working file.

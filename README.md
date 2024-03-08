# Coordinates_Converter

Coordinates_Converter is a Python script designed to facilitate the conversion of celestial coordinates between different systems. This tool enables astronomers and researchers to effortlessly transform equatorial coordinates (Right Ascension and Declination) to galactic coordinates (l and b). Built upon popular Python libraries such as `openpyxl` and `astropy`, Coordinates_Converter ensures accuracy and efficiency in coordinate transformation tasks.

## Features

- **Coordinate Conversion:** Converts equatorial coordinates (RA and Dec) to galactic coordinates (l and b) using the robust `astropy.coordinates` module.
- **Flexible Input:** Accepts input data from various sources, including text files, Excel spreadsheets, databases, and custom data structures.
- **Interactive Interface:** Provides a user-friendly interface with clear instructions, making it easy to navigate and operate for users of all skill levels.
- **Customization Options:** Offers customizable filtering options for galactic coordinates based on user-defined criteria, facilitating tailored data analysis.
- **Versatile Application:** Suitable for a wide range of astronomical research projects, including cataloging celestial objects, studying galactic structures, and more.

## How to Use

1. Ensure Python 3.x is installed on your system.
2. Install the required libraries: `openpyxl`, `astropy`.
3. Prepare your input data in an Excel spreadsheet with the appropriate format:
   - Ensure the Excel file contains a sheet with the target name, RA, and Dec columns.
   - Save the Excel file.
4. Place the Python script (`Coordinates_Converter.py`) in the desired directory.
5. Run the script using a Python interpreter.
6. Follow the prompts to provide the path to the Excel file containing the input data.
7. Follow the prompts to specify the sheet name and the columns containing the target name, RA, and Dec.
8. Optionally, use filtering options to refine the output based on specific criteria.


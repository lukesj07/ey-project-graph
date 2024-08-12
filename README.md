## Overview
This project is developed for the InfoSec division of EY to provide a visual representation of the overall status and progress of various projects. The tool generates a radar chart image that displays project statuses categorized by different service areas.

## Features
- **Project Status Sorting**: Automatically categorizes projects into "GREEN", "AMBER", "ON HOLD", and "RED" based on their overall health, and sorts them based off of what service category they belong to.
- **Dynamic Position Calculation**: Ensures that project markers do not overlap on the radar chart, providing a clear and concise visualization.

## Installation

1. Install Python from the official [Python website](https://www.python.org/downloads/). Make sure to select the option to add Python to your PATH during installation.
2. Open Command Prompt and install the required libraries using:

   ```bash
   pip install pandas matplotlib openpyxl numpy
   ```
   NOTE: command and process varies slightly on different operating systems

## Usage
1. **Prepare the Data**: Ensure your base image (`img.jpg`) stored in an Excel spreadsheet (`spreadsheet.xlsx`) with the following columns:
   - `Project Name`
   - `Overall Health`
   - `%Project Duration Completed`
   - `Service Category`
   - `Radar ID`

     Make sure the image file (`img.jpg`), the spreadsheet (`spreadsheet.xlsx`), and the main Python script (`main.py`) are all located in the same folder.

2. **Customize the Radar Chart**: Adjust the parameters such as `IMG_WIDTH`, `IMG_HEIGHT`, and sector bounds within the script to fit your specific requirements.

3. **Run the Script**
   ```bash
   python main.py
   ```
   NOTE: command varies, most of the time the file can be double clicked to run

## File Structure
- `spreadsheet.xlsx`: The Excel file containing project data.
- `img.jpg`: The background image used for the radar chart.
- `main.py`: The main script that generates the radar chart.
- `radar.png`: The output radar chart image.

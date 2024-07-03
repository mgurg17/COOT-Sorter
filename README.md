# COOT Sorter

## Description
COOT Sorter is a specialized tool designed to automate the process of sorting students into College Outdoor Orientation Trips (COOT). Developed by a COOT Coordinator, this application streamlines the complex task of assigning students to trips based on their preferences while considering various constraints such as gender balance, roommate/teammate placement, and trip capacities.

## Features
- User-friendly GUI for easy file selection and processing
- Efficient sorting algorithm using Pandas DataFrames
- Considers student preferences, gender balance, roommate/teammate placements, and trip capacities
- Outputs sorted assignments and statistics in Excel format

## Requirements
- Python 3.11 or later
- Required libraries: pandas, numpy, tkinter, ttkthemes

## Installation
1. Clone this repository or download the source code.
2. Install the required libraries:
pip install pandas numpy tkinter ttkthemes

## Usage
1. Run `SorterGUI.py` to launch the application.
2. Click "Select File" to choose your input Excel file.
3. Click "Process File" to run the sorting algorithm.
4. Save the output file when prompted.

### Input File Format
The input Excel file should contain two sheets:
1. **Students**: Columns should include Student ID, Gender, Team, POC, Dorm, Water, Tent, and preferences (A-J) corresponding to trip categories.
2. **Trips**: Columns should include Trip, Category, Capacity, Water, and Tent.

## How It Works
The sorting algorithm:
1. Randomizes the order of students to ensure fairness.
2. For each student, it goes through their list of preferences.
3. For each preference, it checks every trip within that category for eligibility.
4. Eligibility is determined by:
- Roommate/teammate placement (avoiding conflicts)
- Gender balance
- Trip capacity

## Executables
Executable versions for both macOS and Windows are available for users who prefer not to run the Python scripts directly.

## Future Improvements
- Implement alternative methods for specifying preferences
- Enhance the sorting algorithm for improved efficiency and results

## Contributing
Contributions, issues, and feature requests are welcome. Feel free to check [issues page](https://github.com/yourusername/coot-sorter/issues) if you want to contribute.

## Author
Matt Gurgiolo

## License
This project is open source and available under the [MIT License](LICENSE.txt).

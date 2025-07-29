# Timesheet Analysis Tool

A Python application that automates the retrieval and analysis of work timesheets from SDMataClick.com, providing visual reports of work hours against targets.


## Features

- Automated login and timesheet retrieval from SDMataClick.com
- Detailed analysis of daily and weekly work hours
- Visual charts comparing actual hours vs target hours
- HTML report generation with interactive charts
- Persistent storage of ID for convenience


## Requirements

- Python 3.6+
- Chrome browser installed
- ChromeDriver (automatically handled by the script)


## Installation

1. Clone this repository or download the files
2. Install the required packages:
	- `pip install -r requirements.txt`


## Usage

1. Run the application by double-clicking `Timesheet.bat` or executing:
	- `python main.py`
2. On first run, you'll be prompted to enter your ID number
3. The script will:
  	- Log in to SDMataClick.com
	- Retrieve your timesheet data
	- Analyze your work hours
	- Generate an HTML report
	- Open the report in your default browser


## Configuration

The tool automatically saves your ID number in `timesheet_config.json` after the first run. To change your ID, either:
	1. Delete the `timesheet_config.json` file and run the tool again
	2. Edit the file directly to update the `id_number` value


## Report Contents

The generated HTML report includes:

### Daily Analysis
- Arrival and departure times
- Total hours worked each day
- Comparison against daily target (default: 9 hours)
- Visual charts of daily hours and differences

### Weekly Analysis
- Total hours worked each week
- Days meeting/exceeding target
- Average daily hours
- Weekly difference from target
- Visual charts of weekly performance


## Customization

To change the daily target hours (default: 9), modify this line in `main.py`:
```python

daily_target_hours = 9  # Change this value to your desired target

```


## Troubleshooting

If you encounter issues:
1. Ensure you have a stable internet connection
2. Verify your ID number is correct
3. Check that SDMataClick.com is accessible
4. Make sure you have the latest version of Chrome installed
5. If you get an error explaining that the code cannot find or click an element, just rerun it


## Dependencies
The tool uses the following Python packages (automatically installed via requirements.txt):
	1. beautifulsoup4
	2. pandas
	3. selenium


## Author
Lee Kaplan (and so much ChatGPT)




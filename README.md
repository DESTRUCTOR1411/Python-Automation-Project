# Excel Price Processing Automation

A Python automation tool that processes Excel workbooks to apply price corrections and generate visual charts automatically.

## Description

This project automates the tedious task of processing pricing data in Excel files. It reads price values from column C, applies a 50% correction factor, stores the corrected prices in column D, and generates a bar chart visualization of the corrected data.

Perfect for businesses that need to:
- Apply bulk discounts or price adjustments
- Convert between different pricing models
- Generate quick visual reports from pricing data
- Automate repetitive Excel data processing tasks

## Features

- **Automated Price Processing**: Reads prices from column C and applies 50% correction
- **Data Cleaning**: Handles currency symbols ($) and comma separators automatically
- **Error Handling**: Skips blank cells to prevent processing errors
- **Visual Charts**: Automatically generates bar charts from corrected data
- **In-Place Processing**: Modifies existing Excel files while preserving original data

## Requirements

```
openpyxl>=3.0.0
```

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/excel-price-automation.git
cd excel-price-automation
```

2. Install required dependencies:
```bash
pip install openpyxl
```

## Usage

### Basic Usage

```python
from excel_processor import process_workbook

# Process your Excel file
process_workbook('your_data.xlsx')
```

### Expected Excel Format

Your Excel file should have:
- **Column A**: Any data (optional)
- **Column B**: Any data (optional)
- **Column C**: Original prices (e.g., "$100.00", "$1,234.56")
- **Column D**: Will be populated with corrected prices
- **Row 1**: Headers (will be skipped during processing)

### Example

**Before Processing:**
| A | B | C | D |
|---|---|---|---|
| Item | Description | Original Price | Corrected Price |
| Product 1 | Description 1 | $100.00 | |
| Product 2 | Description 2 | $200.50 | |

**After Processing:**
| A | B | C | D |
|---|---|---|---|
| Item | Description | Original Price | Corrected Price |
| Product 1 | Description 1 | $100.00 | 50.0 |
| Product 2 | Description 2 | $200.50 | 100.25 |

Plus a bar chart will be added at cell E2 showing the corrected prices.

## Code Structure

```python
def process_workbook(filename):
    """
    Main function that processes an Excel workbook
    
    Args:
        filename (str): Path to the Excel file to process
        
    Operations:
        1. Loads the workbook and accesses 'Sheet1'
        2. Iterates through rows starting from row 2
        3. Cleans price data (removes $ and , symbols)
        4. Applies 50% correction factor
        5. Stores corrected values in column D
        6. Creates a bar chart from corrected data
        7. Saves the modified workbook
    """
```

## Customization

You can easily modify the script for different use cases:

- **Change correction factor**: Modify `* 0.5` to your desired multiplier
- **Different columns**: Update column references (currently uses column 3 for input, 4 for output)
- **Chart position**: Change `'e2'` to your preferred chart location
- **Sheet name**: Update `'Sheet1'` if your data is in a different sheet

## Error Handling

The script includes basic error handling:
- Skips blank cells to prevent processing errors
- Handles currency formatting ($, commas)
- Strips whitespace from values

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## Future Enhancements

- [ ] Support for multiple sheets
- [ ] Configurable correction factors
- [ ] Different chart types (line, pie, etc.)
- [ ] Batch processing of multiple files
- [ ] GUI interface
- [ ] Command-line arguments support
- [ ] Better error handling and logging

## Contact

Raj Rai - rairaj1411@gmail.com
Project Link: https://github.com/yourusername/excel-price-automation

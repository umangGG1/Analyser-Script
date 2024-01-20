# Excel File Analysis Script

This script analyzes an Excel file containing employee data and prints information about their work patterns based on specific criteria.

## Features

- Identifies employees who have worked for 7 consecutive days.
- Detects employees with less than 10 hours between shifts but greater than 1 hour.
- Highlights employees who have worked for more than 14 hours in a single shift.

## Getting Started

1. **Clone the Repository:**
   ```bash
   git clone https://github.com/your-username/your-repository.git
   cd your-repository
2. **Run the Script:**
    Make sure you have Node.js installed.
    Install dependencies: npm install xlsx
   
4. **Execute the script:**
    node analyzeFileScript.js

Important Notes
Excel File Format:
Ensure your Excel file follows the format specified in the script comments.

Date and Time Format:
The script assumes consistent date and time formats in the Excel file.

# Excel VBA Insert New row and Merge it also

This VBA macro for Excel inserts an empty row at the top of the active sheet and adjusts the merged cells in the first row based on the length of the data in the second row. If the first row is already merged, the macro ensures it is correctly unmerged and re-merged according to the new length of the second row.

## Features

- Inserts an empty row at the top of the active sheet if not already present.
- Checks if the first row is merged and adjusts merging as necessary.
- Dynamically merges cells in the first row based on the length of data in the second row.
- Ensures consistent formatting and layout of merged cells.

## Installation

1. Open Excel.
2. Press `Alt + F11` to open the VBA editor.
3. Go to `Insert > Module` to create a new module.
4. Copy and paste the VBA code from below into the module.
5. Press `F5` to run the macro or close the VBA editor and run it from `Developer > Macros` in the Excel ribbon.

## Usage

1. Ensure your active sheet contains data in the second row that you want to base the merging on.
2. Run the macro to insert an empty row at the top and adjust the merged cells accordingly.

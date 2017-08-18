# Excel VBA Words Combinations
A simple script to create combinations of words from Excel's cells. 
For example

| First Column | Second Column | Third Column |
| ------------ | ------------- | ------------ |
| red | small | gloves |
| green | big | boots |
| blue | | |

results in 

- red small gloves
- red small boots
- green small gloves
- green big boots

etc..

## How to
Insert the different classes of words in columns (always  starting at the same height) without empty cells and fill their rows without empty cells again. Select the first (top left) cell and start the script. It will add a new sheet with the combinations created. Words don't need to start from the A1 cell. Looking at the example at the top of this document the cell with "red" is the one that need to be selected in the end. The script travels to the right until the first empty space and down until the first empty cell for every column. No data is written on already created sheets so this script should be safe 99.99999% of times. It hasn't been tested for big set of data.

## To import the script
You can use either the given .ods/.xlsm file or the script.txt one if you want to manually copy and paste it into your Worksheet. You may need (90%) to edit some security settings to run macros. The macro is called combination_composer.

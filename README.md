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
Insert the different classes of words in columns (always  starting at the same height) without empty cells and fill their rows without empty cells again. Select the first (top left) cell and start the script. It will add a new sheet with the combinations created. Words don't need to start from the A1 cell. Looking at the example at the top of this document the cell with "red" is the one that need to be selected in the end. The script travels to the right until the first empty space and down until the first empty cell for every column. 

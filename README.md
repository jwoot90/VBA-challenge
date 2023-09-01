I created variables for :
Symbol Name to store the Ticker Value
Total Stock Value to store the Total Volume
Summary Table Row to store the values for the lookup table after looping through the stock data
Total Open Value and Total Close Value to calculate the changes in stock value over each year
Yearly Percentage for the Summary Table
Ticker Index 1 , 2 & 3 to store the values for greatest percent increase ,decrease and total volume

I looped through all worksheets using  "For Each ws In Worksheets"

Looped through stock data with for loop to get summary table data 

Used the Ticker Index variable to store the index for the position of the values in column Q
I then 
put the values of these variables in the index function in column I to find the appropriate ticker

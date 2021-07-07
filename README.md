For this project I was asked to create a VBA coding script to run through a large dataset of stocks and then return the sumation of each ticker, its yearly change, percent change, and then the total stock volume that was done by each ticker each year. 


You can locate the script itself, inside of the the Repository VBA-Challenge with the name VBAScript. I performed this task with Windows 10, it should be available for the Mac verison as well but I have had instances when the code is too big Mac is unable to read it through all the way so it may be necessary to break down the one large script into smaller sub scripts. 


For the script, you should intialize the worksheet to set the active worksheet as the current one you are in and then perform a loop for each worksheet that you have within the Workbook. 

You then need to locate the last row of the worksheet and work your way up from the end. Next, you create the headers for the items that you want to create, including the ticker, yearly and percentage change, and then the total volume amount. Next you need to set each value equal to zero so not to start off with an incorrect value. 

Next you need to loop through each row, starting with the second row as to avoid headers when calculating. With this addition, you need find the ticker with the same name and continually add them up until you loop to a ticker with a new name, then reset back to zero and run the process again. To find the the yearly change you subtract the total close price vs the total open price.


You then use the color index to find if the change is greater than zero for green, less than zero for red, or if anything else change to a different color yellow. 

For percent change,  you would take the yearly change and then divide it by the open price to figure out the difference in percentage. 

Finally to run through the the greatest percentage increase, greatest percentage decrease and largest volume, you loop through the set and if the cell as you run through isn't the largest, then you loop to the next number to find the largest increase, decrease, and then the maximum volume.

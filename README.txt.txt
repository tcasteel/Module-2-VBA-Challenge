For the beginning of the VBA set:

	* I wrote down all of the Dims needed for the for loops and the ranges that were used within the code.
		They are set as Dims so that they can have specific values attached to them.

	* After the Dims were set I had to set them as Variables, so that most of them could state values, with the stated values, I also had to set a For Each so that the data would go across all the worksheets.
		Then had to set the ranges for the lookup arrays so that we can find the max and min values later.
		Then had to set most of the Dims to start at 0 to give a balanced start, while setting the sumtable to 2 so that it would start on the second row and not override the title.
		
	* I then stated the lastrow variable, as there are over 100k rows total to go through, so with that it helps it stop where the data ends.
		After stating the lastrow I set the titles for each of the vaules that were going to be looked at with using ws.Range() and then the Cell that it was going to be in.
	
	* The first for loop is up to bat, with this for loop I was able to gather the information that I needed with each column.
		For the ticker column, I had it loop around if the one ahead was the same and if not it would put it down into the sumtable.
			- For the Volume I had it count with the tickers so that when it didnt read the same ticker anymore it would add all previous values together.
		For the Year Column, I worte it so that it would take the closing value and subtract the opening value.
		For the percent Column it was very similar but with subtraction it would dvide it by the amount of times the ticker was on that value.
		For every time that it goes through the loop if they dont equal it adds to the sumtable so that when it looks for the next ticket it doesnt overide the previous data.
		It also resets the volume and ticker amount as we wouldnt want it to trail over to the next ticker.
		Once the if statement ends it resets the Year and percent so that it doesnt add again to the next loop.
		Then it goes to the next row of data to continue the set of data.

	*Next two for loops are for the cell colors to have it green for the positve results and red for the negative ones for both the year and the percent column.
		If the value is above 0 it sets the cell to green otherwise itll set it to red.

	* towards the end it set to have the table for the greatest % max, % min, and value.
		For the first set it is set to be able to look through the column and find which value is the highest percent on the sumtable, while the number format automatically sets it to the % value.
		For the second set it does the same thing but looks for the min value.
		the third didnt need to have any formating as it is the full value and it goes throughout the entire value column on the sumtable and finds the largest one.

	* Then with the Next ws it moves onto the next worksheet to be able to do the same process on the next worksheet.


	*Then for testing purposes I made a clear sub, so that it would clear all of the data that was gathered but leaves the titles, so that I was able to make sure that everything was working the way that it was meant to.
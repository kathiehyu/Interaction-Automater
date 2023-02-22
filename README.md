# Resident Interaction Automater
only works for Woodruff -- both North and South\
interactionautomater.py is the program you will run from the terminal (GUI coming soon maybe?)\
roster.xlsx is the excel sheet you will use to put all the needed information (resident names, room numbers, your name, etc)\
# Setting up and getting to know roster.xlsx
(row, column)\
- enter your name in cell (1, 1)\
- enter your building in cell (1, 2) with format WDN or WDS\
- enter your resident names into column 1 starting from row 2. you can copy & paste from your own excel if you have one and it should format correctly. If you have relatively fewer residents, just leave the remaining rows blank. I'll take care of it 😊\
- enter resident rooms into column 2 starting from row 2. you can leave it as 101Aa, etc. no need to remove the Aa (unless you want to). make sure they are with the correct resident name\
- the number in cell (1, 25) plus one is the number of the row of the first resident that it will fill out the form for when the program is run. For instance, when it is 1, the first resident it will fill out the form for is the resident on row 1 + 1 = 2. You will usually start at 1, and the automater will skip over residents with dates that you have not inputted. This is here so that if the program stops running for some reason, you can rerun it starting from the last row it completed. Don't forget to set it back to 1 once you finish.\
- if you want to receive emails of the form response, put your email in cell (1, 26). if not, leave it blank. i recommend it so you can see which ones were submitted and what the inputs were, but if you're not a fan of spam emails, don't do it.\

All of these will remain the same every time you use this automater (except the number in cell (1, 25), which you will only change if the program stops running before it finishes)

# Filling out the excel
- In column 3 and 4, labelled Date 1 and Date 2, enter the date number of your interaction, you can use single digits. Don't enter the month or year. For instance, if you've had two interactions with one resident on January 12 and January 24, in the Date 1 column, put 12, and Date 2 column, put 24 (or the other way around if that's your cup of tea). This is assuming that you are filling out this form within the same month of the interaction. Otherwise, sorry :( (maybe in the future I will implement choosing the month). If you only have one interaction, put the date into one column and leave the other blank. If you have no interactions with them, leave both blank.\
- In columns 5 through 10 (inclusive), mark the cell with the method of your interaction. They are labelled at the top row. You can mark with any text.\
- In columns 11 through 17, mark the cell with the topic of your interaction.\
- If you selected Intentional Conversation theme, in columns 18 through 23, mark the cell with the topic of your intentional conversation. If you didn't select intentional conversation theme, leave them blank.\
- If you have any comments on the interation, type it into the 24th column labelled comments. The excel won't show your whole comment if it's too long, for aesthetic purposes, but it's there. Leave it blank for no comment\
- Repeat for each resident you want to fill out the form for.\

# Running the program
Sorry, GUI for non CS people not created yet :')\
you must have python3 and selenium on your computer, I think
- In your terminal or command line, go into the directory with the program interactionautomer.py\
- Make sure that roster.xlsx is in the same directory\
- run the command python3 interactionautomater.py\
- the program will create a new window. You don't have to stay on the window, but you can't switch to another desktop or the program will stop running\
- the terminal/command line will have records of the inputs, you can use it to make sure it's on the right track, and if the program stops running, you can look at the last completed resident, and set the cell (1, 25) to that row. (keep in mind that that number **plus one** is the row that it will start at)\
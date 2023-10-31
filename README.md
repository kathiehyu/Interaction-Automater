# Resident Interaction Automater
Selenium script to automate filling out Resident Interactions Qualtrics form for Resident Assistants at Georgia Tech<br>
**Please** read the entire README up to running the program. It contains instructions on how to set up and use the program.
- only works for Woodruff -- both North and South
- interactionautomater.py is the program you will run from the terminal (GUI coming soon maybe?)
- roster.xlsx is the excel sheet you will use to put all the needed information (resident names, room numbers, your name, etc)

## Before you run the program . . .
Before using this program at all, there are a couple of technologies you should install
- Git 
    - [Installation instructions for Mac & Windows](https://www.atlassian.com/git/tutorials/install-git)

- Python
    - [Installation instructions for Mac Windows](https://kinsta.com/knowledgebase/install-python/#mac)
- Pip
    - Installation instructions for:
        - [Mac](https://www.geeksforgeeks.org/how-to-install-pip-in-macos/)
        - [Windows](https://www.geeksforgeeks.org/how-to-install-pip-on-windows/)


### Common Errors on Mac

1. After installing python, you may get the error <span style="color: #E03E3E;">zsh: command not found: python</span>. To fix this, run this command:
    ```
    echo "alias python=/usr/bin/python3" >> ~/.zshrc
    ```
    Restart your terminal, and it should work

2. After installing pip, you may get the error <span style="color: #E03E3E;">zsh: command not found: pip</span> when trying to run ```pip install pipenv```. To fix this, try running this instead:
    ```
    pip3 install pipenv
    ```
    Your computer may recognize ```pip3``` as a command instead of ```pip```, so give it a shot.

## Setting up and getting to know roster.xlsx
(row, column)
- enter your name in cell (1, 1)
- enter your building in cell (1, 2) with format WDN or WDS
- enter your resident names into column 1 starting from row 2. you can copy & paste from your own excel if you have one and it should format correctly. If you have relatively fewer residents, just leave the remaining rows blank
- enter resident rooms into column 2 starting from row 2. you can leave it as 101Aa, etc. no need to remove the Aa (unless you want to). make sure they are with the correct resident name
<img width="185" alt="Screen Shot" src="https://user-images.githubusercontent.com/99346474/228655994-db289f09-033a-4284-a386-c84251f43c2c.png">

- the number in cell (1, 25) plus one is the number of the row of the first resident that it will fill out the form for when the program is run. For instance, when it is 1, the first resident it will fill out the form for is the resident on row 1 + 1 = 2. You will usually start at 1, and the automater will skip over residents with dates that you have not inputted. This is here so that if the program stops running for some reason, you can rerun it starting from the last row it completed. Don't forget to set it back to 1 once you finish.
- if you want to receive emails of the form response, put your email in cell (1, 26). if not, leave it blank. i recommend it so you can see which ones were submitted and what the inputs were, but if you're not a fan of spam emails, don't do it.

All of these will remain the same every time you use this automater (except the number in cell (1, 25), which you will only change if the program stops running before it finishes)

## Filling out the excel
- In column 3 and 4, labelled Date 1 and Date 2, enter the date number of your interaction, you can use single digits. Don't enter the month or year. For instance, if you've had two interactions with one resident on January 12 and January 24, in the Date 1 column, put 12, and Date 2 column, put 24 (or the other way around if that's your cup of tea). This is assuming that you are filling out this form within the same month of the interaction. Otherwise, sorry :( (maybe in the future I will implement choosing the month). If you only have one interaction, put the date into one column and leave the other blank. If you have no interactions with them, leave both blank.
- In columns 5 through 10 (inclusive), mark the cell with the method of your interaction. They are labelled at the top row. You can mark with any text.
- In columns 11 through 17, mark the cell with the topic of your interaction.
- If you selected Intentional Conversation theme, in columns 18 through 23, mark the cell with the topic of your intentional conversation. If you didn't select intentional conversation theme, leave them blank.
<img width="329" alt="Screen Shot" src="https://user-images.githubusercontent.com/99346474/228656254-c4dabf05-139d-4fca-85fa-0dfa3ed5abf3.png">

- If you have any comments on the interation, type it into the 24th column labelled comments. The excel won't show your whole comment if it's too long, for aesthetic purposes, but it's there. Leave it blank for no comment
- Repeat for each resident you want to fill out the form for.

## Installing requirements, Running the program
You only need to do steps 1 and 2 once, the first time.\
When you want to run the program again, go into the directory again (use cd <path> to navigate through your files) and start from step 3.\
In your terminal, wherever you want to hold this program, run these commands:
1. Download the repository into your computer and enter
```
git clone https://github.com/kathiehyu/Interaction-Automater.git
cd Interaction-Automater
```
\
2. Install pipenv
```
pip install pipenv
```
\
3. Open the virtual environment to run the program
```
pipenv shell
```
\
4. Update the dependencies
```
pipenv update
```
### 5. Running the program
Make sure that you've properly updated the roster.xlsx according to the directions above
```
python3 interactionAutomater.py
```
- the program will create a new window. You don't have to stay on the window, but you can't switch to another desktop or the program will stop running
- the terminal will have records of the inputs, you can use it to make sure it's on the right track, and if the program stops running, you can look at the last completed resident, and set the cell (1, 25) to that row minus one. (keep in mind that that number **plus one** is the row that it will start at)
## Technologies
- Selenium, Python
## To Do
- [ ] Create GUI

# Scheduler
Takes the input of a students name, session, and needed classes and attempts to assign them to a valid schedule at School X.

Keeps track of the number of students assigned to each period for each class and tries to put students in the least crowded ones. 
Attempts to schedule classes one at a time in this fashion in random order until a valid schedule is produced.

Produces a master student roster in Excel and an individual student schedule in Word

Should be paired with Rostermaker.py and Unscheduler.py to edit and format its output

known issues: 'while' loops will get stuck if an unsolvable schedule is given.
Cancel the program with ctrl+c and try again with alternate classes.

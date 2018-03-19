# Scheduler
#!/usr/bin/python3



#scheduler - Takes the input of a students name, session, and needed classes and attempts to
#           assign them to a valid schedule at School X.

#           Keeps track of the number of students assigned to each period for each class and
#           tries to put students in the least crowded ones. Attempts to schedule classes
#           one at a time in this fashion in random order until a valid schedule is produced.

#           Produces a master student roster in Excel and an individual student schedule in Word

#           Should be paired with Rostermaker.py and Unscheduler.py to edit and format its output

# known issues: 'while' loops will get stuck if an unsolvable schedule is given.
#               Cancel the program with ctrl+c and try again with alternate classes.
#              Use in the wild has led to imbalanced classrooms. Hard to monitor and maintain
#              

import shelve, sys, os
import openpyxl, docx
from random import randint
from random import shuffle
os.chdir('C:\\Programs')


###this creates the roomcount variables
##shelfFile = shelve.open('periodcounter')
##rooms=['112','113','114','115','108','109']
##periodlist = []
##for x in rooms: #rooms got deleted so it needs to be fixed if remaking
##    for y in range(1,7):s
##          periodlist.append('%sPeriod%sAM'% (x, str(y)))
##    for z in range(1,7):
##            periodlist.append('%sPeriod%sPM'% (x, str(z)))
##    for x in periodlist:
##        shelfFile[x]=0
##
##print(list(shelfFile.keys()))
##print(list(shelfFile.values()))




classinfo = [
            {'subject':'Reading 1', 'room':'114', 'periods':[1,3,4,6], 'teacher':'Mrs. Kat'}, # This list of dictionaries hold all the info for the school schedule setup.
                                                                                            
            {'subject':'Reading 2', 'room':'108', 'periods':[3,4],'teacher':'Mr. G/Mrs. A'},
            {'subject':'Leyendo PM', 'room':'114', 'periods':[1,2,3,4], 'teacher':'Mrs. B'},
            {'subject':'Government', 'room':'108', 'periods':[5,6],'teacher':'Mrs. A'},
            {'subject':'Music Appreciation', 'room':'108', 'periods':[5],'teacher':'Mrs. A'},
            {'subject':'English', 'room':'109', 'periods':[1,2,3,4,5,6],'teacher':'Mr. C, Mr. D'},
            {'subject':'Math1', 'room':'112', 'periods':[3,4,5,6],'teacher':'Mr. E'},
            {'subject':'Math2', 'room':'112', 'periods':[1,2],'teacher':'Mr. E},
            {'subject':'Chemistry', 'room':'112', 'periods':[1,2,3,4],'teacher':'Mr. E'},
            {'subject':'Science', 'room':'113', 'periods':[3,4,5,6],'teacher':'Ms. F'},
            
            {'subject':'Ciencias PM', 'room':'113', 'periods':[1,2,3,4,5,6],'teacher':'Ms. F'},
            
            {'subject':'History', 'room':'113', 'periods':[1,2,3,4,5,6],'teacher':'Mr. H'},
            {'subject':'Historia PM', 'room':'113', 'periods':[1,2,3,4],'teacher':'Mr. H'},
            {'subject':'Career', 'room':'115', 'periods':[1,2,5,6],'teacher':'Mr. G'},
            {'subject':'Trabajo PM', 'room':'115', 'periods':[1,2],'teacher':'Mr. G'},
            {'subject':'Elective', 'room':'115', 'periods':[1,2,3,4,5,6],'teacher':'Mr. I'},
            {'subject':'HOPE', 'room':'109', 'periods':[3,4,5,6],'teacher':'Mr. D'},
            ]



def scheduler(name,am, room): #defines the main function of the program. This function assigns the schedule and writes it to word/excel. Called by Tkinter UI
    #name argument is the name of the student
    #am argument is either AM or PM session
    #room argument is the classes the user wants the student to have
    inputlist = []
    for x in room:  
        inputlist.append(x) #converts room argument into a list
    am = am[0] #converts list to string

    

  
    shelfFile2 = shelve.open('roster') #checks if the student is already scheduled. 
    if name in list(shelfFile2.keys()):
        print( '%s is already scheduled' % name)
        shelfFile2.close()
        return

    shelfFile = shelve.open('periodcounter') #opens the persistent list of number of students registered in each class
    


# ready to begin looping
    sp = ['empty','empty','empty','empty','empty','empty'] #first round schedule list. Entries often overwritten
    spfinal = ['empty','empty','empty','empty','empty','empty'] #final schedule list. Should be rewritten on only one loop

    def subject(inputs, am):  # Defines the function which is the main engine for assigning classes. It is called by isLegit loop.
        
        for x in range(len(classinfo)): 
            info = classinfo[x]
            if (inputs ==  info['subject']) and ('Reading 1' in inputs): #assignment of reading 1 class
                
                import shelve
                shelfFile = shelve.open('periodcounter')
                

                classvalues = [1000, 1000, 1000, 1000, 1000, 1000] #default list of number of students in each period. Valid classes are much lower numbers, so it avoids using them.
                                                                    #1000 numbers are placeholders to keep the list in the 6 period format.

                for i in info['periods']: #populates the default classvalues list with the actual class counts
                    
                    if i<3: #makes sure only the 'starter' periods are counted
                        classvalues[i-1]=shelfFile['%sPeriod%s%s' %(info['room'],str(i),am)]
                       
                
                while info['subject'] not in str(sp): #loop that attempts to assign the reading class to the least crowded period. 

                    lowest = classvalues.index(min(classvalues)) #finds the class with the least number of students
                    if sp[lowest] == 'empty':
                        if lowest == 0 and sp[3] == 'empty': #makes sure both period 1 and 4 are empty before assigning.
                            
                            sp[lowest] = info['subject']
                            spfinal[lowest] = info['subject']
                            sp[3] = info['subject']
                            spfinal[3] = info['subject']
                            
                        elif lowest == 1 and sp[4] == 'empty': #makes sure both period 2 and 5 are empty before assigning.
                            
                            sp[lowest] = info['subject']
                            spfinal[lowest] = info['subject']
                            sp[4] = info['subject']
                            spfinal[4] = info['subject']
                        elif lowest == 2 and sp[5] == 'empty':#makes sure both period 3 and 6 are empty before assigning.
                            
                            sp[lowest] = info['subject']
                            spfinal[lowest] = info['subject']
                            sp[5] = info['subject']
                            spfinal[5] = info['subject']
                        elif (sp[0] != 'empty' or sp[3] != 'empty') and (sp[1] != 'empty' or sp[4] != 'empty') and (sp[2] != 'empty' or sp[5] != 'empty'): #creates a reading slot in an invalid manner to break the loop. MUST be caught by the isFalse check.
                            
                            sp[lowest] = info['subject']+'Fake' #should be changed to 'Fugazi'
                            spfinal[lowest] = info['subject']+'Fake'
                            
                        else:
                            classvalues[lowest] += 100  
                    else:
                        classvalues[lowest] += 100 #temporarily inflates the student count for the lowest slot so it doesnt get picked again


            elif (inputs ==  info['subject']) and ('Reading 2' in inputs): #assignment of reading 1 class
                
                import shelve
                shelfFile = shelve.open('periodcounter')
                

                classvalues = [1000, 1000, 1000, 1000, 1000, 1000] #default list of number of students in each period. Valid classes are much lower numbers, so it avoids using them.
                                                                    #1000 numbers are placeholders to keep the list in the 6 period format.

                for i in info['periods']: #populates the default classvalues list with the actual class counts
                    if i<4: #makes sure only the 'starter' periods are counted
                        classvalues[i-1]=shelfFile['%sPeriod%s%s' %(info['room'],str(i),am)]
                       
                
                while info['subject'] not in str(sp): #loop that attempts to assign the reading class to the least crowded period. 



                    lowest = classvalues.index(min(classvalues)) #finds the class with the least number of students
                    
                    if sp[lowest] == 'empty':
                        
                        if lowest == 2 and sp[3] == 'empty': #makes sure both period 1 and 4 are empty before assigning.
                            
                            sp[lowest] = info['subject']
                            spfinal[lowest] = info['subject']
                            sp[3] = info['subject']
                            spfinal[3] = info['subject']
                        elif sp[lowest] == 'empty': #creates a reading slot in an invalid manner to break the loop. MUST be caught by the isFalse check.
                            
                            sp[lowest] = info['subject']+' Fake' #should be changed to 'Fugazi'
                            
                            spfinal[lowest] = info['subject']+' Fake'
                            
                            
                                
                        else:
                            classvalues[lowest] += 100
                    else:
                        classvalues[lowest] += 100 
                        
        
         
            elif (inputs ==  info['subject']) and ('Leyendo PM' in inputs): #assignment of reading 1 class
                
                import shelve
                shelfFile = shelve.open('periodcounter')
                

                classvalues = [1000, 1000, 1000, 1000, 1000, 1000] #default list of number of students in each period. Valid classes are much lower numbers, so it avoids using them.
                                                                    #1000 numbers are placeholders to keep the list in the 6 period format.

                for i in info['periods']: #populates the default classvalues list with the actual class counts
                    if i<4: #makes sure only the 'starter' periods are counted
                        classvalues[i-1]=shelfFile['%sPeriod%s%s' %(info['room'],str(i),am)]
                       
                
                while info['subject'] not in str(sp): #loop that attempts to assign the reading class to the least crowded period. 



                    lowest = classvalues.index(min(classvalues)) #finds the class with the least number of students
                    
                    if sp[lowest] == 'empty':
                        
                        if lowest == 2 and sp[3] == 'empty': #makes sure both period 1 and 4 are empty before assigning.
                            
                            sp[lowest] = info['subject']
                            spfinal[lowest] = info['subject']
                            sp[3] = info['subject']
                            spfinal[3] = info['subject']
                        elif lowest == 1 and sp[2] == 'empty': #makes sure both period 2 and 3 are empty before assigning.
                            
                            sp[lowest] = info['subject']
                            spfinal[lowest] = info['subject']
                            sp[4] = info['subject']
                            spfinal[4] = info['subject']

                        elif sp[lowest] == 'empty': #creates a reading slot in an invalid manner to break the loop. MUST be caught by the isFalse check.
                            
                            sp[lowest] = info['subject']+' Fake' #should be changed to 'Fugazi'
                            
                            spfinal[lowest] = info['subject']+' Fake'
                        
                            
                                
                        else:
                            classvalues[lowest] += 100
                    else:
                        classvalues[lowest] += 100   
 


            elif inputs ==  info['subject']: # assignment for regular classes. Pretty much the same as above.
                import shelve
                shelfFile = shelve.open('periodcounter')
                
                classvalues = [1000, 1000, 1000, 1000, 1000, 1000]

                for i in info['periods']:
                    classvalues[i-1]=shelfFile['%sPeriod%s%s' %(info['room'],str(i),am)]
                
                while info['subject'] not in sp:

                    lowest = classvalues.index(min(classvalues))

                    if sp[lowest] == 'empty':
                        sp[lowest] = info['subject']
                        spfinal[lowest] = info['subject']
                        
                    else:
                        classvalues[lowest] += 100
                        
    
    
    isLegit = False 
    

    while isLegit == False: #keeps trying the scheduler function until it finds a valid one.
        shuffle(inputlist) #shuffles the class inputs so that resulting schedules are not always the same.
       

        for inputs in inputlist: #Runs the scheduler using each of the user inputted classes.
            subject(inputs,am)
            

            
      
            #resets sp for double classes. Generalizes the names of the classes so that the while loops wont be satisfied by the first instance of, say, English 1
            for i in range(6): 
                if sp[i] != 'empty':
                    sp[i]='full'
        #checks if spfinal is a legit schedule
        whitelist1 = ['empty'] #remake into a list of lists
        whitelist2 = ['empty']
        whitelist3 = ['empty']
        whitelist4 = ['empty']
        whitelist5 = ['empty']
        whitelist6 = ['empty']
        for i in range(len(classinfo)): #populates the whitelist
            info = classinfo[i]
            for x in info['periods']:
                if x == 1:
                    whitelist1.append(info['subject'])
                if x == 2:
                    whitelist2.append(info['subject'])
                if x == 3:
                    whitelist3.append(info['subject'])
                if x == 4:
                    whitelist4.append(info['subject'])
                if x == 5:
                    whitelist5.append(info['subject'])
                if x == 6:
                    whitelist6.append(info['subject'])


        if (spfinal[0] in str(whitelist1) and spfinal[1] in str(whitelist2) and spfinal[2] in str(whitelist3) and spfinal[3] in str(whitelist4) and spfinal[4] in str(whitelist5) and spfinal[5] in str(whitelist6)) and ('Fake' not in str(spfinal)): 
            isLegit= True #ends the loop and pushes it on to saving and printing. #'Fake' condition catches bad reading assignments.

        if isLegit == False: #clears a failed attempt
            sp = ['empty','empty','empty','empty','empty','empty']
            spfinal = ['empty','empty','empty','empty','empty','empty']



    for i in range(6): #increases the permanent student number counter for each class once the schedule is finalized.
        for x in range(len(classinfo)):
            info = classinfo[x]
            if spfinal[i] == info['subject']:
                shelfFile['%sPeriod%s%s' %(info['room'],(i+1),am)] += 1

        
    shelfFile2[name] = [name, am, spfinal[0], spfinal[1], spfinal[2], spfinal[3], spfinal[4], spfinal[5]] #saves the kid's schedule 

    print(name) #prints the output of the schedule operation for quick review.
    print(am)
    for i, x in list(enumerate(spfinal)): #prints the classes
        print('Period ' +str(i+1) + ': ' + x)

    #prints permanent class count for monitoring purposes
    for i in range(6):
        for x in range(len(classinfo)):
            info = classinfo[x]
            if spfinal[i] == info['subject']:
               print('the count for %s period %s is: %s' % (info['subject'], str(i+1),shelfFile['%sPeriod%s%s' %(info['room'],(i+1),am)])) 

    
#writes to an excel sheet
    
    wb = openpyxl.load_workbook('Schedule.xlsx')
    sheet = wb.active
    blank = sheet.max_row + 1
    for c in range(6):
        
        sheet.cell(row=blank, column=1).value = name
        
        sheet.cell(row=blank, column=2).value = am
        spfinal[c]
        sheet.cell(row=blank, column=(c+3)).value = spfinal[c]
    wb.save('Schedule.xlsx')
    wb.close()
#writes to a word document
    
    doc=docx.Document('C:\\Programs\\Schedules\\blank.docx')
    doc.paragraphs[3].runs[2].text = name #name
    doc.paragraphs[6].add_run(' '+ '__' + am + '___').underline = True #session
    doc.tables[0].columns[1].cells[1].paragraphs[0].add_run(spfinal[0]).bold = True #classes
    doc.tables[0].columns[1].cells[2].paragraphs[0].add_run(spfinal[1]).bold = True
    doc.tables[0].columns[1].cells[3].paragraphs[0].add_run(spfinal[2]).bold = True
    doc.tables[0].columns[1].cells[4].paragraphs[0].add_run(spfinal[3]).bold = True
    doc.tables[0].columns[1].cells[5].paragraphs[0].add_run(spfinal[4]).bold = True
    doc.tables[0].columns[1].cells[6].paragraphs[0].add_run(spfinal[5]).bold = True
    for y in range(1,7):
        for x in range(len(classinfo)):
            info = classinfo[x]
            if doc.tables[0].columns[1].cells[y].text == info['subject']:
                doc.tables[0].columns[4].cells[y].paragraphs[0].add_run(info['teacher']).bold = True #writes the correct teacher name
                doc.tables[0].columns[5].cells[y].paragraphs[0].add_run(info['room']).bold = True #writes the right room name       

    doc.save('C:\\Programs\\Schedules\\%s.docx' % name)
    
    shelfFile2.close()
    shelfFile.close()

#the user interface

from tkinter import *



def list_of_entries(): #connected to the submit button on interface
    
    room = [] #list of inputs to be scheduled
    session=[]
    name = e1.get()

    for x in range(len(classinfo)):  #assigns names to checked boxes and adds them to room argument list
        info = classinfo[x]
        if varia[x].get() == 1:
            room.append(info['subject'])
        if x in twiceindexlist:
            y = twiceindexlist.index(x)
            if twicevaria[y].get()== 1:
                room.append(info['subject'])
    


    if var17.get() == 1: #am or pm arguments
        session.append('AM')

    if var18.get() == 1:
        session.append('PM')

    
    
    if len(room)>6:  #prevents some invalid inputs.
        print('You have chosen too many classes') 
        return
    while len(room)<= 6:
        room.append('empty')
    if len(session)<1:
        print('please provide a session')  
        return

    
        
    scheduler(name, session, room) #runs the main function
    print ('done') 
    
    
def unschedule(): #unscheduler function. Removes the student from the schedule master list as well as the permanent class counts
    
    name = e1.get()
    shelfFile = shelve.open('periodcounter')
    shelfFile2 = shelve.open('roster')
    
    #removes from schedule excel
    wb = openpyxl.load_workbook('Schedule.xlsx')
    sheet = wb.active
    savelist= []
    newwb= openpyxl.Workbook()
    newsheet = newwb.active

    for r in range(1,sheet.max_row + 1):
        if sheet.cell(row=r, column=1).value != name:
            for i in range(1,8):
                newsheet.cell(row=r,column=i).value = sheet.cell(row=r,column=i).value
    assigned = shelfFile2[name]
    
    for i in range(2,8):  #uncounts  
        for x in range(len(classinfo)):  
            info = classinfo[x]
            if assigned[i] == info['subject']:
                shelfFile['%sPeriod%s%s' %(info['room'],str(i-1),assigned[1])] -= 1
    os.remove('C:\\Programs\\Schedules\\%s.docx' % name)
    shelfFile2.pop(name)
    wb.close()
    newwb.save('Schedule.xlsx')
    newwb.close()
    shelfFile2.close()
    shelfFile.close()
    print('done')

master= Tk()

varia = [] #makes a list of intvars based on the number of classes in the school
for x in range(len(classinfo)): 
    varia.append(IntVar())


#trying to make appropriate x2 buttons
twiceindexlist = []
for x in range(len(classinfo)): #makes a list of which classinfo indexes are x2 able
    info = classinfo[x]

    isReading = False
    isReadingOkay = False
    isTwiceOkay = False
   
    if 'Reading' in info['subject']: #checks if the class is reading
        isReading = True
    if isReading == True and len(info['periods'])>3: #checks to see if there are 2 or more reading pairs
        isReadingOkay = True
        
    elif (isReading == False) and (len(info['periods'])>1): #checks if there are 2 or more periods for a nonreading class
        isTwiceOkay = True
    if isReadingOkay == True or isTwiceOkay == True: #adds the index of valid x2 classes to the list
        twiceindexlist.append(x)

twicevaria = [] #gives the x2 buttons an IntVar
for x in range(len(twiceindexlist)):
    twicevaria.append(IntVar())




var17 = IntVar() #AM and PM session
var18 = IntVar()

#makes the actual input window
e1 = Entry(master) #name window
e1.grid(row=1, column=0)

Label(master, text="Name").grid(row=0, sticky = W)

c=[] #makes the checkboxes based on the classinfo list
twicecount = 0 #counter for twicevaria list index

for x in range(len(classinfo)):
    info = classinfo[x]
    check = Checkbutton(master,text=info['subject'], variable=varia[x]).grid(row=(x+3), column = 0, sticky=W) #regular checkboxes
    c.append(check)
    if x in twiceindexlist:
        xcheck = Checkbutton(master,text='x2', variable=twicevaria[twicecount]).grid(row=(x+3), column = 1, sticky=W) #x2 checkboxes
        twicecount+=1 #connects the x2 checkboxes to successive twicevaria IntVars.

c17 = Checkbutton(master, text="AM", variable=var17).grid(row=2, column = 0, sticky=W) #am and pm buttons
c18 = Checkbutton(master, text="PM", variable=var18).grid(row=2, column = 1, sticky=W)



Button(master, text='Schedule', command=list_of_entries).grid(row=(len(classinfo)+3), column=0, sticky=W, pady=4) #button that starts it all
Button(master, text='UnSchedule', command=unschedule).grid(row=(len(classinfo)+3), column=1, sticky=W, pady=4)



    




mainloop( ) #housekeeping



            

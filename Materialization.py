import pandas as pd
import os
from openpyxl import load_workbook
from datetime import datetime
import matplotlib.pyplot as plt
import numpy as np

"""
The function that starts it all.
"""
def beginning(file1,file2,file3,file4):
  print("\nMATHEMATICS | PROGRAMMING | PROFESSION | FINANCE\n")

  files = [file1,file2,file3,file4]
        
  sidesofsquare = {"MATHEMATICS":mathematics,
                   "PROGRAMMING":programming,
                   "PROFESSION":profession,
                   "FINANCE":finance}
                        
  choice = input("Side: ").upper()
  
  if choice in sidesofsquare:
      pass
  else:
    print("\nAre you cognitively impaired?...\n")
    return
 
  keys_list = list(sidesofsquare.keys()).index(choice)
  sidesofsquare[choice](choice,files[keys_list])

"""
writing() will be used to record instances of points of time
"""
def writing(self,action):
 excel_file = self.file

 book = load_workbook(excel_file)

  # Getting instance of point of time
 now = datetime.now()

 df = pd.DataFrame({
     "Year": [now.year],
     "Month": [now.month],
     "Day": [now.day],
     "Hour":[now.hour],
     "Minute": [now.minute]})

 if action not in book.sheetnames:
     ws = book.create_sheet(action)
  
 else:
     ws = book[action]

 max_row = ws.max_row

 if all(ws.cell(row=max_row,column=col).value is None for col in range(1,6)):
     pass
 else:
     max_row += 1

 for index, row in df.iterrows():
     for col_num, value in enumerate(row):
         ws.cell(row=max_row + index, column=col_num+1, value=value)

 book.save(excel_file)
 book.close()



"""
--------------------
Mathematics Quadrant
--------------------
"""
class mathematics():
  def __init__(self,choice,file):

    self.choice = choice
    self.file = file
    print("\nOUTSET | HALT | VIEW | MILESTONE | TESTING\n")

    mathematicsSides = {"OUTSET":self.Outset,
                        "HALT":self.Halt,
                        "VIEW":self.View,
                        "MILESTONE":self.Milestone,
                        "TESTING":self.Test}

    mathematics_choice = input(f"{choice}-Side: ").upper()

    if mathematics_choice in mathematicsSides:
      mathematicsSides[mathematics_choice]()
    else:
      print("\nOh, I see, you are a dumbass...\n")

  """
  Annotates the beginning instance time of the Quadrant
  """
  def Outset(self):
    action = "OUTSET"
    writing(self,action)

  """
  Annotates the final instance time of the Quadrant
  """
  def Halt(self):
    action = "HALT"
    writing(self,action)


  """
  Portrays the amount of hours spend on the Quadrant, and a
  Cartesian Plane (only the positive values), or also known as
  a scatter plot of the amount of time (in hours) with respect to
  the time of the day the quadrant ended.
  """
  def View(self):
    # I. First we read the excel that stores the parameters
    view_outsets = pd.read_excel(self.file,sheet_name=0,header=None)
    view_halts = pd.read_excel(self.file,sheet_name=1,header=None)

    # II. We prepare to compare by listing all instances
    outsets = view_outsets.values.tolist()
    halts = view_halts.values.tolist()

    # III. Compare hours and minutes of outsets and halts for further ploting

    xaxis = []
    yaxis = []

    for x, y in zip(outsets,halts):
       hour = float(y[3]-x[3])
       minute = round((y[4]/60)-(x[4]/60),3)
       hours = hour + minute
       xaxis.append(round(hours,3))
       yaxis.append(round(y[3]+(x[3]/60),3))

    totalhours = round(sum(xaxis),2)
      
    # IV. Now we need the x and y axis; i.e hours, time of the day respectively
    a = np.array(xaxis)
    b = np.array(yaxis)
    plt.scatter(a,b)
    plt.title(f"Total Hours {totalhours}")
    plt.xlabel("Hours")
    plt.ylabel("Ending Day Time")
    plt.show()

          
      
          
       

  def Milestone():
    pass

  def Test():
    pass

"""
--------------------
Programming Quadrant
--------------------
"""
class programming():
  def __init__(self,choice,file):
    self.choice = choice
    self.file = file
      
    print("\nOUTSET | HALT | VIEW | MILESTONE")
    programmingSides = {"OUTSET":self.Outset,
                        "HALT":self.Halt}
    programming_choice = input(f"{choice}-Side: ").upper()

    if programming_choice in programmingSides:
        programmingSides[programming_choice]()
    else:
        print("\nI recognize a dumbass when I see one...\n")


  def Outset(self):
    action = "OUTSET"
    writing(self,action)

  def Halt(self):
    action = "HALT"
    writing(self,action)

"""
-------------------
Profession Quadrant
-------------------
"""
class profession():
  def __init__(self,choice,file):
    print("\nOUTSET | HALT | VIEW | MILESTONE")
    professionSides = {"OUTSET":programming.Outset}

  def Outset():
    pass



"""
----------------
Finance Quadrant
----------------
"""
class finance():
  def __init__(self,choice):
    print("\nOUTSET | HALT | VIEW | MILESTONE")
    financeSides = {"OUTSET":programming.Outset}

  def Outset():
    pass




if __name__ == "__main__":
    # File directory in your machine
    directory = "/Users/damiamalfaro/Desktop/Materialization/Excel"
    
    # Allocates Excel Files into the program
    files = os.listdir(directory)
    sorted_files = sorted(files, key=lambda x: int(x[0]))
    file_dictionary = {}
    for i, file in enumerate(sorted_files):
        file_dictionary[f"file_{i+1}"] = os.path.join(directory, file)

    all_files = [i for i in file_dictionary.values()]
        
    # Outset    
    beginning(*all_files)
























































import pandas as pd
import os
from openpyxl import load_workbook
from datetime import datetime


"""
The function that starts it all.
"""
def beginning(file1,file2,file3,file4):
  print("MATHEMATICS | PROGRAMMING | PROFESSION | FINANCE")

  files = [file1,file2,file3,file4]

  sidesofsquare = {"MATHEMATICS":mathematics,
                   "PROGRAMMING":programming,
                   "PROFESSION":profession,
                   "FINANCE":finance}

  choice = input("Side: ").upper()
  
  print(sidesofsquare)



  keys_list = list(sidesofsquare.keys()).index(choice)

  if choice in sidesofsquare:
    sidesofsquare[choice](choice,files[keys_list])
  else:
    print("Are you cognitively impaired...")
    

"""
writing() will be used to record instances of points of time
"""
def writing(self,action):
    # Importing & Converting Excel File
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
Mathematics
"""
class mathematics():
  def __init__(self,choice,file):

    self.choice = choice
    self.file = file
    print("\nOUTSET | HALT | VIEW | MILESTONE | TESTING")

    mathematicsSides = {"OUTSET":self.Outset,
                        "HALT":self.Halt}

    mathematics_choice = input(f"{choice}-Side: ").upper()

    if mathematics_choice in mathematicsSides:
      mathematicsSides[mathematics_choice]()
    else:
      print("\nOh, I see, you are a dumbass...")

  def Outset(self):
    action = "OUTSET"
    writing(self,action)

  def Halt(self):
    action = "HALT"
    writing(self,action)

  def View():
    pass

  def Milestone():
    pass

  def Test():
    pass

"""
Programming
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


class profession():
  def __init__(self,choice,file):
    print("\nOUTSET | HALT | VIEW | MILESTONE")
    professionSides = {"OUTSET":programming.Outset}

  def Outset():
    pass


class finance():
  def __init__(self,choice):
    print("\nOUTSET | HALT | VIEW | MILESTONE")
    financeSides = {"OUTSET":programming.Outset}

  def Outset():
    pass




if __name__ == "__main__":
    # File directory in your machine
    directory = ""

    # Allocates Excel Files into the program
    files = os.listdir(directory)
    sorted_files = sorted(files, key=lambda x: int(x[0]))
    file_dictionary = {}
    for i, file in enumerate(sorted_files):
        file_dictionary[f"file_{i+1}"] = os.path.join(directory, file)

    all_files = [i for i in file_dictionary.values()]
        
    # Outset    
    beginning(*all_files)
























































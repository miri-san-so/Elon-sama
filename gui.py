from tkinter import Tk, Frame, Label, Image, Pack, Canvas, Scrollbar, Listbox, Button
from PIL import ImageTk, Image
import os
from retrieve import getTitleUsingExcel, getHoursUsingExcel
from tkinter import ttk
import os
from webbrowser import open
from generate_graph import giveGraph
import argparse, sys, glob, random
import gc

themePurple = {"canvasBg": "#202020", "headerBg":"#191919", "appName":"#AD5CC1", "windowTitle":'#AD5CC1', "tasksFrame":"#1A1A1A", "tasks":"#701587"}
themeCyan = {"canvasBg": "#202020", "headerBg":"#191919", "appName":"#5AFFC4", "windowTitle":'#31E4A4', "tasksFrame":"#1A1A1A", "tasks":"#009C80"}
themeLight = {"canvasBg": "#F0F0F0", "headerBg":"#FFFFFF", "appName":"#303030", "windowTitle":'#303030', "tasksFrame":"#FFFFFF", "tasks":"#303030"}
def openExcel():
    excelFile = (os.getcwd()+"\\tasks.xlsx")
    open(excelFile)
    
def giveElon():
    elonPNG = []
    elons = "C:\\Users\\SANFAL\\Desktop\\elon-sama\\python files\\elons"
    for root, dirs, files in os.walk(elons):
        for file in files:
            if file.endswith(".png"):
                elonPNG.append(os.path.join(root, file))
        gc.collect()
    return(elonPNG)

def createWindow(args):
    
    # CLI argument for changing theme
    if args.t == "cyan":
        theme = themeCyan
    if args.t == "purple":
        theme = themePurple
    if args.t == "light":
        theme = themeLight
        
    # Creating Root Window
    tk = Tk()

    # Setting Title
    tk.title('Elon sama')
    
    # Setting Geometry of the Window
    tk.geometry('1200x700')
    
    # Making the window resizable only Verticaly / on Y-Axis Only
    tk.resizable(False, True)
    
    # Changing the background of window
    tk.configure(background=theme["canvasBg"])

    # Header for the app
    headbar = Frame(tk, width=1200, height=77, bg=theme["headerBg"])
    headbar.pack(side="top")


    # Elon Sama Title
    appName = Label(tk, text="ELON SAMA.", font=(
        "Arial", 15, "bold"), bg=theme['headerBg'], fg=theme["appName"])
    appName.place(x=23, y=23)


    # Today (Window Name)
    today = Label(tk, text="Today", font=(
        "Arial", 36, "bold"), bg=theme['canvasBg'], fg=theme["windowTitle"])
    today.place(x=23, y=97)

    # Spawning Elon (adding elon to the bottom right of the window)
    frame = Frame(tk)
    frame.place(x=826, y=438)

    canvas = Canvas(frame, bg=theme['canvasBg'], width=374,
                    height=262, bd=0, relief='ridge', highlightthickness=0)
    canvas.pack()

    # > selecting random elon musk image from ./elons/ folder
    filePath = giveElon()
    newElon = random.choice(filePath)

    # Setting Images on Canvas
    photoimage = ImageTk.PhotoImage(file=newElon)
    
    # Adding image to bottom right
    canvas.create_image(187, 131, image=photoimage)


    # ->> Creating tasks <<-
    
    # Values for Placing title and sections
    initialYValue = 173
    initialTitleValue = 190

    # tasks and hour values
    taskss = getTitleUsingExcel(5)  # MAKE THIS DYNAMIC 
    hourss = getHoursUsingExcel(5)  # MAKE THIS DYNAMIC 
    hoursInt = [int(i) for i in hourss]
    
    # Looping Through to Create Tasks
    for i in range(0, 5):   # MAKE THIS DYNAMIC 
       
        # Creating Task Sections / Frame
        task = Frame(tk, width=450, height=60, bg=theme["tasksFrame"])
        task.place(x=23, y=initialYValue)

        # Task Title
        taskTitle = Label(tk, text=taskss[i].capitalize(), fg=theme["tasks"], bg=theme['tasksFrame'], font=("Arial", 14))
        taskTitle.place(x=30, y=initialTitleValue)
        
        # Task Hour
        taskHours = Label(tk, text=hourss[i]+" Hour", fg=theme["tasks"], bg=theme['tasksFrame'], font=("Arial", 14))
        taskHours.place(x=370, y=initialTitleValue)
        initialYValue += 70
        initialTitleValue += 70

    # Open Excel Button
    openExcelButtonFrame = Frame(tk, width=100, height=100)
    openExcelButton = Button(openExcelButtonFrame, text="Open Excel", fg=theme['headerBg'], bg=theme["windowTitle"], width=20, height=2, bd=0, font=("Arial", 10), activebackground=theme["tasks"],command= lambda: openExcel())
    openExcelButton.pack()
    openExcelButtonFrame.place(x=23, y=660)
        
    # Generate Graph Button
    generateGraphButtonFrame = Frame(tk, width=100, height=100)
    generateGraphButton = Button(generateGraphButtonFrame, text="Generate Graph",  fg=theme['headerBg'], bg=theme["windowTitle"], width=20, height=2, bd=0, font=("Arial", 10), activebackground=theme["tasks"], command=lambda:giveGraph(5))
    generateGraphButton.pack()
    generateGraphButtonFrame.place(x=210, y=660)
    
    # This Week Button
    thisWeekFrame = Frame(tk, width=100, height=100)
    thisWeek = Button(thisWeekFrame, text="This Week",  fg=theme['headerBg'], bg=theme["windowTitle"], width=20, height=2, bd=0, font=("Arial", 10), activebackground=theme["tasks"])
    thisWeek.pack()
    thisWeekFrame.place(x=397, y=660)
    
    # Total Hours 
    totalHours = Label(tk, text="Total: "+str(sum(hoursInt))+" Hours", font=("Arial", 36, "bold"), fg=theme["windowTitle"], bg=theme["canvasBg"])
    totalHours.place(x=820, y=97)
  
    # Add Hours Button
    addHoursFrame = Frame(tk, width=100, height=100)
    addHours = Button(addHoursFrame, text="ADD HOURS", fg=theme["tasks"], bg=theme["tasksFrame"], width=30, height=2, bd=0, font=("Arial", 10), activebackground=theme["tasks"])
    addHours.pack()
    addHoursFrame.place(x=930, y=173)
  
    tk.mainloop()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--t', type=str, default="purple")
    args = parser.parse_args()
    createWindow(args)

main()

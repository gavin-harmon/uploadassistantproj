"""
This is the primary script. It is in development currently.

Any changes to this document need to tested to 1) execute in the console 2) execute from the terminal
and 3) have a functional, shareable .exe output via pyinstaller

===================

"""

#libraries and modules used

import os
import sys
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import *
from tkinter import ttk
from pandastable import Table, TableModel
import mouse
import importlib
"""data import modules"""
global vdata, sdata
#set variable

"""controls the working directory for when being processed by document and executable builders"""
os.chdir(os.path.dirname(os.path.abspath(__file__)))

"""Import Submission"""
def fetch_sdata():
    """Loads the input file.

    Before doing so it reads all the file names in "Submission" folder. currently the load process can only
    work with one file, but this may change as the project develops.
  """
    global sdata

    spath = os.path.join(os.path.dirname(__file__), "../Submission")
    print(spath)
    files = os.listdir(spath)

    """Get a list of only excel files in the path, find several extentions formats, case sensitive"""
    files = [files.lower() for files in files]
    files_xls = [f for f in files if f[-3:] in ('lsx' , 'lsm' ,'xls')]

    """     empty list to append to"""
    pathfiles =[]

    """     create list of files with path"""
    for f in files_xls:
        makepathsfiles = os.path.join(str(spath), str(f))
        pathfiles.append(makepathsfiles)
#    """     empty dataframe to append to"""
#    df = pd.DataFrame()
    """     Read Summarize and append to df"""
    for f in pathfiles:
        global sdata
        sdata = pd.read_excel(f, sheet_name ='Ptf_Monitoring_GROSS_Reins', na_values = [0], header=3, converters={'Business Partner Name': str,
            'Type of Business': str, 'Type of Account': str, 'Distribution Type': str, 'LOB': str, 'Distribution Channel': str,
            'Sub LOB': str,'Business Partner ID Number': str,  'Product Name': str, 'Product ID Number': str,  'Product Family': str,
            'Standard Product': str, })
        sdata.columns = sdata.columns.str.strip()

        """Remove rows with null business units"""
        sdata = sdata[sdata['Business Unit'].notnull()]
        return sdata

def fetch_manfields():
    """find mandatory fields"""
    """Import Template"""

    global manfields

    tpath = os.path.join(os.path.dirname(__file__), "../Template")
    files = os.listdir(tpath)

    files = [files.lower() for files in files]

    """Get a list of only excel files in the path, find several extentions formats, case sensitive"""
    template = [f for f in files if f[-3:] in ('lsx', 'lsm', 'xls')]
    template = template[0]

    # template file path
    pathsfile = os.path.join(str(tpath), str(template))

    # empty dataframe
    mdf = pd.DataFrame()

    # read the lines with mandatory descriptions and fields
    mand = pd.read_excel(pathsfile, sheet_name='Ptf_Monitoring_GROSS_Reins', na_values=[0], header=0, nrows=3)
    mdf = mdf.append(mand)

    # Transpose and drop non field name data
    mdf = mdf.T.reset_index()

    # turn to a simple list rather than a pandas dataframe
    manfields = mdf[2][mdf[0] == "Mandatory"].values.tolist()



"""     empty list to append to, eash time fetch each uploaded DataFrame will be a member of
 this list vdata[0], vdata[1] and so on. Reference the lates upload always as vdata[-1]"""
vdata =[]

global sdata
global manfields
fetch_sdata()
fetch_manfields()

vdata.append(sdata.replace(np.nan, '', regex=True))

def validate():
    global vmessagelist, vdflist

    from validations import valid
    valid(vdata[-1], manfields)

    vdata

    import validations

    v1_list = validations.valmessage.values()
    vmessagelist = list(v1_list )

    v2_list = validations.valdf.values()
    vdflist = list(v2_list )

def change_df(self, input_val):
    """ This changes the DataFrame displayed in the UI pandas table. This is likely to be
    rebuilt to support many different dataset displays through validation
    *this runs from buttons as a lambda

    :param self: this is the table that needs to be changed
    :param input_val: a DataFrame, the final dataframe is set in teh below if statement
    :return:
    """
    # Responds to button
    if len(llock) == 0 :
        """Check if locked, this prevents it from running when the user has not explicitly
        indicated they want to reload the data
        
        :functions updateModel, TableModel : These pandastable functions change the DataFrame of the table
        :functions redraw(): replaces the previous pandastable in the UI with the one created to memory in the line above 
        """
        global vdata, sdata, ui_df

        vdata = []
        fetch_sdata()
        vdata.append(sdata.replace(np.nan, '', regex=True))
        ui_df = vdata[-1]
        self.updateModel(TableModel(vdata[-1]))
        self.redraw()
    else:
        lock()


def submitbuttonaction(df):
    """
    see change_df() notes, this completes the same task but is callable from a mouse bound event (see python mouse documentation)
    :param df: current df to be changed
    :return: initial table load or refresh
    """
    if len(llock) == 0 :
        global vdata, ui_df, sdata

        vdata =[]
        fetch_sdata()
        vdata.append(sdata.replace(np.nan, '', regex=True))
        ui_df = vdata[-1]
        ui.updateModel(TableModel(vdata[-1]))
        ui.redraw()

        lock()
    else:
        lock()

"""Button actions, used below in button option "command" , functions exist in scripts,
read functions as 'from [script in this folder] import [DataFrame]'"""

def subfolderopen():
    """System commands to open the folder location. Written to read the
    current file path so that it can work from any location where the folder tree is in tact.
    """
    path = os.path.join(os.path.dirname(__file__), "../Submission").replace('dist', '')
    path = path.replace('dist', '')
    os.startfile(path)

def resourcefolderbuttonaction():
    """System commands to open the folder location. Written to read the
    current file path so that it can work from any location where the folder tree is in tact.
    """
    path = os.path.join(os.path.dirname(__file__), "../Resources").replace('dist', '')
    path = path.replace('dist', '')
    os.startfile(path)

def dltemplatebuttonaction():
    """System commands to open the Template. Written to read the current file
    path so that it can work from any location where the template exists and the folder tree is in tact.
    """
    path = os.path.join(os.path.dirname(__file__), "../Template").replace('dist', '')
    path = path.replace('dist', '')
    file = os.listdir(path)

    filepath = os.path.join(str(path), str(file[0]))

    os.startfile(filepath)

def askassistbuttonaction():
    """imports function from module, see module for notes"""
    from buttonactionemailhelp import emailer
    emailer("", "Upload Application Assistance Required", " <Dana.Mark@allianz.com>;"
                                                          " <angela.chenxx@allianz.com>; <Federico.Guerreschi@allianz.com>; <gavin.harmon@allianz.com>")




# def BR012buttonaction(self):
# # -- inprogress importing messages and data from the 1st validation, not finalizaed
#     if len(llock) == 0 :
#         global messages, mandatoryvalblanks
#         print("xx")
#         from ValBR012 import messages
#         from ValBR012 import mandatoryvalblanks
#         lock()
#     else:
#         lock()

# mouse moves help control responses to user actions actions slight movement to trigger a command
def mousemove(self):
    """
    This is used to trigger as action insequence immediately after a new page has been selected an loaded
    user effect - page navigated away from, waiting for new page to load, as opposed to "push a button and wait for something to happen"
    :param self: the button it is called from
    :return: inperceptable mouse movement
    """
    mouse.move(0, 1, absolute=False, duration=0)

# for is a command needs to run multiple functions *verify this is being used
def combine_funcs(*funcs):
    def combined_func(*args, **kwargs):
        for f in funcs:
            f(*args, **kwargs)
    return combined_func

# empty set to deactivate button response
llock = {}

# if the script has been run on the screen, this will be locked and not run again, the user pressing "back" unlocks this by setting llock={}
def lock():
    global llock
    llock = {1}

def unlock(self):
    global llock
    llock = {}

# The code for changing pages was derived from: http://stackoverflow.com/questions/7546050/switch-between-two-frames-in-tkinter
#Tkinter basics can be found here https://docs.python.org/3/library/tk.html


class root(tk.Tk):
    """this is the main display, it is replaced by other pages as buttons get pushed or lifted to the user display
    The code for changing pages was derived from: http://stackoverflow.com/questions/7546050/switch-between-two-frames-in-tkinter
    Tkinter basics can be found here https://docs.python.org/3/library/tk.html
        """

    def __init__(self, *args, **kwargs):

            tk.Tk.__init__(self, *args, **kwargs)
            container = tk.Frame(self)
            container.pack(side="top", fill="both", expand=True)

            container.grid_rowconfigure(0, weight=1)
            container.grid_columnconfigure(0, weight=1)

            self.frames = {}
            #app title that displays in the window
            self.winfo_toplevel().title("Upload_Assistant")

            # all pages must exist here
            for F in (StartPage,PageSix):
                frame = F(container, self)

                self.frames[F] = frame

                frame.grid(row=0, column=0, sticky="nsew")

            self.show_frame(StartPage)

    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()


class StartPage(tk.Frame):
    """
    in development
    """
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)

        global vdflist, vmessagelist

        submit = tk.Frame(self, height= 850, width=1150, bg='#dedede', )
        submit.pack( fill="both",  expand=True)

        TestApp = tk.Frame(submit, height=600, width=400, bg='#dedede', )
        TestApp.place( relx=0.135, rely=0.2, relwidth=0.7, relheight=0.07)

        label = tk.Label(submit, text="Data Submission Results", font= ('Helvetica', 22), bg='#004a93', fg="white", )
        label.place( relx=0.135, rely=0.1375, relwidth=0.7, relheight=0.07)

        box = tk.Frame(submit, bg='#004a93',  )
        box.place( relx=0.16, rely=0.2875, relwidth=0.65, relheight=0.55)
        validate()

        tablemessage = vmessagelist[6]
        tabledf = vdflist[6]

        labelmessage = tk.Label(submit, text=tablemessage, font= ('Helvetica', 12, ), wraplength=700,  justify= 'center', bg='white', fg='black', anchor= 'center')
        labelmessage.place( relx=0.16, rely=0.225, relwidth=0.65, relheight=0.06, )

        pt = Table(box, dataframe= tabledf , showtoolbar=True, showstatusbar=True)
        pt.show()

        nextbutton = tk.Button(submit, text="Next", font= ('Helvetica', 16), bg='#004a93', fg="white",
                            command=lambda: controller.show_frame(PageFive))
        nextbutton.place( relx=0.395, rely=.85, relwidth=0.18, relheight=0.07, )
        submit.bind("<Enter>", mousemove)
        nextbutton.bind("<Motion>", submitbuttonaction)

        homebutton = tk.Button(submit, text="Start Over", font= ('Helvetica', 16), bg='#004a93', fg="white",
                            command=lambda: controller.show_frame(StartPage))
        homebutton.place( relx=0.395, rely=.925, relwidth=0.18, relheight=0.07, )
        homebutton.bind("<Button-1>", unlock)


class PageSix(tk.Frame):
    """
    in development
    """
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)


        submit = tk.Frame(self, height= 850, width=1150, bg='#dedede', )
        submit.pack( fill="both",  expand=True)

        TestApp = tk.Frame(submit, height=600, width=400, bg='#dedede', )
        TestApp.place( relx=0.135, rely=0.2, relwidth=0.7, relheight=0.07)

        label = tk.Label(submit, text="Data Submission Results", font= ('Helvetica', 22), bg='#004a93', fg="white", )
        label.place( relx=0.135, rely=0.1375, relwidth=0.7, relheight=0.07)

        box = tk.Frame(submit, bg='#004a93',  )
        box.place( relx=0.16, rely=0.2675, relwidth=0.65, relheight=0.55)

        nextbutton = tk.Button(submit, text="Next", font= ('Helvetica', 16), bg='#004a93', fg="white",
                            command=lambda: controller.show_frame(PageFive))
        nextbutton.place( relx=0.395, rely=.85, relwidth=0.18, relheight=0.07, )
        submit.bind("<Enter>", mousemove)
        nextbutton.bind("<Motion>", submitbuttonaction)

        homebutton = tk.Button(submit, text="Start Over", font= ('Helvetica', 16), bg='#004a93', fg="white",
                            command=lambda: controller.show_frame(StartPage))
        homebutton.place( relx=0.395, rely=.925, relwidth=0.18, relheight=0.07, )
        homebutton.bind("<Button-1>", unlock)

app = root()
app.mainloop()

"""
standard tkinter lines that allow the script to run on loop as an application
"""


if __name__ == '__main__':
    app

"""
execution gaurd

"""
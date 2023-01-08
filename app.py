from tkinter import *
from tkinter import ttk
from tkinter import filedialog

import pandas as pd
import time

ROW = 1
RUN = True
# for the loop
ind = 0
selected_item = None
after_id = None

def print_val():
    global RUN
    global ind
    global selected_item
    global after_id
    # indx = ind
    if RUN:

        print("index is ", ind)
        print(tree.item(selected_item[ind])['values'][0])
        ind+=1
        if ind >= len(selected_item):
            RUN=False
    after_id = root.after(2000, print_val)

# Define a function to stop the loop
def on_end():
    global after_id
    global RUN

    if after_id or RUN:
        root.after_cancel(after_id)
        after_id = None
        RUN = False

def auto_all():
    print("Inside auto ", auto.get())
    
def upload():
    print("Inside upload")
    
    global df
    global cols
    filename = filedialog.askopenfilename()
    df = pd.read_excel(filename)
    cols = list(df.columns)
    print(df.shape)
    
    #Set the Menu initially
    col= StringVar()
    col.set("Select Any Column")

    #Create a dropdown Menu
    drop= OptionMenu(root, col,*cols, command=print_col)
    drop.grid(row=3, sticky=W)
    # ROW+=1
    
def print_col(x):
    print(x)
    global col
    col = x

    
    # Add new data in Treeview widget
    # tree["column"] = list(df.columns)
    tree["column"] = [x]
    tree["show"] = "headings"

    # For Headings iterate over the columns
    for col in tree["column"]:
        tree.heading(col, text=col)

    tree.delete(*tree.get_children())
    # Put Data in Rows
    df_rows = df[x].to_numpy().tolist()
    # df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tree.insert("", "end", values=row)
        

        
def loop_tree():
    
    print("loop tree, df.shape= ", df.shape)
    if auto.get():
        # loop through all df
        for i in range(len(df)):
            print("First name: ", df.loc[i, 'first_name'])
            print("Last name: ", df.loc[i, "last_name"])
    else:
        print("AUTO IS OFF OR INACCESSIBLE")
        print("loop: ", col)
        
        global RUN
        global ind
        global selected_item
        RUN = True
        ind=0
        
        
        selected_item = tree.selection()
        
        
        print_val()
        
        # for val in selected_item:
        #     print(tree.item(val)['values'][0])
    
    # loop through treeView
    # time.sleep(int(delay.get()))
    #Create list of 'id's
    '''
    listOfEntriesInTreeView=tree.get_children()

    for each in listOfEntriesInTreeView:
        # print(tree.item(each)['values'][0])
        print(df[df[col]==tree.item(each)['values'][0]])
    '''

def vaidate_num():
    
    if delay.get().isdigit():
        print(delay.get(), " seconds wait")
        
        time.sleep(int(delay.get()))
        loop_tree()
        
    elif delay.get()=="":
        print("<blank delay>")
        loop_tree()
        
    else:
        print(delay.get())
        print("INVALID DELAY INPUT!!!")
        # return False
    
    # return True

# read data
# df = pd.read_excel("UserDetails2.xlsx")

# cols = list(df.columns)

root = Tk()
root.title("Notepad")

auto = IntVar()
check = Checkbutton(root, text="AUTO ALL", variable=auto, command=auto_all).grid(row=1, sticky=W)
ROW+=1

# upload button
ttk.Button(root, text="UPLOAD EXCEL FILE", command=upload).grid(row=2, sticky=W)
ROW+=1




# Add a Treeview widget
tree = ttk.Treeview(root, selectmode='extended')

# Constructing vertical scrollbar with treeview
verscrlbar = ttk.Scrollbar(root,
                           orient ="vertical",
                           command = tree.yview)

verscrlbar.grid(row=4, column=1, sticky='ns')

# horscrlbar = ttk.Scrollbar(root, orient ="horizontal", command = tree.xview)
# horscrlbar.grid(row=5, sticky='ew')

# Configuring treeview
tree.configure(yscrollcommand = verscrlbar.set
               # ,xscrollcommand = horscrlbar.set
              )



tree.grid(row=4, sticky=W)
# ROW+=1


# delay entry
Label(root, text="DELAY").grid(row=6, sticky=W, padx=10)
delay = StringVar()
e=Entry(root, textvariable=delay, width=5).grid(row=6, sticky=E, padx=10)
# reg = root.register(vaidate_num)
# e.config(validate ="key", validatecommand =(reg, '% P'))

# BEGIN & END button
Button(root, text="BEGIN", command=vaidate_num).grid(row=7, sticky=W)
Button(root, text="END", command=on_end).grid(row=7, sticky=E)



mainloop()
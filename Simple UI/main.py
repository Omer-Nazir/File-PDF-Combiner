
import tkinter as tk
import tkinterdnd2 as tkdnd
import tkinter.ttk as ttk
from icon import Icon
import modified_tv as mtv

helptext = """

Supported File Types:
    VALID_WORD_EXTENSIONS = ['.doc', '.docm', '.docx', '.dot', '.dotm', '.dotx', '.odt', '.rtf', '.txt', '.wps', '.xml', '.xps']
    VALID_IMAGE_EXTENSIONS = ['.jpeg', '.jpg', '.png', '.gif']
    VALID_EXCEL_EXTENSIONS = ['.xlsx', '.xlsm', '.xlsb', '.xltx', '.xltm', '.xls', '.xlt', '.xml', '.xlam', '.xla', '.xlw', '.xlr']
    VALID_PPT_EXTENSIONS = ['.ppt','.pptx']

KeyBoard Shortcuts:
> Escape -> Undo selection / cancel bookmark editing
> Enter / Return -> Exceute button / Edit bookmark of selection
> Double-Click on bookmark -> Edit bookmark
> Ctrl + a -> Select all files
> Delete -> Delete Selected files from table

Some Tips:
> Files can be dragged and dropped into the Treeview
> Incase of a Directory Dropped, all the valid files in that directory and its sub directories will be added
> Press Escape or click within treeview anywhere except entries to deselect entries
> Pressing Escape or clicking outside the textbox while editing bookmark will cancel editting of the bookmark
> If a bookmark is empty, it will not be added in final compiled files
> Clicking on Headings will order the items by its respective column
> You can also select and arrange multiple files.
    > Clicking on b1 and then Shift clicking on b2 will select all files between b1 and b2 
    > Ctrl clicking on b1 will add b1 to selection

For better bookmark customization you can download Free Apps like jpdfbookmark and for general pdf tools PDF24

Developer : Omer Nazir
Github : https://github.com/OM3R-Nazir/File-PDF-Combiner
"""

class App(tkdnd.Tk):
    def __init__(self):
        tkdnd.Tk.__init__(self)
        self.title('File - PDF Combiner')
        self.icon = Icon()
        self.tempicon = self.icon.tempicon
        try:
            self.iconbitmap(self.tempicon)
        except:
            pass
        # Screen Grid Configurations

        self.rowconfigure(0, weight=0, minsize=25)
        self.rowconfigure(1, weight=1, minsize=25)
        noofcols = 9

        for i in range(noofcols):
            self.columnconfigure(i, weight=1)

        # Treeview
        self.tv = mtv.modifiedTreeview(self)
        self.tv['columns'] = ('fname','filetype','bookmark','fpath','status')

        # Formatting Columns
        st = True
        self.tv.column("#0", width=0,stretch=0)
        self.tv.column("fname", anchor='w', minwidth=25, stretch=st)
        self.tv.column("filetype", anchor='w', minwidth=25, width = 70, stretch=0)
        self.tv.column("bookmark", anchor='w', minwidth=25, stretch=st)
        self.tv.column("fpath", anchor='w', minwidth=25, stretch=st)
        self.tv.column("status", anchor='w', minwidth=25, width = 70, stretch=0)

        # Formatting Headings
        self.tv.heading("fname",text="File Name",anchor='w')
        self.tv.heading("filetype",text="File Type",anchor='w')
        self.tv.heading("bookmark",text="Bookmark",anchor='w')
        self.tv.heading("fpath",text="File Path",anchor='w')
        self.tv.heading("status",text="Status",anchor='w')

        # Configuring YScrollbar
        yscrollbar = ttk.Scrollbar(self, orient='vertical')
        yscrollbar.configure(command=self.tv.yview)
        self.tv.config(yscrollcommand=yscrollbar.set)

        # Configure Treeview + Scrollbar Grid

        self.tv.grid(row = 1, column = 0, columnspan=noofcols, sticky='news', padx = (20,0), ipadx=10)
        yscrollbar.grid(row = 1,column = noofcols, sticky='ns',padx = (0,20))

        buttontexts = [
            "Add Files","Add Folder","Move up","Move down","Edit Bookmark","Delete","Delete all","Help"
        ]

        buttonfuncs = [
            self.tv.add_files_from_explorer,
            self.tv.add_folder_from_explorer,
            self.tv.move_selection_up,
            self.tv.move_selection_down,
            self.tv.edit_bookmark_of_selected,
            self.tv.delete_selection,
            self.tv.delete_all,
            self.Help_Window
        ]

        self.buttons = []

        for i in range(len(buttonfuncs)):
            self.buttons.append(ttk.Button(self,text=buttontexts[i],command=buttonfuncs[i]))
            self.buttons[i].bind('<Return>',buttonfuncs[i])

        # CheckBox
        self.cbox_bkmark = ttk.Checkbutton(self,text='Add Bookmarks')
        self.cbox_bkmark.bind('<Return>',lambda e: self.cbox_bkmark.invoke())
        self.cbox_bkmark.invoke()

        # Adding Compile Button
        self.buttons.append(ttk.Button(self,text='Compile',command=lambda: self.call_compile_files()))
        self.buttons[-1].bind('<Return>', lambda e: self.call_compile_files())

        # Drawing self.buttons
        self.buttons[0].grid(row= 0, column=0, padx=(20,10), pady= 10)
        self.buttons[1].grid(row= 0, column=1, padx=10, pady= 10)
        self.buttons[2].grid(row= 0, column=2, padx=10, pady= 10)
        self.buttons[3].grid(row= 0, column=3, padx=10, pady= 10)
        self.buttons[4].grid(row= 0, column=4, padx=10, pady= 10)
        self.buttons[5].grid(row= 0, column=5, padx=10, pady= 10)
        self.buttons[6].grid(row= 0, column=6, padx=10, pady= 10)
        self.cbox_bkmark.grid(row= 0, column=7, padx=10, pady= 10)
        self.buttons[-1].grid(row= 0, column=8, padx=(0,0), pady= 10)
        self.buttons[-2].grid(row= 2, column=0, padx=(20,10), pady= 10)

        self.compiling_label = ttk.Label(self,text='')
        self.compiling_label.grid(row = 2, column=1, columnspan=2, padx=10, pady= 10)

        # Help Screen
        self.hwin = None
        self.bind('<F1>',lambda e: self.Help_Window())


    def call_compile_files(self):
        # function to call compile_files() function from our treeview class
        if self.tv.compiling:

            self.tv.cancel_compiling = True

            for b in self.buttons[:-2]:
                b.configure(state='normal')
            self.buttons[-1].configure(text='Compile')
            self.cbox_bkmark.configure(state='normal')
        
        cbox_state = self.cbox_bkmark.state()
        addbm = False
        for i in cbox_state:
            if i == 'selected':
                addbm = True

        for b in self.buttons[:-2]:
            b.config(state='disabled')
        self.buttons[-1].config(text='Cancel')
        self.cbox_bkmark.config(state='disabled')
        
        self.tv.compile_files(addbm)

        for b in self.buttons[:-2]:
            b.configure(state='normal')
        self.buttons[-1].configure(text='Compile')
        self.cbox_bkmark.configure(state='normal')
  
        

    def Help_Window(self):
        
        if not (self.hwin is None or not self.hwin.winfo_exists()):
            self.hwin.focus()
            return
        self.hwin = tk.Toplevel(self)
        try:
            self.hwin.iconbitmap(self.tempicon)
        except:
            pass
        self.hwin.title('Help Window')

        label = ttk.Label(self.hwin,text = helptext)
        label.pack(padx=20,pady=20)




if __name__ == '__main__':
    window = App()
    window.mainloop()
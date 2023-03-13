
import customtkinter as ctk
from tkinter import PhotoImage
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
> Files cannot be dragged and dropped into the this version of File - PDF combiner (visit github site for info)
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

class App(ctk.CTk):
    def __init__(self):
        ctk.CTk.__init__(self)
        self.title('File - PDF Combiner')
        self.icon = Icon()
        self.tempicon = self.icon.tempicon
        self.iconbitmap(self.tempicon)
        

        # Defining Frames

        sidebar = ctk.CTkFrame(self, width=140, corner_radius=0)

        # Treeview
        self.tv = mtv.modifiedTreeview(self)
        self.tv['columns'] = ('fname','filetype','bookmark','fpath','status')

        # Formatting Columns
        st = True
        self.tv.column("#0", width=0,stretch=0)
        self.tv.column("fname", anchor='w', minwidth=25, stretch=st)
        self.tv.column("filetype", anchor='w', minwidth=25, width = 100, stretch=0)
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
        yscrollbar = ctk.CTkScrollbar(self)
        yscrollbar.configure(command=self.tv.yview)
        self.tv.config(yscrollcommand=yscrollbar.set)

        # Defining Buttons
        buttontexts = [
            "Add Files","Add Folder","Move up","Move down","Edit Bookmark","Delete","Delete All","Help"
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
            self.buttons.append(ctk.CTkButton(sidebar,text=buttontexts[i],command=buttonfuncs[i]))
            self.buttons[i].bind('<Return>',buttonfuncs[i])

        # CheckBox
        self.cbox_bkmark = ctk.CTkCheckBox(sidebar,text='Add Bookmarks')
        self.cbox_bkmark.bind('<Return>',lambda e: self.cbox_bkmark.toggle())
        self.cbox_bkmark.toggle()

        # Adding Compile Button
        self.buttons.append(ctk.CTkButton(sidebar,text='Compile',command=lambda : self.call_compile_files(), 
                                     fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE")))

        self.buttons[-1].bind('<Return>',lambda e: self.call_compile_files())

        # Screen Grid Configuration

        self.rowconfigure(0, weight = 1, minsize = 25)
        self.grid_columnconfigure(1, weight = 1, minsize = 40)

        # Total rows = 13
        noofrows = 15

        # SideBar Frame
        logo_label = ctk.CTkLabel(sidebar, text="File-PDF Combiner", font=ctk.CTkFont(size=20, weight="bold"))
        help_label = ctk.CTkLabel(sidebar, text="For Help Press F1", font=ctk.CTkFont(size=13, weight="normal"))
        addf_label = ctk.CTkLabel(sidebar, text="Add Files", font=ctk.CTkFont(size=15))
        managef_label = ctk.CTkLabel(sidebar, text="Manage Table", font=ctk.CTkFont(size=15))
        compile_label = ctk.CTkLabel(sidebar, text="Compile Files", font=ctk.CTkFont(size=15))

        logo_label.grid(row = 0, column = 0, padx = 20, pady=(20,0))
        help_label.grid(row = 1, column = 0, padx = 20, pady=(0,10))

        # Drawing Buttons
        addf_label.grid(row= 2, column=0, padx=10, pady= (20,0))
        self.buttons[0].grid(row= 3, column=0, padx=10, pady= 10)
        self.buttons[1].grid(row= 4, column=0, padx=10, pady= 10)
        # row = 5 -> resizeable empty space
        managef_label.grid(row= 6, column=0, padx=10, pady= (20,0))
        self.buttons[2].grid(row= 7, column=0, padx=10, pady= 10)
        self.buttons[3].grid(row= 8, column=0, padx=10, pady= 10)
        self.buttons[4].grid(row= 9, column=0, padx=10, pady= 10)
        self.buttons[5].grid(row= 10, column=0, padx=10, pady= 10)
        self.buttons[6].grid(row= 11, column=0, padx=10, pady= 10)
        # buttons[6].grid(row= 9, column=0, padx=(20,10), pady= 10) # no help button
        # row = 12 -> resizeable empty space
        compile_label.grid(row= 13, column=0, padx=10, pady= (20,0))
        self.cbox_bkmark.grid(row =14, column=0, padx=10, pady= 10)
        self.buttons[8].grid(row= 15, column=0, padx=10, pady= 10)
        
        self.compiling_label = ctk.CTkLabel(sidebar, text="\n", font=ctk.CTkFont(size=15))
        self.compiling_label.grid(row = 16, column=0, padx=10, pady= (10,30))
        sidebar.grid_rowconfigure((5,12),weight=1)
        sidebar.grid(row=0, column=0, rowspan=noofrows, sticky='ns')


        # Configure Treeview + Scrollbar Grid

        self.tv.grid(row = 0, column = 1, rowspan=noofrows, sticky='news', padx = (20,0), ipadx=10, pady=20)
        yscrollbar.grid(row = 0,column = 2, rowspan=noofrows, sticky='ns',padx = (0,20), pady = 20)

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
        
        addbm = self.cbox_bkmark.get()

        for b in self.buttons[:-2]:
            b.configure(state='disabled')
        self.buttons[-1].configure(text='Cancel')
        self.cbox_bkmark.configure(state='disabled')
        
        self.tv.compile_files(addbm)
        
        for b in self.buttons[:-2]:
            b.configure(state='normal')
        self.buttons[-1].configure(text='Compile')
        self.cbox_bkmark.configure(state='normal')

        
    def Help_Window(self):
        if not (self.hwin is None or not self.hwin.winfo_exists()):
            self.hwin.focus()
            return
        self.hwin = ctk.CTkToplevel(self)
        self.hwin.title('Help Window')

        try:
            self.hwin.iconbitmap(self.tempicon)
        except:
            pass

        label = ctk.CTkLabel(self.hwin,text = helptext,justify='left')
        label.pack(padx=20,pady=20)
    
    def removetemp(self):
        self.icon.removetemp()



if __name__ == '__main__':
    window = App()
    window.mainloop()
    window.removetemp()
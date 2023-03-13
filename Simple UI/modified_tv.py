from tkinter import ttk   
from tkinter import filedialog as fd
import tkinter as tk
import tkinterdnd2 as tkdnd
import os
import pdfconvertion as pdfc
from functools import partial
import time

valid_img_names = ['*.jpeg', '*.jpg', '*.png', '*.gif']
valid_word_names = ['*.doc', '*.docm', '*.docx', '*.dot', '*.dotm', '*.dotx', '*.odt', '*.rtf', '*.txt', '*.wps', '*.xml', '*.xps']
valid_excel_names = ['*.xlsx', '*.xlsm', '*.xlsb', '*.xltx', '*.xltm', '*.xls', '*.xlt', '*.xml', '*.xlam', '*.xla', '*.xlw', '*.xlr']
valid_ppt_names = ['*.ppt','*.pptx']
valid_pdf_names = ['*.pdf']
valid_all_names = valid_img_names+valid_pdf_names+valid_excel_names+valid_ppt_names+valid_word_names

# defining for Add Files function later used in class modified tree view
filetypes=(
    ('All Files',valid_all_names),
    ('Image Files',valid_img_names),
    ('Wrod Files',valid_word_names),
    ('Excel Files',valid_excel_names),
    ('PowerPoint Files',valid_ppt_names),
    ('PDF Files',valid_pdf_names)
)


def valid_file_tuple(filepath):
    # This replacement [ "\" -> "/" ] is done to keep iid same for drag and drop files
    # and files from scaning folder using os walk
    filepath = filepath.replace("\\","/")
    filename = os.path.basename(filepath)
    bookmark = os.path.splitext(filename)[0]
    extension = os.path.splitext(filename)[1]
    if extension in pdfc.VALID_FILE_EXTENSIONS:
        return (filename,extension,bookmark,filepath,'')

def get_tuple_from_paths(listfpath):
    listf = []

    for i in range(len(listfpath)):
        
        # is path is of a directory scan whole directory for walid files
        if os.path.isdir(listfpath[i]):
            for root, directories, files in os.walk(listfpath[i]):
                for file in files:
                    filepath = os.path.join(root,file)
                    if tmp := valid_file_tuple(filepath): listf.append(tmp)
        
        # in case if path is not of directory (then it is of file) simply
        # adding the file
        else:
            if tmp := valid_file_tuple(listfpath[i]): listf.append(tmp)
    
    return listf

# a bit of this code was taken from following link and modified for treeview
# https://stackoverflow.com/questions/14459993/tkinter-listbox-drag-and-drop-with-python

class modifiedTreeview(ttk.Treeview):

    def __init__(self, master, **kw):

        ttk.Treeview.__init__(self, master, selectmode='extended')

        # drag and drop bind
        self.dnd_bind('<<Drop>>', self.drag_and_drop_files)

        # mouse binds
        self.bind('<Button-1>', self.onClick)
        self.bind('<B1-Motion>', self.moveSelection)
        self.bind('<Double-Button-1>',self.onDoubleClick)
        self.bind('<Control-1>', self.toggleSelection)
        self.bind('<Leave>',  self.onLeave)
        self.bind('<Enter>',  self.onEnter)

        # keyboard binds
        self.bind('<Control-a>',self.selectall)
        self.bind('<Delete>', self.delete_selection)
        self.bind('<Escape>', self.reset_selection)
        self.master.bind('<Up>',self.arrow_up,add = '+')
        self.master.bind('<Down>',self.arrow_down,add = '+')
        self.bind('<Return>',self.edit_bookmark_of_selected)

        self.drop_target_register(tkdnd.DND_FILES, tkdnd.DND_TEXT)
        self.selectionClicked = False
        self.left = False
        self.ctrlClicked = False
        self.scroll_func_lock = False
        self.compiling = False
        self.cancel_compiling = True

# Heading Sorting Mechnism was reffered from ::
# https://stackoverflow.com/questions/1966929/tk-treeview-column-sort

    def heading(self, column, **kwargs):
        if not hasattr(kwargs, 'command'):
            kwargs['command'] = partial(self.sort_heading, column, False)
        return super().heading(column, **kwargs)
    
    def reverse_heading_sort(self,column,reverse):
        return super().heading(column, command=partial(self.sort_heading, column, not reverse))

    def sort_heading(self, column, reverse):

        l = [(self.set(k, column), k) for k in self.get_children('')]
        l.sort(key = lambda s: (s[0].upper(),s[0] ), reverse=reverse)
        for index, (_, k) in enumerate(l):
            self.move(k, '', index)
        self.reverse_heading_sort(column,reverse)

# End of Sorting Functions

# Drag item Sorting Functions:

    def onLeave(self, event):
        # prevents changing selection when dragging
        # already selected items beyond the edge of the listbox
        if self.selectionClicked:
            self.left = True
            return 'break'
        
    def onEnter(self, event):
        #TODO
        self.left = False

    def onClick(self, event):

        self.ctrlClicked = False

        i = self.identify_row(event.y)
        
        if i in self.selection():
            self.selectionClicked = True
            return 'break'

        # if user clicks on empty area, reset selection
        if i == '':
            self.reset_selection()
        

    def toggleSelection(self, event):
        self.ctrlClicked = True

    def moveSelection(self, event):

        if self.ctrlClicked:
            return

        selection = self.selection()

        if not self.selectionClicked or len(selection) == 0:
            return
                
        grid_info = self.grid_info()
        row = grid_info.get('row')
        col = grid_info.get('column')
        colspan = grid_info.get('columnspan')

        temp, temp, temp, height = self.master.grid_bbox(col, row, col+colspan)
        # move selection up or down if mouse is on header or near bottom of treeview

        if event.y <= 20 or event.y >= height-5:

            # Setting loop to allow scroll even when user is not moving mouse

            if self.scroll_func_lock:
                return
            
            self.scroll_func_lock = True

            dir = -1 if event.y<= 20 else 1

            self.scroll_func(dir)

        else:

            self.scroll_func_lock = False

            
            moveto = self.index(self.identify_row(event.y))

            # This loop checks if moveto location is within selection and 
            # stops if it is preventing weird behaviours within selection
            for s in self.selection():
                if moveto == self.index(s):
                    return 'break'
                
            selection = self.selection()

            if moveto < self.index(selection[0]):
                selection = reversed(selection)
            
            for s in selection:
                self.move(s, self.parent(s), moveto)
        
        return

    def scroll_func(self,dir):

        fid1 = self.master.bind('<Leave>',self.release_scroll_func_lock)
        fid2 = self.master.bind('<ButtonRelease>',self.release_scroll_func_lock)
        
        start = time.time()
        
        while self.scroll_func_lock:
            self.master.update()
            # following function decreases delay with time
            delay = 0.2/(time.time() - start+2)
            time.sleep(delay)
        
            if self.scroll_func_lock:
                self.yview_scroll(dir,'units')
     
        self.master.unbind('<Leave>',fid1)
        self.master.unbind('<ButtonRelease>',fid2)

    def release_scroll_func_lock(self,event):
        self.scroll_func_lock = False

# Functions that enable entries Editing Field
# Help was taken from following link:
# https://stackoverflow.com/questions/18562123/how-to-make-ttk-treeviews-rows-editable

    def onDoubleClick(self, event):

        # what row and column was clicked on
        rowid = self.identify_row(event.y)
        columnid = self.identify_column(event.x)

        # if 2nd column was double clicked
        # columnid condition can be removed if all entries
        # are to be made editable
        if rowid != '' and columnid == '#3':
            self.edit_bookmark(rowid,columnid)
            return

    def edit_bookmark_of_selected(self, event=''):
        rowid = self.selection()[0]
        columnid = '#3'
        self.edit_bookmark(rowid,columnid)

    def edit_bookmark(self,rowid,columnid):
        
        self.selection_set(rowid)

        # get column position info
        tvgi = self.grid_info()
        gridx, gridy, temp, temp = self.master.grid_bbox(tvgi.get('column'), tvgi.get('row'))
        gwidth = self.column(columnid,'width')
        x,y,temp,bheight = self.bbox(rowid, columnid)
        epx = x + tvgi.get('padx')[0] + gridx
        epy = y + gridy # + tvgi.get('pady')[0]
        
        # place Entry popup properly         
        text = self.item(rowid,'values')[2]
        editbm = EntryBox(self.master,self,rowid)
        editbm.place(x=epx, y=epy, width=gwidth, height=bheight+2)
        editbm.insert(0,text)
        editbm.focus()


# Other Small Functions for basic functionality

    def selectall(self, event=''):
        self.selection_add(self.get_children())

    def delete_selection(self, event=''):
        selected_items = self.selection()
        
        for selected_item in selected_items[::-1]:
            self.delete(selected_item)

    def delete_all(self,event=''):
        self.selectall()
        self.delete_selection()
        
    def reset_selection(self, event=''):
        if len(self.selection()) > 0:
            self.selection_remove(self.selection())

    def move_selection_up(self, event=''):
        selection = self.selection()
        moveto = self.index(selection[0]) - 1
        selection = reversed(selection)
        for s in selection:
            self.move(s, self.parent(s), moveto)

    def move_selection_down(self, event=''):
        selection = self.selection()
        moveto = self.index(selection[-1]) + 1
        for s in selection:
            self.move(s, self.parent(s), moveto)

    def arrow_up(self, event=''):
        if self.focus() == '':
            self.selection_set(self.get_children()[-1])
            self.focus(self.get_children()[-1])
        
    def arrow_down(self, event=''):
        if self.focus() == '':
            self.selection_set(self.get_children()[0])
            self.focus(self.get_children()[0])

## Functions to add Entries into Treeview

# Drag and Drop Item into treeview functiom:
    def drag_and_drop_files(self, event):
        if event.data:
            if event.widget == self:
                listfpath = self.tk.splitlist(event.data)
        
        sorted(listfpath, key=lambda s: (s[0].upper(),s))
        listf = get_tuple_from_paths(listfpath)

        for i in range(len(listf)):
            try:
                self.insert(parent='', index='end', iid = listf[i][3], values = listf[i])
            except:
                continue

    def add_files_from_explorer(self, event=''):
        filepaths = fd.askopenfilenames(filetypes=filetypes)
        for filepath in filepaths:
            filename = os.path.basename(filepath)
            bookmark = os.path.splitext(filename)[0]
            extension = os.path.splitext(filename)[1]
            try:
                self.insert(parent='', index='end', iid = filepath, values = (filename,extension,bookmark,filepath,''))
            except:
                continue

    def add_folder_from_explorer(self, event=''):
        folderpath = fd.askdirectory()
        for root, directories, files in os.walk(folderpath):
            for file in files:
                filepath = os.path.join(root,file)
                if tmp := valid_file_tuple(filepath):
                    try:
                        self.insert(parent='', index='end', iid = tmp[3], values = tmp)
                    except:
                        continue


# Custom function that compiles files to PDF for this APP

    def compile_files(self,add_bookmark):

        # Lock for avoiding multiple calls 
        if self.compiling: return
        self.compiling = True

        self.cancel_compiling = False
        allfiles = self.get_children()
        total = len(allfiles)
        current = 0
        failed = 0

        if  total == 0:
            self.compiling = False
            return
        
        merger = pdfc.pdfmerging()

        # resetting status
        if self.item(allfiles[0])['values'][3] != '':
            for file in allfiles:
                self.set(file,'status','')

        
        for file in allfiles:
            current += 1

            status = 'Merging...'
            self.set(file,'status',status)
            
            labelstr = 'Status: Compiling ' + str(current) + '/' + str(total) + '. Failed: '+ str(failed) 
            self.master.compiling_label.config(text=labelstr)
            
            self.master.update()

            if merger.addfile(file,self.item(file)['values'][2],add_bookmark):
                status = 'Merged'
            else:
                failed += 1
                status = 'Failed'
            
            self.set(file,'status',status)
            self.master.update()
            if self.cancel_compiling:
                break
        
        if len(merger.mergingExceptions) != 0:
            self.Merging_ExceptionWindow(merger.mergingExceptions)
            self.bell()
            self.master.update()
            time.sleep(1)
    
        filedest = fd.asksaveasfilename(filetypes=(('PDF Files',valid_pdf_names),),defaultextension='.pdf')
        compile_status = merger.close_merger_task(filedest)

        if merger.closingException != '':
            self.SavingFile_ExceptionWindow(merger.closingException)
            self.bell()
            self.master.update()
            time.sleep(1)

        labelstr = 'Compiled ' + str(current-failed) + '/' + str(total) + '. Failed: '+ str(failed) 
        self.master.compiling_label.configure(text=labelstr)

        if compile_status == False:
            pass

        self.cancel_compiling = False
        self.compiling = False
    
    def Merging_ExceptionWindow(self,exception_list):
        exception_win = tk.Toplevel(self.master)
        exception_win.title('Errors While Merging')
        try:
            exception_win.iconbitmap(self.master.tempicon)
        except:
            pass
        exception_win.rowconfigure(1, weight=1, minsize=25)
        exception_win.rowconfigure(0, weight=1, minsize=25)
        exception_win.rowconfigure(1, weight=1, minsize=25)

        label = ttk.Label(exception_win,text = 'Following exceptions Occured')
        label.pack(padx=20,pady=(20,0))
        
        textbox = tk.Text(exception_win, )
        yScrollbar = ttk.Scrollbar(exception_win, orient='vertical')
        yScrollbar.config(command=textbox.yview)
        textbox.config(yscrollcommand=yScrollbar.set)
        
        for e in exception_list:
            textstr = 'For file: ' + e[0] + '\n Error : '+ str(e[1]) +  '\n\n\n'
            textbox.insert('end', textstr) 
        yScrollbar.pack(side='right',fill='y')
        textbox.pack(padx=(20,0),pady=(5,20), expand=1)

        textbox.configure(state='disabled')

    def SavingFile_ExceptionWindow(self,exception):
        exception_win = tk.Toplevel(self.master)
        exception_win.title('Errors While Saving')
        try:
            exception_win.iconbitmap(self.master.tempicon)
        except:
            pass
        exception_win.rowconfigure(1, weight=1, minsize=25)
        exception_win.rowconfigure(0, weight=1, minsize=25)
        exception_win.rowconfigure(1, weight=1, minsize=25)

        label = ttk.Label(exception_win,text = 'Following exceptions Occured')
        label.pack(padx=20,pady=(20,0))
        
        textbox = tk.Text(exception_win, )
        yScrollbar = ttk.Scrollbar(exception_win, orient='vertical')
        yScrollbar.config(command=textbox.yview)
        textbox.config(yscrollcommand=yScrollbar.set)
        tstring = 'Error : ' + str(exception)
        textbox.insert('end',tstring) 
        yScrollbar.pack(side='right',fill='y')
        textbox.pack(padx=(20,0),pady=(5,20), expand=1)

        textbox.configure(state='disabled')

# Entry box widget for editting entries

class EntryBox(ttk.Entry):
    def __init__(self,master,tree,iid):
        ttk.Entry.__init__(self,master)
        self.tree = tree
        self.iid = iid
        self.fid = []
        self.fid.append(self.master.bind('<1>',self.destroy_entrybox))
        self.fid.append(self.master.bind('<Button>',self.checkwidget))
        self.fid.append(self.master.bind('<Return>',self.return_value))
        self.fid.append(self.master.bind('<Escape>',lambda e: self.destroy_wd()))
        self.fid.append(self.master.bind('<MouseWheel>',lambda e: self.destroy_wd()))
        self.fid.append(self.bind('<Button>',self.checkwidget))
        self.tree.bind('<Button-1>',self.checkwidget)

    def destroy_wd(self):
        binds = ('<1>','<ButtonPress>','<Return>','<Escape>','<MouseWheel>')
        for i in range(len(binds)):
            self.master.unbind(binds[i],self.fid[i])
        self.tree.bind('<Button-1>',self.tree.onClick)
        self.destroy()

    # if any event occurs outside destroy Entry Box
    def destroy_entrybox(self,event):
        if event.widget != self:
            self.destroy_wd()
    
    # only return value if pressed Enter
    def return_value(self,event):

        if event.widget != self:
            self.destroy_wd()
            return
        
        self.tree.set(self.iid,'#3',self.get())
        self.destroy()

    def checkwidget(self,event):
        if event.widget != self:
            self.destroy_wd()


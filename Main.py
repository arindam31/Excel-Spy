#FU3_5

import wx
import re
import os.path,sys
from MyExcel_Search import *
from Writeit import *


class Example(wx.Frame):
  
    def __init__(self, parent, title):
        super(Example, self).__init__(parent, title=title, 
            size=(650, 650))
            
        self.InitUI()
        self.Centre()
        
#*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.
#--------------------------------------------------------------------------------
        
        favicon = wx.Icon('desktop.ico', wx.BITMAP_TYPE_ICO)
        wx.Frame.SetIcon(self, favicon)
        self.Show()     
        
    def InitUI(self):
    
        panel = wx.Panel(self)

        font = wx.SystemSettings_GetFont(wx.SYS_SYSTEM_FONT)
        font.SetPointSize(9)

        font1 = wx.SystemSettings_GetFont(wx.SYS_SYSTEM_FONT)
        font1.SetPointSize(10)

        vbox = wx.BoxSizer(wx.VERTICAL)

#*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.
#--------------------------------------------------------------------------------


#My variables


        self.finalfileholder = [] #Holds the final list of files
        self.finalpathholder=[] ##Holds the final list of paths
        temppath=os.path.abspath(os.path.dirname(sys.argv[0]))
        self.dirName2 = temppath  #If user does'nt decide a loc himself, this will be the default location.
        self.file_fullname_and_path = [] #This will hold full file and path name , we need this to send to our searching module
        self.file_count = 0
        #self.copy_of_filename=[]
        #temp_var=[]
        self.result=[]
        self.result_New=[]
        self.list_of_cellnames=[] #Cell List
        self.sheet_name=[] #Sheet List
        self.list_filenames = []
#*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.
#--------------------------------------------------------------------------------
        
#First Horizontal Box
       
        hbox1 = wx.BoxSizer(wx.HORIZONTAL)

  #Static Text for 'Search'
        st1 = wx.StaticText(panel, label='Search')
        st1.SetFont(font)
        hbox1.Add(st1, flag=wx.RIGHT|wx.TOP, border=8)

  #Search Box
        
        self.Search_Text = wx.TextCtrl(panel,-1,"Enter search string",size=wx.Size(-1,30))
        hbox1.Add(self.Search_Text, proportion=1)        
        self.Search_Text.SetFont(font1)

  #Search Button
        
        btn_Process = wx.Button(panel,-1,label='Search')
        self.Bind(wx.EVT_BUTTON, self.Process, btn_Process)
        hbox1.Add(btn_Process,flag=wx.RIGHT|wx.LEFT|wx.TOP, border=5)
        vbox.Add(hbox1, flag=wx.EXPAND|wx.LEFT|wx.RIGHT|wx.TOP, border=10)
        vbox.Add((-1, 10))

#*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.
#--------------------------------------------------------------------------------
        
#Second Horizontal Box
        
        hbox2 = wx.BoxSizer(wx.HORIZONTAL)
        st2 = wx.StaticText(panel, label='File Repository')
        st2.SetFont(font)
        hbox2.Add(st2)
        vbox.Add(hbox2, flag=wx.LEFT | wx.TOP, border=10)

        vbox.Add((-1, 10))

#*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.
#--------------------------------------------------------------------------------        
#Third Horizontal Box
        
 #File Repository
        
        self.Files_List = wx.ListCtrl(panel,-1,style=wx.LC_REPORT|wx.BORDER_SUNKEN|wx.LC_SINGLE_SEL)
        self.Files_List.InsertColumn(0, 'File Name')
        self.Files_List.InsertColumn(1, 'Path')
        self.Files_List.SetColumnWidth(0, 200)
        self.Files_List.SetColumnWidth(1, 300)        

        bt_Add = wx.Button(panel,-1,label='Add')
        self.bt_Remove = wx.Button(panel,-1,label='Remove')
        self.bt_Clear = wx.Button(panel,-1,label='Clear')
        
        hbox3 = wx.BoxSizer(wx.HORIZONTAL)
        vbox1 = wx.BoxSizer(wx.VERTICAL)
        
        vbox1.Add(bt_Add,flag=wx.RIGHT|wx.LEFT, border=5)
        vbox1.Add(self.bt_Remove,flag=wx.RIGHT|wx.LEFT|wx.TOP, border=5)
        vbox1.Add(self.bt_Clear,flag=wx.RIGHT|wx.LEFT|wx.TOP, border=5)
        
        hbox3.Add(self.Files_List, proportion=1, flag=wx.EXPAND)
        hbox3.Add(vbox1,flag=wx.RIGHT|wx.LEFT, border=10)
        
        vbox.Add(hbox3, proportion=1, flag=wx.LEFT|wx.RIGHT|wx.EXPAND, 
            border=10)
        self.Files_List.SetFont(font) 
        vbox.Add((-1, 25))

#*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.
#--------------------------------------------------------------------------------
        
#Fourth Horizontal Box

        hbox4 = wx.BoxSizer(wx.HORIZONTAL)
        self.cb1 = wx.CheckBox(panel, label='Ignore Case')
        self.cb1.SetFont(font)
        hbox4.Add(self.cb1)
        self.cb2 = wx.CheckBox(panel, label='Regular Expression')
        self.cb2.SetFont(font)
        hbox4.Add(self.cb2, flag=wx.LEFT, border=10)
        self.cb3 = wx.CheckBox(panel, label='Exact match')
        self.cb3.SetFont(font)
        hbox4.Add(self.cb3, flag=wx.LEFT, border=10)
        self.cb4 = wx.CheckBox(panel, label='Exclude')
        self.cb4.SetFont(font)
        hbox4.Add(self.cb4, flag=wx.LEFT, border=10)
        vbox.Add(hbox4, flag=wx.LEFT, border=10)

        vbox.Add((-1, 25))

#*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.
#--------------------------------------------------------------------------------
        
#Fifth Horizontal Box
        
        hbox5 = wx.BoxSizer(wx.HORIZONTAL)
        self.bt_Export = wx.Button(panel, label='EXPORT', size=(70, 30))
        self.bt_Edit = wx.Button(panel, label='EDIT', size=(70, 30))
        hbox5.Add(self.bt_Edit, flag=wx.RIGHT|wx.BOTTOM, border=5)
        hbox5.Add(self.bt_Export)
        bt_Reset = wx.Button(panel, label='RESET', size=(70, 30))
        hbox5.Add(bt_Reset, flag=wx.LEFT|wx.BOTTOM, border=5)
        vbox.Add(hbox5, flag=wx.ALIGN_RIGHT|wx.RIGHT|wx.BOTTOM, border=10)
        

        vbox.Add((-1, 25))

#*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.
#--------------------------------------------------------------------------------
        #Sixth Horizontal Box
 
        #List Control for Result display
        hbox6 = wx.BoxSizer(wx.HORIZONTAL)
        self.lc = wx.ListCtrl(panel,-1,style=wx.LC_REPORT|wx.BORDER_SUNKEN|wx.LC_SINGLE_SEL)
        self.lc.InsertColumn(0, 'Result')
        self.lc.InsertColumn(1, 'Cell')
        self.lc.InsertColumn(2, 'Sheet')  
        self.lc.InsertColumn(3, 'File')   
        self.lc.SetColumnWidth(0, 200)
        self.lc.SetColumnWidth(1, 50)
        self.lc.SetColumnWidth(2, 60)
        self.lc.SetColumnWidth(3, 200)
        hbox6.Add(self.lc, proportion=1, flag=wx.EXPAND)
        vbox.Add(hbox6, proportion=1, flag=wx.LEFT|wx.RIGHT|wx.EXPAND, 
            border=10)
        vbox.Add((-1, 5))

#*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.*.
#--------------------------------------------------------------------------------
        #Seventh Horizontal Box


        hbox7 = wx.BoxSizer(wx.HORIZONTAL)
        self.statusbar = self.CreateStatusBar()
        self.statusbar.SetStatusText('Ready')
        hbox7.Add(self.statusbar, proportion=1, flag=wx.EXPAND)
        vbox.Add(hbox7,flag=wx.TOP|wx.EXPAND,border=5)


        panel.SetSizer(vbox)
       

#Menu section
        menubar=wx.MenuBar()
        pehla=wx.Menu()
        doosra=wx.Menu()
        teesra =wx.Menu()
        option_menu=wx.Menu()
        info=wx.Menu()

#Menu Items
        
        item1_1=pehla.Append(wx.ID_OPEN,"&Add\tAlt-A","This is add files") #Sub-Items of First menu pull down list
        item1_2=pehla.Append(wx.ID_EXIT,"&Quit\tAlt-Q","This will exit app") #The last comment will show on status bar when mouse is on that option
        item3_2=teesra.Append(wx.ID_ABOUT,"A&bout\tAlt-B","About Section")

        
        menu_1=menubar.Append(pehla,'&File')    #Naming of Menu items
        menu_2=menubar.Append(doosra,'&Edit')
        menu_3=menubar.Append(teesra,'&Info')
        item2_1=option_menu.Append(wx.ID_ANY,'Export File Location')
        doosra.AppendMenu(wx.ID_ANY,"&Options\tAlt-O",option_menu)
        self.SetMenuBar(menubar)

#Events
        self.Bind(wx.EVT_MENU, self.OnFileExit,item1_2)
        self.Bind(wx.EVT_MENU, self.OnFileOpen,item1_1)
        self.Bind(wx.EVT_BUTTON, self.OnReset, bt_Reset)
        self.Bind(wx.EVT_BUTTON, self.OnExport, self.bt_Export)
        self.Bind(wx.EVT_MENU, self.OnOptions, item2_1)
        self.Bind(wx.EVT_BUTTON, self.OnFileOpen,bt_Add)
        self.Bind(wx.EVT_BUTTON, self.OnRemove,self.bt_Remove)
        self.Bind(wx.EVT_BUTTON, self.OnClear,self.bt_Clear)
        self.Bind(wx.EVT_MENU, self.OnAbout,item3_2)
        self.Bind(wx.EVT_BUTTON, self.OnEdit,self.bt_Edit)
        
#Set Launch state of buttons        
        self.bt_Export.Disable()
        self.bt_Clear.Disable()
        self.bt_Remove.Disable()

#^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^Function Definations^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
#::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::

        
    def OnFileOpen(self, event):
        """ File|Open event - Open dialog box. """
        self.dirname = '' #This is to set current working directory is the our default folder which will open when fileDialog will be called . 
        dlg = wx.FileDialog(self, "Choose single or multiple files", self.dirname, "", "*.xls",wx.FD_MULTIPLE)
        self.tempfiles_holder=[]
        self.temppath_holder=[]
        
        if (dlg.ShowModal() == wx.ID_OK):
            
            self.fileName = dlg.GetFilenames() #File name list
            #self.dirNamewithpath = dlg.GetPaths() #Directory name with file name 
            self.dirName = [os.path.dirname(i) for i in dlg.GetPaths()]
            
            self.bt_Clear.Enable()
            self.bt_Remove.Enable()

            for i in range(len(self.fileName)):
                if self.fileName[i] not in self.finalfileholder:
                    self.tempfiles_holder.append(self.fileName[i])
                    self.temppath_holder.append(self.dirName[i])
                else:
                    wx.MessageBox('File already exists', 'Info',wx.OK | wx.ICON_INFORMATION)
                    return
            
            for k in range(len(self.tempfiles_holder)):
                           self.finalfileholder.append(self.tempfiles_holder[k])
                           self.finalpathholder.append(self.temppath_holder[k])
                          
                           
            no_files=self.Files_List.GetItemCount()#This will be the index . self.Files_List is listctrl for file repository
            for  j in range(len(self.tempfiles_holder)):
                self.Files_List.InsertStringItem(no_files,self.tempfiles_holder[j])
                
                self.Files_List.SetStringItem(no_files,1,self.temppath_holder[j])
                no_files=no_files+1

                

            for k in range(len(self.tempfiles_holder)):
                
                           self.file_fullname_and_path.append(self.temppath_holder[k]+'\\'+self.tempfiles_holder[k])
            

                               

#^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-
            
            
    def OnFileExit(self, event):
        """ File|Exit event """
        self.Close()
    def Process(self,event):
        if len(self.file_fullname_and_path)==0:
            wx.MessageBox('Add some files first', 'Info',wx.OK | wx.ICON_INFORMATION)
            return
        else:
            pass
            
        
        self.result=[]
        self.result_New=[]
        self.list_of_cellnames=[] #Cell List
        self.sheet_name=[] #Sheet List
        self.list_filenames = []
        
        Pattern=self.Search_Text.GetValue()
        rawpattern=Pattern #Will be used when we need it in a raw format
        print rawpattern

#^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-
#^-^-^-^-^-^-^-^-^-^-For empty search space check^-^-^-^-^-^-^-^-
        
        if len(Pattern)==0:
            return
        else:
            Patternlist=Pattern.split(",")#This is to seperate strings by comma .
            
        Patternlist = [pat.strip() for pat in Patternlist]


#^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-
#^-^-^-^-^-^-^-^-^-^-For enabling search backlash and more^-^-^-^-^-^-^-^-        

        RE_cb = self.cb2.GetValue()
        
        if RE_cb == False:
            for pat in Patternlist:
                print pat
                try:
                    pat = pat.replace("\\",r"[\\]").replace("^","[\^]").replace(".","[.]").replace("?","[?]").replace("+","[+]") #Will replace special symbols that need escaping and bracketing
                except:pass
                    
            for pat in Pattern:
                try:
                    re.compile(Pattern) #This is for '?' and '+' who return error.
                except:
                    Pattern="["+Pattern+"]" #Convert the symbols to strings inside bracket

        

#^-^-^-^-^-^-^-^-^-^-Just checking the regular expression flag status^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-^-
                                      

        
        elif RE_cb == True:
            self.cb1.SetValue(False) #If RE is on , no need of Case Search to be on .
            self.cb3.SetValue(False) #If RE is on , no need of Exact Match to be on .
            Pattern = rawpattern #Did this as we dont want the backlashes escaped
            for pat in Patternlist:
                print pat
                if pat[-1] == '\\':
                    wx.MessageBox("Regular Expression Pattern can't end with '\\'", 'WARNING',wx.OK |wx.ICON_INFORMATION)
                    return
                try:
                    re.compile(Pattern)
                except:
                    wx.MessageBox("Invalid Regular Expression Pattern", 'Info',wx.OK | wx.ICON_EXCLAMATION)
                    self.statusbar.SetStatusText("Ready")
                    self.lc.DeleteAllItems()
                    return            
                
            else:pass

        #Results taken only after determining their final state (after Regex flag status is checked)

        EX_cb = self.cb3.GetValue()
        case_cb = self.cb1.GetValue()
        exclude_cb = self.cb4.GetValue()
      
         #We call our xl function in MyExcel_Search
        
        self.result,self.list_of_cellnames,self.sheet_name,self.list_filenames = xl(self.file_fullname_and_path,Patternlist,case_cb,RE_cb,EX_cb,exclude_cb) #Collect cellname and search results


        self.result_New = []
        

        if self.result:#If the list is not empty,then we need the export button deactivated. 
            
            self.bt_Export.Enable()
            for j in range(len(self.result)):
                try:
                    self.result_New.append(str(self.result[j])) #We need only in Strings format to pass to list control
                except UnicodeEncodeError:
                    pass
                
            self.loadResult() #Called it to load the list with results
        else:
            wx.MessageBox('Sorry, NO matches found', 'Info',wx.OK | wx.ICON_INFORMATION)  #If list has come back with no results
            self.lc.DeleteAllItems() #Need to clear result display area
            self.statusbar.SetStatusText("Ready")

        
        

    def loadResult(self):
        
        # clear the Result listctrl
        self.lc.DeleteAllItems()
        self.index=0  #First declaration
        # load each data row
        for i in range(len(self.result_New)):

            self.lc.InsertStringItem(self.index,self.result_New[i])#max rows value and starting point , here resource management by dunamic allocation
            temp1=self.list_of_cellnames[i] #We pass cell no one at a time 
            temp2=self.sheet_name[i]
            temp3=self.list_filenames[i]
            self.lc.SetStringItem(self.index,1,temp1)
            self.lc.SetStringItem(self.index,2,temp2)
            self.lc.SetStringItem(self.index,3,temp3)
            self.index+=1
        self.statusbar.SetStatusText(str(len(self.result))+" matches found") #Display results on staus bar

    def OnReset(self,event):
        #This will clear all display sections.
        self.lc.DeleteAllItems()
        self.Search_Text.Clear()
        self.Files_List.DeleteAllItems() #File Respository 
        #We are clearing all list so all previous data is flashed.
        self.result=[]
        self.result_New=[]
        self.finalpathholder=[]
        self.finalfileholder=[]
        self.list_of_cellnames=[] #Cell List
        self.sheet_name=[] #Sheet List
        #Clear status bar and deactivate Export button
        self.bt_Export.Disable()
        self.bt_Clear.Disable()
        self.bt_Remove.Disable()

        self.statusbar.SetStatusText("Ready")
        

    def OnExport(self,event):
        self.Savefile = wx.TextEntryDialog(self, 'Enter File name','Save Results') #Save file name user entry
        self.Savefile.SetValue("Result")
        if self.Savefile.ShowModal() == wx.ID_OK:
            ResultFileName = self.Savefile.GetValue()
            self.Savefile.Destroy()

            #Export function in write.py called
            #Sending Filename , results , cell no , sheet names , file name , and directory to save the file.
            Export(ResultFileName,self.result,self.list_of_cellnames,self.sheet_name,self.list_filenames,self.dirName2)

    def OnOptions(self,event):
        
        dlg2 = wx.DirDialog(self, "Choose a Directory",style=wx.DD_DEFAULT_STYLE | wx.DD_NEW_DIR_BUTTON)
        
        if (dlg2.ShowModal() == wx.ID_OK):
            self.dirName2 = dlg2.GetPath()


    def OnRemove(self, event):
        index_focus = self.Files_List.GetFocusedItem()
        if not index_focus == -1: #We are doing this , because we need to only delete from list when something was actually highlighted
            self.Files_List.DeleteItem(index_focus)
            self.finalfileholder.remove(self.finalfileholder[index_focus])
            self.finalpathholder.remove(self.finalpathholder[index_focus])
            self.file_fullname_and_path.remove(self.file_fullname_and_path[index_focus])
            

        if not len(self.finalfileholder):
            #self.bt_Export.Disable() - Disabling for now
            self.bt_Remove.Disable()
            self.bt_Clear.Disable()
            
              
        

    def OnClear(self,event):
        self.Files_List.DeleteAllItems()
        self.bt_Clear.Disable()
        self.bt_Remove.Disable()
        self.finalfileholder=[]
        self.file_fullname_and_path = []

    def OnEdit(self,event):
        index_focus_edit = self.lc.GetFocusedItem()
        if index_focus_edit == -1:
            wx.MessageBox('Please select a search result', 'Info',wx.OK | wx.ICON_INFORMATION)
            return
        else:
            file_edit=self.list_filenames[index_focus_edit]
            from  os import startfile
            startfile(file_edit)
   

    def OnAbout(self,event):

        description = """\tXcel Search is a small search application that can search within excel files with .xls
             format only. I designed this to solve a common problem which sometimes people face ,
             which is merely searching for one thing in a large number of files .I have tried to solve
             this problem,since I faced it when i needed but I did'nt had any choice.
             I hope this helps someone ."""

        licence = """This is a free app.I have not acquired a copyright or licence . Frankly I don't even know how to.
            People who want to take inspiration from this , my best wishes.""" 


        info = wx.AboutDialogInfo()

#        info.SetIcon(wx.Icon('icons/hunter.png', wx.BITMAP_TYPE_PNG))
        info.SetName('Xcel Search')
        info.SetVersion('1.0.2')
        info.SetDescription(description)
        info.SetCopyright('(C) 2012 Arindam Roychowdhury')
       # info.SetWebSite('arindam31@yahoo.co.in')
        info.SetLicence(licence)
        info.AddDeveloper('Arindam Roychowdhury')
        

        wx.AboutBox(info)
    

        

#^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
#::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::


if __name__ == '__main__':
  
    app = wx.App()
    Example(None, title='Excel Spy')
    app.MainLoop()

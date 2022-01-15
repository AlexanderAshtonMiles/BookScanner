'''
Book Scanner v0.1
Alexander Miles, 2021
www.AlexanderMiles.com

This software comes with no warranty, use at your own risk, etc. 
GNU General Public License v3.0
'''

import wx
from glob import glob
from pyzbar.pyzbar import decode
from PIL import Image
import json
import urllib
from isbnlib import * #meta is_isbn13 is_isbn10
import csv
import os
import shutil

def scale_bitmap(bitmap, width, height):
    #image = wx.ImageFromBitmap(bitmap)
    image = bitmap.ConvertToImage()
    H = image.GetHeight()
    W = image.GetWidth()
    #print("Height found was " + str(H))
    #print("Width found was " + str(H))
    if H > W:
        finalH = height
        finalW = int( W/H * height )
    else:
        finalW = width
        finalH = int( H/W * width )
    #print("Aspect ratio is " + str(finalH/finalW))
    image = image.Scale(finalW, finalH, wx.IMAGE_QUALITY_HIGH)
    result = wx.Bitmap(image)
    return result

class MyApp(wx.App):
    def __init__(self):
        super().__init__(clearSigInt=True)

        # Init frame
        self.OnInit()
        
    def OnInit(self):
        frame = MyFrame(None, "Book Recognizer GUI, v0.1", (400,400), (1300,500) )
        frame.Show()
        frame.Fit()
        return True
        
    def OnClose(self, event):
        return

      
class MyFrame(wx.Frame):
    def __init__(self, parent, title, pos, size):
        super().__init__(parent=parent, title=title, pos=pos, size=size)
        self.OnInit()
        
    def OnInit(self):
        panel = MyPanel(parent=self)
        global statusbar
        statusbar = self.CreateStatusBar(1)
        statusbar.SetStatusText('Pick an input folder to scan')
        
class MyPanel(wx.Panel):
    def __init__(self,parent):
        super().__init__(parent=parent)
        self._dont_show = False
        self.PhotoMaxSize = 500
        self.CSVPath = "unset"

        # Generating elements
        
        self.staticbitmap = wx.StaticBitmap(self)
        
        title = wx.StaticText(parent=self, id=wx.ID_ANY, label="Book Recognizer GUI")

        fnameList = ['File names'] 
        self.lst = wx.ListBox(self, size = (100,-1), choices = fnameList, style = wx.LB_SINGLE)
        
        csvButton = wx.Button(parent =self, id = wx.ID_ANY, label="Choose CSV File for Output")
        loadButton = wx.Button(parent =self, id = wx.ID_ANY, label="Load Directory")
        cancelButton = wx.Button(parent =self, id = wx.ID_ANY, label="Cancel")
        skipButton = wx.Button(parent =self, id = wx.ID_ANY, label="Skip file")
        
        # Now for the fields we actually need
        
        labelTitle = wx.StaticText(self, wx.ID_ANY, "Title")
        labelISBN = wx.StaticText(self, wx.ID_ANY, "ISBN-13")
        labelAuthor = wx.StaticText(self, wx.ID_ANY, "Author")
        labelYear = wx.StaticText(self, wx.ID_ANY, "Year")
        labelPublisher = wx.StaticText(self, wx.ID_ANY, "Publisher")
        labelSubject = wx.StaticText(self, wx.ID_ANY, "Subject Tags")
        
        labelSubjectButtons = wx.StaticText(self, wx.ID_ANY, "Common tags")
        subject1Button = wx.Button(parent =self, id = wx.ID_ANY, label="Fiction")
        subject2Button = wx.Button(parent =self, id = wx.ID_ANY, label="Nonfiction")
        subject3Button = wx.Button(parent =self, id = wx.ID_ANY, label="Graphic Novel")

        
        self.inputTitle = wx.TextCtrl(self, wx.ID_ANY|wx.EXPAND, '')
        self.inputISBN = wx.TextCtrl(self, wx.ID_ANY, '')
        self.inputAuthor = wx.TextCtrl(self, wx.ID_ANY, '')
        self.inputYear = wx.TextCtrl(self, wx.ID_ANY, '')
        self.inputPublisher = wx.TextCtrl(self, wx.ID_ANY, '')
        self.inputSubject = wx.TextCtrl(self, -1, style = wx.EXPAND|wx.TE_MULTILINE|wx.TE_WORDWRAP)
        
        barcodeButton = wx.Button(parent =self, id = wx.ID_ANY, label="Scan Barcode")
        metaButton = wx.Button(parent =self, id = wx.ID_ANY, label="Retrieve Metadata")
        writeRowButton = wx.Button(parent =self, id = wx.ID_ANY, label="Write Row to CSV")
        
        labelMetaSource = wx.StaticText(self, wx.ID_ANY, "Metadata Source")
        self.inputMetaSource = wx.Choice(self, choices=['Google Books','Wikipedia','OpenLibrary'])

        #labelGrabSubj = wx.StaticText(self, wx.ID_ANY, "Retrieve Subject Tags? (OpenLibrary)")
        self.inputGrabSubj = wx.CheckBox(parent = self, label="Retrieve Subject Tags? (OpenLibrary)")
        self.inputGrabOnClick = wx.CheckBox(parent = self, label="Scan and Retrieve on Click?")
        self.inputGrabOnClick.SetValue(True)
        
        # Making sizers to place elements

        # Top level sizers
        evenBiggerSizer = wx.BoxSizer(wx.HORIZONTAL)
        mainSizer = wx.BoxSizer(wx.VERTICAL)
        rightSideSizer = wx.BoxSizer(wx.VERTICAL)
        imgSizer = wx.BoxSizer(wx.HORIZONTAL)
        
        # 2nd level sizers
        titleSizer = wx.BoxSizer(wx.HORIZONTAL)
        inputOneSizer = wx.BoxSizer(wx.HORIZONTAL)
        inputTwoSizer = wx.BoxSizer(wx.HORIZONTAL)
        inputThreeSizer = wx.BoxSizer(wx.HORIZONTAL)
        inputFourSizer = wx.BoxSizer(wx.HORIZONTAL)
        lstSizer = wx.BoxSizer(wx.HORIZONTAL)
        buttonSizer = wx.BoxSizer(wx.HORIZONTAL)
        
        bookTitleSizer = wx.BoxSizer(wx.HORIZONTAL)
        ISBNSizer = wx.BoxSizer(wx.HORIZONTAL)
        AuthorSizer = wx.BoxSizer(wx.HORIZONTAL)
        YearSizer = wx.BoxSizer(wx.HORIZONTAL)
        PublisherSizer = wx.BoxSizer(wx.HORIZONTAL)
        SubjectSizer = wx.BoxSizer(wx.HORIZONTAL)
        metaSizer = wx.BoxSizer(wx.HORIZONTAL)
        
        # 3rd level sizers
        SubjectButtonsSizer = wx.BoxSizer(wx.VERTICAL)
        SubjectButtonsSizer.Add(labelSubject, proportion=0, flag=wx.ALL, border=5)
        SubjectButtonsSizer.Add(labelSubjectButtons, proportion=0, flag=wx.ALL, border=5)
        SubjectButtonsSizer.Add(subject1Button, proportion=1, flag=wx.ALL, border=5)
        SubjectButtonsSizer.Add(subject2Button, proportion=1, flag=wx.ALL, border=5)
        SubjectButtonsSizer.Add(subject3Button, proportion=1, flag=wx.ALL, border=5)
        
        imgSizer.Add(window=self.staticbitmap, flag=wx.ALL, border=5)
        titleSizer.Add(window=title, proportion=0, flag=wx.ALL, border=5)

        lstSizer.Add(window=self.lst, proportion=1, flag=wx.ALL|wx.EXPAND, border=5)
        
        buttonSizer.Add(window=loadButton, proportion=0, flag=wx.ALL, border=5)
        buttonSizer.Add(window=cancelButton, proportion=0, flag=wx.ALL, border=5)

        mainSizer.Add(titleSizer, 0, wx.CENTER|wx.ALIGN_TOP|wx.ALL, 5)
        mainSizer.Add(csvButton, 1, wx.CENTER|wx.ALIGN_TOP|wx.ALL, 5)
        mainSizer.Add(lstSizer, 0, wx.ALIGN_LEFT|wx.ALL|wx.EXPAND, 5)
        mainSizer.Add(buttonSizer, 0, wx.CENTER|wx.ALL, 5)
        mainSizer.Add(skipButton, 1, wx.CENTER|wx.ALL|wx.EXPAND, 5)
        
        bookTitleSizer.Add(labelTitle, 0, wx.ALIGN_LEFT|wx.ALL, 5)
        bookTitleSizer.Add(self.inputTitle, 1, wx.ALIGN_LEFT|wx.EXPAND|wx.ALL, 5)
        ISBNSizer.Add(labelISBN, 0, wx.CENTER|wx.ALL, 5)
        ISBNSizer.Add(self.inputISBN, 1, wx.CENTER|wx.ALL|wx.EXPAND, 5)
        AuthorSizer.Add(labelAuthor, 0, wx.CENTER|wx.ALL, 5)
        AuthorSizer.Add(self.inputAuthor, 1, wx.CENTER|wx.ALL|wx.EXPAND, 5)
        YearSizer.Add(labelYear, 0, wx.ALIGN_LEFT|wx.ALL, 5)
        YearSizer.Add(self.inputYear, 1, wx.ALIGN_LEFT|wx.ALL|wx.EXPAND, 5)
        PublisherSizer.Add(labelPublisher, 0, wx.ALIGN_LEFT|wx.ALL, 5)
        PublisherSizer.Add(self.inputPublisher, 1, wx.ALIGN_LEFT|wx.ALL|wx.EXPAND, 5)
        SubjectSizer.Add(SubjectButtonsSizer, 0, wx.ALIGN_LEFT|wx.ALL, 5)
        SubjectSizer.Add(self.inputSubject, -1, wx.EXPAND, 5)
        
        rightSideSizer.Add(bookTitleSizer, 0, wx.ALIGN_LEFT|wx.ALL|wx.EXPAND, 5)
        rightSideSizer.Add(ISBNSizer, 0, wx.ALIGN_LEFT|wx.ALL|wx.EXPAND, 5)
        rightSideSizer.Add(AuthorSizer, 0, wx.ALIGN_LEFT|wx.ALL|wx.EXPAND, 5)
        rightSideSizer.Add(YearSizer, 0, wx.ALIGN_LEFT|wx.ALL|wx.EXPAND, 5)
        rightSideSizer.Add(PublisherSizer, 0, wx.ALIGN_LEFT|wx.ALL|wx.EXPAND, 5)
        rightSideSizer.Add(SubjectSizer, -1, wx.ALIGN_LEFT|wx.ALL|wx.EXPAND, 5)
        rightSideSizer.Add(barcodeButton, 0, wx.ALIGN_LEFT|wx.ALL|wx.EXPAND, 5)
        rightSideSizer.Add(metaButton, 0, wx.ALIGN_LEFT|wx.ALL|wx.EXPAND, 5)
        rightSideSizer.Add(writeRowButton, 0, wx.ALIGN_LEFT|wx.ALL|wx.EXPAND, 5)
        
        metaSizer.Add(labelMetaSource, 0, wx.ALIGN_LEFT|wx.ALL, 5)
        metaSizer.Add(self.inputMetaSource, 0, wx.ALIGN_LEFT|wx.ALL, 5)
        mainSizer.Add(metaSizer, 0, wx.ALL, 5)
        mainSizer.Add(self.inputGrabSubj, 0, wx.ALIGN_LEFT|wx.ALL, 5)
        mainSizer.Add(self.inputGrabOnClick, 0, wx.ALIGN_LEFT|wx.ALL, 5)
        
        evenBiggerSizer.Add(mainSizer, 0, wx.CENTER|wx.ALL, 5)
        evenBiggerSizer.Add(imgSizer, 1, wx.CENTER|wx.ALL, 5)
        evenBiggerSizer.Add(rightSideSizer, 1, wx.CENTER|wx.ALL|wx.EXPAND, 5)
        

        self.SetSizer(evenBiggerSizer)
        
        self.Fit()
        
        self.SetSize
        
        self.Bind(wx.EVT_BUTTON, self.onCSV, csvButton)
        self.Bind(wx.EVT_BUTTON, self.onCancel, cancelButton)
        self.Bind(wx.EVT_BUTTON, self.onLoad, loadButton)
        self.Bind(wx.EVT_LISTBOX, self.onInspect, self.lst)
        self.Bind(wx.EVT_BUTTON, self.onBarcodeScan, barcodeButton)
        self.Bind(wx.EVT_BUTTON, self.onGrabMeta, metaButton)
        self.Bind(wx.EVT_BUTTON, self.onWriteRow, writeRowButton)
        self.Bind(wx.EVT_BUTTON, self.addTag, subject1Button)
        self.Bind(wx.EVT_BUTTON, self.addTag, subject2Button)
        self.Bind(wx.EVT_BUTTON, self.addTag, subject3Button)
        self.Bind(wx.EVT_BUTTON, self.onSkip, skipButton)
        
    def onSkip(self, event):
        img_path = self.RootDir + self.lst.GetString(self.lst.GetSelection())
        
        if not os.path.isdir(self.RootDir + "Skip/"):
            os.mkdir(self.RootDir + "Skip/")

        shutil.move(img_path, self.RootDir + "Skip/" + self.lst.GetString(self.lst.GetSelection()))
            
        # Now that the row is recorded and the file moved, we drop it from our listing, move the selection to
        # the top, clear the input fields, and re-trigger the inspection routine.
        
        self.inputTitle.SetValue("")
        self.inputISBN.SetValue("")
        self.inputAuthor.SetValue("")
        self.inputPublisher.SetValue("")
        self.inputYear.SetValue("")
        self.inputSubject.SetValue("")

        statusbar.SetStatusText('Moved ' + self.lst.GetString(self.lst.GetSelection()) +' to Skip folder')

        sel = self.lst.GetSelection()
        if sel != -1:
            self.lst.Delete(sel)
        self.lst.SetSelection(0)        
        self.onInspect(self)
        

        
    def addTag(self, event):
        new_txt = event.GetEventObject().GetLabel()
        
        subjString = self.inputSubject.GetValue()
        
        if len(subjString) < 1:
            subjString = new_txt
        else:
            subjString = subjString + "; " + new_txt
            
        self.inputSubject.SetValue(subjString)
        
    def onWriteRow(self, event):
        # We have to build a dictionary object with entries matching the column headers
        # used by the csv writer, then call it.

        #entry = {"ISBN-13":self.inputISBN.GetValue(), "Title":self.inputTitle.GetValue(), "Authors":self.inputAuthor.GetValue(),"Subject":self.inputSubject.GetValue(), "Year":self.inputYear.GetValue(), "Publisher":self.inputPublisher.GetValue()}
        entry = [self.inputISBN.GetValue(), self.inputTitle.GetValue(), self.inputAuthor.GetValue(), self.inputSubject.GetValue(), self.inputYear.GetValue(), self.inputPublisher.GetValue()]
        
        #self.writer.writerow( entry )
        self.csvfile.write("\t".join(entry)+"\n")
        
        # To ensure each row actually gets written, we will close and reopen the file
        self.csvfile.close()        
        self.csvfile = open(self.CSVPath, 'a')
        
        statusbar.SetStatusText('Wrote entry to CSV for file ' + self.lst.GetString(self.lst.GetSelection()) +' and moved to Success')

        # Lastly, we want to move the file to a Successfully Stored folder inside it's current folder. First step is
        # confirming the path exists, if it does not, creating it, then moving the file
        
        img_path = self.RootDir + self.lst.GetString(self.lst.GetSelection())
        
        if not os.path.isdir(self.RootDir + "Success/"):
            os.mkdir(self.RootDir + "Success/")

        shutil.move(img_path, self.RootDir + "Success/" + self.lst.GetString(self.lst.GetSelection()))
            
        # Now that the row is recorded and the file moved, we drop it from our listing, move the selection to
        # the top, clear the input fields, and re-trigger the inspection routine.
        
        self.inputTitle.SetValue("")
        self.inputISBN.SetValue("")
        self.inputAuthor.SetValue("")
        self.inputPublisher.SetValue("")
        self.inputYear.SetValue("")
        self.inputSubject.SetValue("")

        sel = self.lst.GetSelection()
        if sel != -1:
            self.lst.Delete(sel)
            
        if self.lst.GetCount() > 0:
            self.lst.SetSelection(0)        
            self.onInspect(self)
        
        
    def onCSV(self, event):
        dlg = wx.FileDialog (None, "Where should we save the output CSV?", defaultDir = '/Users/alexander_miles/Documents/Personal/Books', defaultFile = 'my_books.csv', wildcard = "Comma-delimited Plaintext (.csv)|.csv", style = wx.FD_SAVE )
        dlg.ShowModal()
        statusbar.SetStatusText('Opened' + dlg.GetPath() + ' for saving output')
        self.CSVPath = dlg.GetPath()
        
        self.csvfile = open(self.CSVPath, 'a')
        fieldnames = ['ISBN', 'Title', 'Authors', 'Subject', 'Year', 'Publisher']
        
        self.csvfile.write("\t".join(fieldnames)+"\n")
        
        #self.writer = csv.DictWriter(self.csvfile, fieldnames=fieldnames)
        #self.writer.writeheader()
        
    def onGrabMeta(self, event):
        meta_source = self.inputMetaSource.GetString( self.inputMetaSource.GetSelection() )
        if meta_source == 'Google Books':
            MS = 'goob'
        elif meta_source == 'Wikipedia':
            MS = 'wiki'
        elif meta_source =='OpenLibrary':
            MS = 'openl'
        else:
            print("Invalid metadata souce? But how?")
            
        number = str(self.inputISBN.GetValue())
        try:
            data = meta(number, service=MS)
        except:
            statusbar.SetStatusText("Error grabbing info from metadata provider " +MS+". Select a different provider")
            return
        # e.g. {'ISBN-13': '9781942993094', 'Title': 'My Neighbor Seki, 6', 'Authors': ['Takuma Morishige'], 'Publisher': 'Vertical Comics', 'Year': '2016', 'Language': 'en'}
        if 'Title' in data:
            self.inputTitle.SetValue( data['Title'] )
        if 'Authors' in data:
            if len(data['Authors']) == 1:
                self.inputAuthor.SetValue( data['Authors'][0] )
            else:
                 self.inputAuthor.SetValue(', '.join([A for A in data['Authors'] ]))
        if 'Publisher' in data:
            self.inputPublisher.SetValue( data['Publisher'] )
        if 'Year' in data:
            self.inputYear.SetValue( data['Year'] )
        
        # Checkbox for grabbing subject tags checked?
        if self.inputGrabSubj.GetValue() == True:
            try:
                number = str(self.inputISBN.GetValue())
                link = "https://openlibrary.org/api/books?bibkeys=ISBN:"+number+"&jscmd=data&format=JSON"
                #print(link)
                f = urllib.request.urlopen(link)
                web_result = json.loads(f.read())

                if 'ISBN:'+str(number) in web_result:
                    web_result = web_result['ISBN:'+str(number)]
                    if 'subjects' in web_result:
                        self.inputSubject.SetValue('; '.join( list(set( [ sub['name'].capitalize() for sub in web_result['subjects']] ))))
                else:
                    statusbar.SetStatusText("No metadata available from OpenLibrary for that ISBN!")
            except:
                statusbar.SetStatusText("Error grabbing info from OpenLibrary, server may be down")

        
    def onBarcodeScan(self, event):
        img_path = self.RootDir + self.lst.GetString(self.lst.GetSelection())
        img = Image.open( img_path )
        decode_result = decode(img)
        if len(decode_result) > 0:
            number = str( decode_result[0].data )[2:-1]
        else:
            number = ""
            statusbar.SetStatusText('Barcode could not be scanned, enter ISBN manually')
        self.inputISBN.SetValue( number )
        # print("Click! " + str(number))
        
        
    def onLoad(self, event):
        # Do file load-y stuff? Let's first see if we can change the context of the list box from here
        #self.lst.InsertItems(["A","B","C"], 0)
        
        # That worked! So let's extend that.. Let's first remove all existing entries from the box
        # Delete wants a zero-based item index.. So let's find how many items there are and just loop downards
        
        lastIndex = self.lst.GetCount()
        for I in range( lastIndex-1, -1, -1):
            self.lst.Delete(I)
        
        # self.lst.InsertItems(["A","B","C"], 0)
        
        # That works too. Now, can we open a window and have the user select a directory, glob for *.jpg,
        # and populate with the result?
        
        dlg = wx.DirDialog (None, "Choose input directory", '/Users/alexander_miles/Documents/Personal/Books', wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST)
        dlg.ShowModal()
        #print("Opened: " + dlg.GetPath())
        
        self.RootDir = dlg.GetPath()+"/"
        all_imgs = glob(self.RootDir + "*.jpeg")
        
        just_files = [img.split(self.RootDir)[1] for img in all_imgs]
        self.lst.InsertItems(just_files, 0)
        
        statusbar.SetStatusText('Found '+str(len(just_files))+ ' items in ' + dlg.GetPath())
        
    def onInspect(self, event):
        img_path = self.RootDir + self.lst.GetString(self.lst.GetSelection())
        bitmap = wx.Bitmap( img_path )
        bitmap = scale_bitmap(bitmap, self.PhotoMaxSize, self.PhotoMaxSize)        
        self.staticbitmap.SetBitmap(bitmap)
        
        if self.inputGrabOnClick.GetValue() == True:
            statusbar.SetStatusText('Grabbed barcode and metadata for file: ' + self.lst.GetString(self.lst.GetSelection()))
            self.onBarcodeScan(self)
            self.onGrabMeta(self)

    def onCancel(self, event):
        wx.Exit()

    
if __name__ == "__main__":
    app = MyApp()
    app.MainLoop()
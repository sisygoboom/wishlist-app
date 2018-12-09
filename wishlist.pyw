from decimal import *
from tkinter import *
from tkinter.filedialog import *
import pickle, webbrowser, time, tkinter.messagebox, operator

libs = []

try:
    from pyshorteners import Shortener
    libs.append(1)
except:
    libs.append(0)

try:
    import win32print, win32ui, win32clipboard, pythoncom
    from win32com.shell import shell, shellcon
    libs.append(1)
except: libs.append(0)

version = '1.1.0'

class Application(Frame):
    def __init__(self, master):

        Frame.__init__(self, master)
        self.grid()

        #Prerequisite variables
        self.googshop = IntVar()
        self.googshop.set(1)
        self.urlprint = IntVar()
        self.urlprint.set(1)
        self.auto = IntVar()
        self.order = StringVar()
        self.order.set("-----")
        self.filename = str()
        ###
        if libs[0] == 1:
            API_KEY = 'AIzaSyDeWKwZhgcTC2Z4RctB3EOybyi0K5ECzJE'
            self.shortener = Shortener('Google', api_key=API_KEY)
        ###

        #Frames

        #Frames - Define
        self.mainFrame = Frame(self)
        self.addFrame = Frame(self)
        
        #Frames - Setup
        for frame in (self.mainFrame, self.addFrame):
            frame.grid(row=0, column=0, sticky='news')

        #Frames - Setup - Main Page
        self.listbox = Listbox(self.mainFrame, selectmode=MULTIPLE, width=50, height=20)
        self.listbox.bind('<<ListboxSelect>>', self.onselect)
        self.addBut = Button(self.mainFrame, text='Add', command=self.addScreen)
        self.removeBut = Button(self.mainFrame, text='Remove', command=self.remove)
        self.searchBut = Button(self.mainFrame, text='Search', command=self.search)
        self.selAllBut = Button(self.mainFrame, text='(De)Select All', command=self.selectAll)

        self.applyBut = Button(self.mainFrame, text='Apply', command=self.populate)
        self.minSlider = Scale(self.mainFrame, orient=HORIZONTAL)
        self.maxSlider = Scale(self.mainFrame, orient=HORIZONTAL)
        self.sumLabel = Label(self.mainFrame, text='£0')
        self.orderOption = OptionMenu(self.mainFrame, self.order, '-----', 'Alphabetical', 'Price')
        self.filterBox = Entry(self.mainFrame)
        self.filterBox.bind('<Button-3>', self.popup)

        self.listbox.grid(rowspan=6, columnspan=3)
        self.addBut.grid(column=0, row=6, sticky='news')
        self.removeBut.grid(column=1, row=6, sticky='news')
        self.searchBut.grid(column=2, row=6, sticky='news')
        self.selAllBut.grid(column=3, row=6, sticky='news')
        self.orderOption.grid(row=0, column=3)
        self.minSlider.grid(row=1, column=3)
        self.maxSlider.grid(row=2, column=3)
        self.filterBox.grid(row=3, column=3)
        self.applyBut.grid(row=4, column=3)
        self.sumLabel.grid(row=5, column=3)

        #Frames - Setup - Add Page
        self.successLabel = Label(self.addFrame, text='Success', fg='green')
        self.nameLabel = Label(self.addFrame, text='Name:')
        self.priceLabel = Label(self.addFrame, text='Price: £')
        self.urlLabel = Label(self.addFrame, text='URL: ')

        self.nameBox = Entry(self.addFrame)
        self.priceBox = Entry(self.addFrame)
        self.urlBox = Entry(self.addFrame)

        self.nameBox.bind('<Button-3>', self.popup)
        self.priceBox.bind('<Button-3>', self.popup)
        self.urlBox.bind('<Button-3>', self.popup)

        self.submitBut = Button(self.addFrame, text='Submit', command=self.add)

        self.nameLabel.grid(row=0, column=0)
        self.priceLabel.grid(row=1, column=0)
        self.urlLabel.grid(row=2, column=0)
        self.nameBox.grid(row=0, column=1)
        self.priceBox.grid(row=1, column=1)
        self.urlBox.grid(row=2, column=1)
        self.submitBut.grid(row=3, column=0)

        #Menubar

        #Menubar - Main
        self.menuBar = Menu(self)
        self.menuBar.add_command(label='Main', command=self.mainScreen)

        #Menubar - File...
        menu = Menu(self.menuBar, tearoff=0)        
        self.menuBar.add_cascade(label='File...', menu=menu)
        menu.add_command(label='New', command=self.new)
        menu.add_command(label='Load...', command=self.load)
        #menu.add_command(label='Delete...')
        menu.add_command(label='Save', command=self.save)
        menu.add_command(label='Save As...', command=self.saveas)
        menu.add_command(label='Print', command=self.printer)

        #Menubar - Edit...
        menu = Menu(self.menuBar, tearoff=0)
        self.menuBar.add_cascade(label='Edit...', menu=menu)
        menu.add_command(label='Add', command=self.addScreen)
        menu.add_command(label='Remove', command=self.remove)
        menu.add_command(label='Search', command=self.search)

        #Menubar - Options...
        menu = Menu(self.menuBar, tearoff=0)
        self.menuBar.add_cascade(label='Options...', menu=menu)
        menu.add_checkbutton(label='Include URLs when printing', variable=self.urlprint)
        menu.add_checkbutton(label='Google Shopping alternatives', variable=self.googshop, onvalue=1, offvalue=0)
        menu.add_checkbutton(label='Autosave', variable=self.auto, onvalue=1, offvalue=0)

        #Menubar Configuration
        try:self.master.config(menu=self.menuBar)
        except AttributeError:
            self.master.tk.call(master, 'config', '-menu', self.menuBar)

        self.new()


    def mainScreen(self):
        self.populate(reset=True)
        self.mainFrame.tkraise()

    def addScreen(self):
        self.addFrame.tkraise()

    def populate(self, reset=False):
        size = len(self.wlist)
        self.calibrate(reset)
        self.listbox.delete(0, END)
        if size > 0:
            vals = self.findVal()
            minval = self.minSlider.get()
            maxval = self.maxSlider.get()
            keyw = (self.filterBox.get()).lower()
            tempList = dict()
            for k, v in self.wlist.items():
                p = float(v['price'])
                if p > minval:
                    if p < maxval:
                        if keyw in k.lower():
                            tempList[k] = p

            if self.order.get() == 'Alphabetical':
                tempList = sorted(tempList.items(), key=operator.itemgetter(0))

            if self.order.get() == 'Price':
                tempList = sorted(tempList.items(), key=operator.itemgetter(1))

            if self.order.get() == '-----':
                tempList = tempList.items()

            for n in tempList:
                self.listbox.insert(END, (n[0] + ' = £' + str(n[1])))

    def calibrate(self, reset):
        size = len(self.wlist)
        if size > 0:
            vals = self.findVal()
            self.minSlider.config(from_=vals[0], to=vals[1])
            self.maxSlider.config(from_=vals[2], to=vals[3])
            if reset==True:
                self.minSlider.set(vals[0])
                self.maxSlider.set(vals[3])
                self.filterBox.delete(0, END)

    def findVal(self):
        null = list()
        for k, v in self.wlist.items():
            temp = v['price']
            null.append(float(temp))
        lowest = float(min(null))
        highest = float(max(null))
        minLow = (lowest - 1)
        maxLow = (highest - 1)
        minHigh = (lowest + 1)
        maxHigh = (highest + 1)
        compiled = (minLow, maxLow, minHigh, maxHigh)
        return(compiled)

    def new(self):
        self.wlist = dict()
        self.sumLabel['text'] = "£0"
        self.mainScreen()

    def save(self):
        try:
            pickle.dump(self.wlist, open(self.filename, 'wb'))
        except:
            self.saveas()
            
    def autosave(self):
        if self.auto.get() == 1:
            self.save()

    def saveas(self):
        extensionInfo = ('pickle file', '*.p')
        self.filename = asksaveasfilename(filetypes=[extensionInfo], defaultextension=extensionInfo)
        pickle.dump(self.wlist, open(self.filename, 'wb'))

    def load(self):
        try:
            self.filename = askopenfilename(filetypes=[('pickle file', '*.p')])
            self.wlist = pickle.load(open(str(self.filename), 'rb'))
        except FileNotFoundError:
            print('file not found')
        self.mainScreen()
        self.sumLabel['text'] = "£0"

    def getname(self, i, method):
        listing = self.listbox.get(i)
        location = listing.rfind(' = £')
        if method == 'n':
            value = listing[:location]
        elif method == 'p':
            value = listing[location+4:]
        return value

    def search(self):
        tempList = list()
        for i in self.listbox.curselection():
            name = self.getname(i, 'n')
            x = self.wlist[name]["url"]
            name = name.replace(" ", "+")
            y = 'https://www.google.co.uk/search?q=' + name + '&tbm=shop&safe=active&ssui=on'
            if self.googshop.get() == 1:
                tempList.append(y)
            tempList.append(x)
        for i in tempList:
            webbrowser.open_new_tab(i)
            if tempList.index(i) == 0:
                time.sleep(2)

    def remove(self):
        tempList = list()
        for i in self.listbox.curselection():
            name = self.getname(i, 'n')
            tempList.append(name)
        for i in tempList: del self.wlist[i]
        self.sumLabel['text'] = "£0"
        self.populate()
        self.autosave()

    def add(self):
        name = self.nameBox.get().strip()
        price = self.priceBox.get().strip()
        url = self.urlBox.get().strip()
        i = False

        if libs[0] == 1:
            
            try:
                url = self.shortener.short(url)
                self.urlBox.delete(0, END)
                self.urlBox.insert(0, url)
            except:
                i = tkinter.messagebox.askyesno('Warning',
'''The url you entered did not respond,
if your PC has access to the internet, this url may be broken.
it is also possible that the URL has already been shortened.
Would you like to change the url?''')

        if i == False:
            try:
                price = Decimal(price)
                price = price.quantize(Decimal('0.01'), rounding=ROUND_UP)
                self.wlist[name] = {'price':price, 'url':url}
                self.success()
            except:
                tkMessageBox.showwarning("Invalid Input", "Please only use numbers and decimal points in the price field.")
                print('error')
            self.autosave()

    def onselect(self, evt=None):
        index = self.listbox.curselection()
        total = 0
        for i in range(len(index)):
            price = Decimal(self.getname(index[i], 'p'))
            price = price.quantize(Decimal('.01'), rounding=ROUND_UP)
            total += price
        self.sumLabel['text'] = '£' + str(total)

    def selectAll(self):
        if len(self.listbox.curselection()) == 0:
            self.listbox.selection_set(0, END)
            self.onselect()
        else:
            self.listbox.selection_clear(0, END)
            self.sumLabel['text'] = '£0'

    def success(self):
        self.successLabel.grid(row=3, column=1)
        self.after(2000, lambda: self.successLabel.grid_remove())

    def cut(self):
        self.copy()
        self.widget.delete('sel.first', 'sel.last')

    def copy(self):
        data = self.widget.selection_get()
        r = Tk()
        r.withdraw()
        r.clipboard_clear()
        r.clipboard_append(data)
        r.destroy()

    def paste(self):
        text = self.selection_get(selection='CLIPBOARD')
        self.widget.insert('insert', text)

    def popup(self, event):
        self.widget = event.widget
        rcmenu = Menu(self, tearoff=0)
        rcmenu.add_command(label='cut', command=self.cut)
        rcmenu.add_command(label='copy', command=self.copy)
        rcmenu.add_command(label='paste', command=self.paste)
        
        rcmenu.post(event.x_root, event.y_root)

    def printer(self):
        if libs[1]==1:
            self.Y = 50
            names = []
            prices = []
            urls = []

            for k,v in self.wlist.items():
                names.append(k)
                prices.append(v['price'])
                urls.append(v['url'])

            hDC = win32ui.CreateDC()
            hDC.CreatePrinterDC(win32print.GetDefaultPrinter())
            hDC.StartDoc('wishlist')
            hDC.StartPage()

            for name in names:
                nameLength = len(name)
                nameIndex = names.index(name)
                url = urls[nameIndex]
                urlLength = len(url)
                price = str(prices[nameIndex])

                lines = [name[i:i+40] for i in range(0, nameLength, 40)]
                urllines = [url[i:i+21] for i in range(0, urlLength, 21)]

                hDC.TextOut(3000, self.Y, price)

                if self.urlprint.get() == 1:
                    startY = self.Y
                    for urlline in urllines:
                        hDC.TextOut(3500, self.Y, urlline)
                        self.down(1)
                    lowpoint = self.Y
                    self.Y = startY

                for line in lines:
                    hDC.TextOut(50, self.Y, line)
                    self.down(1)

                if lowpoint > self.Y:
                    self.Y = lowpoint
                self.down(2)

            hDC.EndPage()
            hDC.EndDoc()

        elif libs[1]==0:
            tkinter.messagebox.showwarning('missing library',
                                           'Please install the\
pywin32 python package or ask me for the compiled version before using this feature.')
            
    def down(self, n):
        self.Y += 100*n

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath('.')

    if not os.path.exists(os.path.join(base_path, relative_path)):
        os.makedirs(relative_path)

    return os.path.join(base_path, relative_path)

root = Tk()
root.title('Wishlist')
root.wm_iconbitmap(resource_path('icon.ico'))
root.resizable(False, False)
app = Application(root)

root.mainloop()

#Programmer: Brent E. Chambers (q0m)
#Date: 02/20/2014
#Description: This program is a part of the DEFTRN-FOR training module.  
# Files needed: deftrn.py (Primary Access Class for the XLS Data)
#               deftrn-workbook.xls (Resource File [xls])
#               deftrn-g.py (GUI front end)


import Tkinter as tk
import tkMessageBox
import deftrn
import sets

def clear_forms():
	techniqueBox.delete(0, tk.END)
	try:
		techniqueBox.delete('1.0', tk.END)
		syntaxBox.delete('1.0', tk.END)
		descriptionBox.delete('1.0', tk.END)
	except:
		pass
			
def get_list(event):
	try:
		syntaxBox.delete('1.0', '40.0')
		descriptionBox.delete('1.0', '40.0')
	except:
		pass
	index = techniqueBox.curselection()[0]
	seltext = techniqueBox.get(index)
	indexFinder = liveInstance.tArray.index(seltext)
	#print indexFinder
	#print liveInstance.sArray[indexFinder]
	#print liveInstance.dArray[indexFinder]
	syntaxBox.insert('1.0', str(liveInstance.sArray[indexFinder]))
	descriptionBox.insert('1.0', str(liveInstance.dArray[indexFinder]))

def change_mode(mode):
	#try:
	#	techniqueBox.delete(0, tk.END)
	#except:
	#	pass
	sheet = liveInstance.loadCurrentMode(mode)
	liveInstance.populateData(sheet)
	clear_forms()
	for item in liveInstance.CurrentModeData[0]:
		techniqueBox.insert(0, item)
	
def about():
		tkMessageBox.showinfo("About", "Combat Informatics Training \n Developed by Brent E. Chambers (q0m) \n http://www.brentchambers.net")
		

		
def search():
	results = liveInstance.newsearch(searchBox.get()) #calls the newsearch function of the pentrn instance
	clear_forms()
	for item in results:
		techniqueBox.insert(0, item[0])
	
		
	
####################### GRAPHICAL USER INTERFACE #########################
#initiate graphical user interface
root = tk.Tk()
root.resizable(width='True', height='False')
root.title("DEFTRN Forensics Training Program v1.0")
root.geometry("1175x330+50+200") #expanded this and maybe apply a word wrap
s = "Choose a selection"
#label = tk.Label(root, text=s, width=25)
#label.grid(row=0, column=0)
#techinque box frame
searchBox = tk.Entry(root)
searchBox.grid(row=4, column=0, sticky='e'+'w')#, rowspan=1)

searchButton = tk.Button(root, text='SEARCH', command=search)
searchButton.grid(row=4, column=0, sticky='e', rowspan=2)

clearButton = tk.Button(root, text="CLEAR", command=clear_forms)
clearButton.grid(row=4, column=1, sticky='w', columnspan=2)

techniqueBox = tk.Listbox(root, height=19, width=75, bg='blue', fg='white')
techniqueBox.grid(row=1, column=0, rowspan=3, sticky='n'+'s')
techniqueBox.bind('<ButtonRelease-1>', get_list)
techniqueBox.bind('<Up>', get_list)
techniqueBox.bind('<Down>', get_list)
yscroll = tk.Scrollbar(command=techniqueBox.yview, orient=tk.VERTICAL)
yscroll.grid(row=1, column=1, sticky='n'+'s', rowspan=3)
#syntax box frame

syntaxBox = tk.Text(root, height=14, width=85, wrap='word')
syntaxBox.grid(row=1, column=2, sticky='n'+'s')
yscroll = tk.Scrollbar(command=syntaxBox.yview, orient=tk.VERTICAL)
yscroll.grid(row=1, column=3, sticky='n'+'s')
#description box frame

descriptionBox = tk.Text(root, height=4, width=85, wrap='word')
descriptionBox.grid(row=2, column=2, sticky='n'+'s')
yscroll = tk.Scrollbar(command=descriptionBox.yview, orient=tk.VERTICAL)
yscroll.grid(row=2, column=3, sticky='n'+'s')
#File Menu Presentation and Commands

menubar = tk.Menu(root)
filemenu = tk.Menu(menubar, tearoff=0)
filemenu.add_separator()
filemenu.add_command(label="Exit", command=root.quit)
menubar.add_cascade(label="File", menu=filemenu)


#subfunctions are needed to define the change from different disciplines
discipMenu = tk.Menu(menubar, tearoff=0)
liveInstance = deftrn.PenInstance()
for item in liveInstance.secMode:
	discipMenu.add_command(label=item, command= lambda mode=item:change_mode(mode))
menubar.add_cascade(label="Discipline", menu=discipMenu)


helpMenu = tk.Menu(menubar, tearoff=0)
helpMenu.add_command(label="About", command=about) #command=root.popup Brent Chambers brentchambers.net
menubar.add_cascade(label="Help", menu=helpMenu)

root.config(menu=menubar)
root.mainloop()
from tkinter import *
import os
import winshell
import win32com.client
from tkinter import filedialog
from PIL import Image

iconPath = r"%systemroot%\system32\imageres.dll"
IconName = ""


def fix_path(datapath):
    fixedpath = datapath.replace("\\", "\\\\")
    return fixedpath


def generate_label(text1, x1):
    my_label = Label(master=Top_Frame, text=text1, width=55, fg='blue', bg="#C0C0C0")
    my_label.grid(row=x1, column=0, ipadx=10, ipady=5)


def browseFiles():
    filename = filedialog.askopenfilename(initialdir="/", title="Select a File",
                                          filetypes=(("Apps", "*.exe*"),
                                                     ("all files", "*.*")))
    global x
    global BatText
    x += 1
    BatText = BatText + "start \"\" \"" + fix_path(filename) + "\"\n "
    if len(filename) > 1:
        generate_label(filename, x)

    # Change label contents
    # label_file_explorer.configure(text="Data Path: " + filename)


def browsePng():
    filename = fix_path(filedialog.askopenfilename(initialdir="/", title="Select a File",
                                                   filetypes=(("photos", ('.png', '.ico')),
                                                              ("all files", "*.*"))))
    global iconPath
    if ".png" in filename:
        img = Image.open(fix_path(filename))
        img.save(filename.replace(".png", ".ico"))

    iconPath = filename.replace(".png", ".ico")
    print(iconPath)


def getShortcutName():
    global IconName
    global Entry_Name
    if len(Entry_Name.get()) < 1:
        return "Neue_Verknüpfung"
    else:
        return Entry_Name.get()


def end_file():
    path_to_batFile = f"{newpath}\\{getShortcutName()}.bat"
    myBat = open(path_to_batFile, 'w+')
    myBat.writelines(BatText)
    myBat.close()
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(os.path.join(winshell.desktop(), f'{getShortcutName()}.lnk'))
    shortcut.TargetPath = path_to_batFile
    shortcut.IconLocation = fix_path(iconPath)
    shortcut.save()
    root.destroy()


root = Tk()
Bottom_Frame = Frame(root, height=100, width=300)
Bottom_Frame.grid(row=1, column=0, sticky=S)
Label(Bottom_Frame, text="Name:").grid(row=0, column=0, columnspan=2,)
Entry_Name = Entry(Bottom_Frame, width=40)
Entry_Name.grid(row=1, column=0, columnspan=2, sticky=S)
x = 0
BatText = "@echo off\n"
user = os.path.expanduser('~')
newpath = os.path.expanduser('~') + "\\" + "Shortcut"
if not os.path.exists(newpath):
    os.makedirs(newpath)
root.title('Shortcuts Maker')
root.resizable(width=False, height=True)
root.geometry("400x540")
Top_Frame = Frame(root, height=400, width=400)
Top_Frame.grid(row=0, column=0)
root.rowconfigure(1, weight=1)
root.columnconfigure(1, weight=1)
add_Path = Button(Bottom_Frame, text="Dateipfad hinzufügen", command=browseFiles)
add_Path.grid(row=3, column=1, sticky=S)
add_icon = Button(Bottom_Frame, text="Symbol hinzufügen", command=browsePng)
add_icon.grid(row=3, column=0)
Done_file = Button(Bottom_Frame, text="Verknüpfung erstellen", command=end_file)
Done_file.grid(row=4, column=0, columnspan=2)
Label(Bottom_Frame, text="Drücken sie auf \"Dateipfad hinzufügen\" \num eine Datei oder App zur Verknüpfung "
                         "hinzuzufügen",
      fg="#C0C0C0").grid(row=2, column=0, sticky=N, columnspan=2)
print(iconPath)
root.mainloop()

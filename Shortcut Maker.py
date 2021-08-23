import os
import webbrowser
from tkinter import *
from tkinter import filedialog

import win32com.client
import winshell
from PIL import Image
from PyInstaller.utils.hooks import collect_data_files
from tkinterdnd2 import *

datas = collect_data_files('tkinterdnd2')

iconPath = r"%systemroot%\system32\imageres.dll"
IconName = ""


def fix_path(datapath):
    fixedpath = datapath.replace("\\", "\\\\")
    return fixedpath


def generate_label(text1):
    lb.insert("end", text1)


def add_via_dnd(event):
    global x
    global BatText
    x += 1
    text = fix_path(event.data)
    text2 = text.replace("{", "")
    text3 = text2.replace("}", "")
    BatText = BatText + "start \"\" \"" + text3 + "\"\n "
    lb.insert("end", text3)


def browseFiles():
    filename = filedialog.askopenfilename(initialdir="/", title="Select a File",
                                          filetypes=(("Alle Dateien", "*.*"),
                                                     ("Apps", "*.exe*"),
                                                     ("Fotos", "*.png")))
    global x
    global BatText
    x += 1
    BatText = BatText + "start \"\" \"" + fix_path(filename) + "\"\n "
    if len(filename) > 1:
        generate_label(filename)


def browsePng():
    filename = fix_path(filedialog.askopenfilename(initialdir="/", title="Foto auswählen:",
                                                   filetypes=(("photos", ('.png', '.jpg', '.ico')),
                                                              ("all files", "*.*"))))
    global iconPath
    print(filename)
    if ".png" in filename:
        img = Image.open(fix_path(filename))
        img.save(filename.replace(".png", ".ico"))
        iconPath = filename.replace(".png", ".ico")

    elif ".jpg" in filename:
        img1 = Image.open(fix_path(filename))
        img1.save(filename.replace(".jpg", ".ico"))
        iconPath = filename.replace(".jpg", ".ico")

    else:
        iconPath = filename

    print(iconPath)


def callback(event):
    webbrowser.open_new(event.widget.cget("text"))


def getShortcutName():
    global IconName
    global Entry_Name
    if len(Entry_Name.get()) < 1:
        return "Neue_Verknüpfung"
    else:
        return Entry_Name.get()


def end_file():
    global iconPath
    if len(iconPath) < 3:
        iconPath = r"%systemroot%\system32\imageres.dll"

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
Bottom_Frame.rowconfigure(0, weight=1)
Bottom_Frame.columnconfigure(0, weight=1)
Label(Bottom_Frame, text="Shortcut Name:").grid(row=0, column=0, columnspan=2, )
Entry_Name = Entry(Bottom_Frame, width=40)
Entry_Name.grid(row=1, column=0, columnspan=2, sticky=S)
x = 0
BatText = "@echo off\n"
user = os.path.expanduser('~')
newpath = os.path.expanduser('~') + "\\" + "Shortcut"
if not os.path.exists(newpath):
    os.makedirs(newpath)
root.title('Shortcut Maker')
root.resizable(width=False, height=True)
root.geometry("400x350")
Top_Frame = Frame(root, height=400, width=400)
Top_Frame.grid(row=0, column=0)
Top_Frame.drop_target_register(DND_FILES)
Top_Frame.dnd_bind('<<Drop>>', add_via_dnd)
root.rowconfigure(0, weight=1)
root.columnconfigure(1, weight=1)
lb = Listbox(Top_Frame, width=69, height=70, bd=0, selectmode=SINGLE, justify=LEFT, bg="#c4c4c4")
lb.grid(row=0, column=0, sticky=N)
add_Path = Button(Bottom_Frame, text="Dateipfad hinzufügen", command=browseFiles)
add_Path.grid(row=3, column=1, sticky=S, ipadx=5)
add_icon = Button(Bottom_Frame, text="Symbol hinzufügen", command=browsePng)
add_icon.grid(row=3, column=0, ipadx=5)
Done_file = Button(Bottom_Frame, text="Verknüpfung erstellen", command=end_file)
Done_file.grid(row=4, column=0, columnspan=2)
lbl = Label(Bottom_Frame, text=r"https://twitter.com/pryz208", height=0, fg="#0e526c", cursor="hand2")
lbl.grid(row=5, column=0, columnspan=1)
Label(Bottom_Frame, text="©2021 Beta V1.0", fg="#0e526c").grid(row=5, column=1)
Label(Bottom_Frame, text="Drücken sie auf \"Dateipfad hinzufügen\", \noder ziehen sie eine Datei auf das obere graue "
                         "Feld",
      fg="#4f4f4f").grid(row=2, column=0, sticky=N, columnspan=2)
lbl.bind("<Button-1>", callback)
root.mainloop()

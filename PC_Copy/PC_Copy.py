from tkinter import *
from distutils import dir_util


class CopyLabel:

    def __init__(self, master, row, column, text):
        self.text = text
        self.label = Label(master, text=text)
        self.label.grid(row=row, column=column, sticky=E)


class CopyEntry:

    def __init__(self, master, row, column, defaulttext):
        self.entry = Entry(master)
        self.entry.insert(END, defaulttext)
        self.entry.grid(row=row, column=column)


class CopyCheckbutton:

    def __init__(self, master, row, column, text):
        self.var = IntVar()
        self.text = text
        self.button = Checkbutton(master, text=text, var=self.var)
        self.button.grid(row=row, column=column, sticky=W)


class CopyButton:

    def __init__(self, master, row, column, text, command):
        self.label = Button(master, text=text, command=command)
        self.label.grid(row=row, column=column)


def copy_selected():
    for button in checkbuttons:
        if button.var.get():
            copy(button.text)


def copy_all():
    for button in checkbuttons:
        copy(button.text)


def copy(category):
    old_computer = oldEntry.entry.get()
    new_computer = newEntry.entry.get()
    username = usernameEntry.entry.get()

    if category == "Desktop":
        print(f"Copying Desktop...")
        dir_util.copy_tree(src=f"\\\\{old_computer}/itadmin$/users/{username}/Desktop", dst=f"\\\\{new_computer}/itadmin$/users/{username}/Desktop")

    if category == "Favourites":
        print(f"Copying Favourites...")
        dir_util.copy_tree(src=f"\\\\{old_computer}/itadmin$/users/{username}/Favorites", dst=f"\\\\{new_computer}/itadmin$/users/{username}/Favorites")

    if category == "Documents":
        print(f"Copying Documents...")
        dir_util.copy_tree(src=f"\\\\{old_computer}/itadmin$/users/{username}/Documents", dst=f"\\\\{new_computer}/itadmin$/users/{username}\\Documents")

    if category == "Outlook":
        print(f"Copying Outlook...")
        dir_util.copy_tree(src=f"\\\\{old_computer}/itadmin$/users/{username}/AppData/Roaming/Microsoft/Outlook", dst=f"\\\\{new_computer}/itadmin$/users/{username}/AppData/Roaming/Microsoft/Outlook")
        dir_util.copy_tree(src=f"\\\\{old_computer}/itadmin$/users/{username}/AppData/Roaming/Microsoft/Signatures", dst=f"\\\\{new_computer}/itadmin$/users/{username}/AppData/Roaming/Microsoft/Signatures")

    if category == "Pictures":
        print(f"Copying Pictures...")
        dir_util.copy_tree(src=f"\\\\{old_computer}/itadmin$/users/{username}/Pictures", dst=f"\\\\{new_computer}/itadmin$/users/{username}/Pictures")

    if category == "Apollo":
        print(f"Copying Apollo...")
        dir_util.copy_tree(src=f"\\\\{old_computer}/itadmin$/fp/datadir/pkeys", dst=f"\\\\{new_computer}/itadmin$/fp/datadir/pkeys")
        dir_util.copy_tree(src=f"\\\\{old_computer}/itadmin$/fp/datadir/users", dst=f"\\\\{new_computer}/itadmin$/fp/datadir/users")
        dir_util.copy_tree(src=f"\\\\{old_computer}/itadmin$/fp/machine/dat32com.ini", dst=f"\\\\{new_computer}/itadmin$/fp/machine/dat32com.ini")

    print("Transfer complete.")


# Initialize root window.
root = Tk()
root.title("EZPZ - PC Copy")
root.iconbitmap("./Assets/ezpz.ico")

# Input fields.
oldLabel = CopyLabel(root, 0, 0, "Old PC Name")
oldEntry = CopyEntry(root, 0, 1, "")
newLabel = CopyLabel(root, 1, 0, "New PC Name")
newEntry = CopyEntry(root, 1, 1, "")
usernameLabel = CopyLabel(root, 2, 0, "Username")
usernameEntry = CopyEntry(root, 2, 1, "")

# Checkbuttons.
desktopCheck = CopyCheckbutton(root, 3, 0, "Desktop")
favouritesCheck = CopyCheckbutton(root, 3, 1, "Favourites")
documentsCheck = CopyCheckbutton(root, 4, 0, "Documents")
outlookCheck = CopyCheckbutton(root, 4, 1, "Outlook")
picturesCheck = CopyCheckbutton(root, 5, 0, "Pictures")
apolloCheck = CopyCheckbutton(root, 5, 1, "Apollo")
checkbuttons = [desktopCheck, favouritesCheck, documentsCheck, outlookCheck, picturesCheck, apolloCheck]

# Submit buttons.
copyselectedButton = CopyButton(root, 7, 0, "Copy Selected", copy_selected)
copyallButton = CopyButton(root, 7, 1, "Copy All", copy_all)


root.mainloop()

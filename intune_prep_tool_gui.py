import json
from tkinter import Tk, Label, Button, Entry, filedialog, IntVar, Checkbutton
from tkinter.font import Font
import os
import subprocess
import sys
from pathlib import Path


class IntuneWinAppUtilGUI:
    def __init__(self, master):
        self.master = master
        master.title("Microsoft Win32 Content Prep Tool GUI by Joshua Dwight")

        # Initialize default values
        self.install_file_path = ""
        self.output_dir_path = ""
        self.setup_folder = ""

        # Load saved paths from JSON if it exists
        # Determine if running as script or frozen exe
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))

        # Update the path for the JSON file
        self.json_file_path = os.path.join(application_path, 'preptool.json')

        if os.path.exists(self.json_file_path):
            with open(self.json_file_path, 'r') as f:
                saved_paths = json.load(f)
                self.install_file_path = saved_paths.get('install_file_path', '')
                self.output_dir_path = saved_paths.get('output_dir_path', '')
                self.setup_folder = os.path.dirname(self.install_file_path)

        # Set window dimensions and disable resizing
        master.geometry("600x325")
        master.resizable(False, False)

        self.label_pkg_name = Label(master, text="Package Name")
        self.label_pkg_name.pack()
        self.entry_pkg_name = Entry(master, width=50)
        self.entry_pkg_name.pack()

        # Add some space
        Label(master, text="").pack()

        self.label_install_file = Label(master, text="Install File")
        self.label_install_file.pack()
        self.entry_install_file = Entry(master, width=50)
        self.entry_install_file.pack()
        self.button_install_file = Button(master, text="Choose Install File (F1)", command=self.choose_install_file)
        self.button_install_file.pack()

        # Add some space
        Label(master, text="").pack()

        self.label_output_dir = Label(master, text="Output Directory")
        self.label_output_dir.pack()
        self.entry_output_dir = Entry(master, width=50)
        self.entry_output_dir.pack()
        self.button_output_dir = Button(master, text="Choose Output Directory (F2)", command=self.choose_output_dir)
        self.button_output_dir.pack()

        # Add some space
        Label(master, text="").pack()

        # Create and style the "Create Package" button
        self.create_package_button_font = Font(size=18, weight="bold")
        self.button_create_package = Button(master, text="Create Package (F5)", command=self.create_package, font=self.create_package_button_font)
        self.button_create_package.pack()

        # Checkbox for opening output directory
        self.open_output_var = IntVar()
        self.open_output_checkbox = Checkbutton(self.master, text="Open Output Directory (F3)", variable=self.open_output_var)
        self.open_output_checkbox.pack()

        # Key Bindings
        self.master.bind('<F1>', lambda event: self.choose_install_file())
        self.master.bind('<F2>', lambda event: self.choose_output_dir())
        self.master.bind('<F3>', lambda event: self.toggle_open_output_checkbox())
        self.master.bind('<F5>', lambda event: self.create_package())
        self.master.bind('<F12>', lambda event: self.clear_all_fields())

    # Sample function to toggle the "Open Output Directory" checkbox
    def toggle_open_output_checkbox(self):
        current_value = self.open_output_var.get()
        self.open_output_var.set(not current_value)

    # Sample function to clear all fields
    def clear_all_fields(self):
        self.entry_pkg_name.delete(0,'end')
        self.entry_install_file.delete(0, 'end')
        self.entry_output_dir.delete(0, 'end')
        self.open_output_var.set(0)
        # Add more fields to clear if needed

    def choose_install_file(self):
        initial_dir = self.setup_folder if self.setup_folder else "/"
        self.install_file_path = filedialog.askopenfilename(initialdir=initial_dir, title="Select Install File", filetypes=(("Installer files", "*.exe;*.msi;*.cmd;*.bat;*.ps1"), ("All files", "*.*")))
        self.entry_install_file.delete(0, 'end')
        self.entry_install_file.insert(0, self.install_file_path)
        self.setup_folder = os.path.dirname(self.install_file_path)

    def choose_output_dir(self):
        initial_dir = self.output_dir_path if self.output_dir_path else "/"
        self.output_dir_path = filedialog.askdirectory(initialdir=initial_dir, title="Select Output Directory")
        self.entry_output_dir.delete(0, 'end')
        self.entry_output_dir.insert(0, self.output_dir_path)

    def create_package(self):
        # Update attributes based on manual entries
        self.install_file_path = self.entry_install_file.get()
        self.output_dir_path = self.entry_output_dir.get()
        self.setup_folder = os.path.dirname(self.install_file_path)
        
        if not self.install_file_path or not self.output_dir_path:
            print("Please choose both an install file and an output directory before proceeding.")
            return

        # Save paths to JSON
        with open(self.json_file_path, 'w') as f:
            json.dump({
                'install_file_path': self.install_file_path,
                'output_dir_path': self.output_dir_path
            }, f)

        if getattr(sys, 'frozen', False):
            current_dir = os.path.dirname(sys.executable)
        else:
            current_dir = os.path.dirname(os.path.abspath(__file__))




        intune_tool_path = os.path.join(current_dir, "IntuneWinAppUtil")
        package_name = self.entry_pkg_name.get()
        intune_command = f'"{intune_tool_path}" -c "{self.setup_folder}" -s "{self.install_file_path}" -o "{self.output_dir_path}" -q'
        print(f"Executing: {intune_command}")
        subprocess.run(intune_command, shell=True)

        # Find the .intunewin file in the output directory that corresponds to the install file
        output_dir = Path(self.output_dir_path)
        install_file_name = os.path.basename(self.install_file_path).rsplit('.', 1)[0]
        corresponding_intune_file = next(output_dir.glob(f"{install_file_name}*.intunewin"), None)

        if corresponding_intune_file is not None:
            # Rename it to match the package name
            renamed_intune_file = output_dir / f"{package_name}.intunewin"
            corresponding_intune_file.rename(renamed_intune_file)
        else:
            print("Could not find corresponding .intunewin file to rename.")

        # Open the output directory if checkbox is checked
        if self.open_output_var.get() == 1:
            absolute_output_path = os.path.abspath(self.output_dir_path)
            subprocess.call(f'explorer "{absolute_output_path}"')


# Run the application
root = Tk()
app = IntuneWinAppUtilGUI(root)
root.mainloop()

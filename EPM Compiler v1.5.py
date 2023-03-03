import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

class EPMCompiler:
    def __init__(self, master):
        self.master = master
        master.title("Elevated Plus Maze Compiler")

        # Set Discord-style color scheme
        self.color_scheme = {
            "background": "#d3eeea",
            "foreground": "#FFFFFF",
            "accent": "#7289DA",
            "button_bg": "#454546",
            "button_fg": "#FFFFFF",
        }

        # Set window background color
        master.config(bg=self.color_scheme["background"])

        icon_path = os.path.join(os.path.dirname(__file__), 'icon.ico')
        master.iconbitmap(default=icon_path)

        # Create label and buttons
        self.label = tk.Label(
            master,
            text=" ",
            bg=self.color_scheme["background"],
            fg=self.color_scheme["foreground"],
        )
        self.label.pack()

        self.button_frame = tk.Frame(master, bg=self.color_scheme["background"])
        self.button_frame.pack()

        self.compile_button = tk.Button(
            self.button_frame,
            text="Compile",
            command=self.compile_data,
            padx=40,
            bg=self.color_scheme["button_bg"],
            fg=self.color_scheme["button_fg"],
            activebackground=self.color_scheme["button_bg"],
            activeforeground=self.color_scheme["button_fg"],
            bd=3,
        )
        self.compile_button.pack(side=tk.LEFT, padx=10, pady=10)

        self.export_button = tk.Button(
            self.button_frame,
            text="Export",
            command=self.export_data,
            padx=40,
            bg=self.color_scheme["button_bg"],
            fg=self.color_scheme["button_fg"],
            activebackground=self.color_scheme["button_bg"],
            activeforeground=self.color_scheme["button_fg"],
            bd=3,
        )
        self.export_button.pack(side=tk.RIGHT, padx=10, pady=10)

        # Load logo image
        logo_path = os.path.join(os.path.dirname(__file__), 'logo.png')
        self.logo_image = tk.PhotoImage(file=logo_path)

        # Create logo label
        self.logo_label = tk.Label(master, image=self.logo_image, bg=self.color_scheme["background"])
        self.logo_label.pack()

        # Create info menu
        self.info_menu = tk.Menu(master, bg=self.color_scheme["button_bg"], fg=self.color_scheme["button_fg"])
        self.info_menu.add_command(label="Info", command=self.show_info)
        master.config(menu=self.info_menu)

    def compile_data(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            data = []
            for file_name in os.listdir(folder_path):
                if file_name.endswith('.txt'):
                    with open(os.path.join(folder_path, file_name)) as f:
                        lines = f.readlines()
                        row = {
                            'File name': file_name,
                            'Closed Explore': lines[18][14:21],
                            'Closed Entrance': lines[18][27:34],
                            'Closed Arm Duration': lines[18][39:46],
                            'Open Explore': lines[20][14:21],
                            'Open Entrance': lines[20][27:34],
                            'Open Arm Duration': lines[20][39:48],
                            'Junction Time': lines[15][8:15],

                            'Start Date': lines[4][13:21],
                            'End Date': lines[5][10:19],
                            'Subject': lines[6][9:22],
                            'Experiment': lines[7][12:22],
                            'Group': lines[8][7:22],
                            'Box': lines[9][5:7],

                            'Start Time': lines[10][12:22],
                            'End Time': lines[11][10:22],
                            'MSN': lines[12][5:50],
                            'A': lines[13][5:15],
                            'B': lines[14][5:15],
                            'S': lines[16][5:15],

                            'R01': lines[22][16:21],
                            'R02': lines[22][28:34],
                            'R03': lines[22][40:50],
                            'R04': lines[22][53:60],
                            'R05': lines[22][67:73],

                        }
                        data.append(row)
            self.data = pd.DataFrame(data)
            messagebox.showinfo("Compilation", "Data compiled successfully. Now select Export and choose name and location of your .xslx file")

    def export_data(self):
        if hasattr(self, 'data'):
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx")
            if file_path:
                 # Convert columns to numeric data types
                numeric_cols = ['Closed Explore', 'Closed Entrance', 'Closed Arm Duration',
                'Open Explore', 'Open Entrance', 'Open Arm Duration', 'Junction Time', 'R01', 'R02', 'R03', 'R04', 'R05']
                self.data[numeric_cols] = self.data[numeric_cols].apply(pd.to_numeric)

                 # Export data to Excel
                self.data.to_excel(file_path, index=False)
                messagebox.showinfo("Export", "Data exported successfully.")
        else:
            messagebox.showwarning("Export", "No data to export.")

    def show_info(self):
        messagebox.showinfo("Info", "Elevated Plus Maze Compiler was created to take the Elevated Plus Maze text data files produced by Med Associates Med-PC software and compile it into one easy to analyze excel sheet. Created on 2.23.23 By Paul Hamblin")



root = tk.Tk()
app = EPMCompiler(root)
root.mainloop()

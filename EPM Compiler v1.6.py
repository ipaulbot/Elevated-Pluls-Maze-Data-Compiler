import os
import pandas as pd
import tkinter as tk
import webbrowser
from tkinter import filedialog, messagebox
from tkinter.font import Font

class EPMCompiler:
    def __init__(self, master):
        self.master = master
        master.title("Elevated Plus Maze Compiler")

        # set fixed window size
        master.geometry("305x450")

        # Set style color scheme
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
        # Create new window
        info_window = tk.Toplevel(self.master)
        info_window.title("About EPM Compiler")
        info_window.config(bg=self.color_scheme["background"])

        # set fixed window size
        info_window.geometry("305x450")

        # Create text widget to display information
        info_text = tk.Text(
            info_window,
            bd=0,
            padx=20,
            pady=20,
            wrap=tk.WORD,
            font=("Arial", 8),
            bg=self.color_scheme["background"],
            fg=self.color_scheme["button_bg"],
        )
        info_text.pack()

        # Insert text into widget
        info_text.insert(
            tk.END,
            "Elevated Plus Maze Compiler v1.6\n"
            "Created by Paul Hamblin 2023\n\n"
            "This program compiles Elevated Plus Maze data from text files generated by Med-Associates "
            "behavior tracking software.\n\n"
            "Instructions:\n\n"
            "1. Click 'Compile' and select the folder containing the EPM text files.\n\n"
            "2. Click 'Export' and select the name and location of the output file."
            "The compiled data will be exported to an Excel file (.xlsx).\n"

            "\n\nContact Information:\n"
            "Email: pbhamblin@gmail.com\n"
            "GitHub: github.com/ipaulbot\n"
            "Blog: Igrok\n"
        )

        # Create a font with drop shadow effect
        my_font = Font(family="Arial", size=12, weight="bold")

        # Modify font weight of specific tag
        info_text.tag_add("bold", "1.0", "1.35")
        info_text.tag_config("bold", font=my_font, foreground="#f3cf6f")

        # Make the links clickable
        info_text.tag_add("email_link", "14.6", "14.30")
        info_text.tag_config("email_link", foreground="blue", underline=0)
        info_text.tag_bind("email_link", "<Button-1>", lambda e: webbrowser.open("mailto:pbhamblin@gmail.com"))

        info_text.tag_add("github_link", "15.9", "15.30")
        info_text.tag_config("github_link", foreground="blue", underline=0)
        info_text.tag_bind("github_link", "<Button-1>", lambda e: webbrowser.open("https://github.com/ipaulbot"))

        info_text.tag_add("blog_link", "16.6", "16.30")
        info_text.tag_config("blog_link", foreground="blue", underline=0)
        info_text.tag_bind("blog_link", "<Button-1>", lambda e: webbrowser.open("https://transapient.blogspot.com/2023/03/automating-data-organization-with-ai-my.html"))

        # Disable text widget editing
        info_text.config(state=tk.DISABLED)

root = tk.Tk()
app = EPMCompiler(root)
root.mainloop()

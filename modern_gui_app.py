import customtkinter as ctk
import os
import threading
from tkinter import filedialog, messagebox
from PIL import Image

# Import the refactored scripts
import sys

# Get absolute paths to the subdirectories
if getattr(sys, 'frozen', False):
    # If the application is run as a bundle, the PyInstaller bootloader
    # extends the sys module by a flag frozen=True and sets the app 
    # path into variable _MEIPASS'.
    current_dir = sys._MEIPASS
else:
    current_dir = os.path.dirname(os.path.abspath(__file__))

res_script_path = os.path.join(current_dir, 'reslivemain')
manage_script_path = os.path.join(current_dir, 'resvaduvlive')

# Add to sys.path if not already present
if res_script_path not in sys.path:
    sys.path.append(res_script_path)
if manage_script_path not in sys.path:
    sys.path.append(manage_script_path)

# Import residentialscript
residentialscript = None
residential_error = None
try:
    import residentialscript
except ImportError as e:
    residential_error = str(e)
    print(f"Error importing residentialscript: {e}")

# Import manage_builtup_area
manage_builtup_area = None
manage_error = None
try:
    import manage_builtup_area
except ImportError as e:
    manage_error = str(e)
    print(f"Error importing manage_builtup_area: {e}")

ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Real Estate Data Manager")
        self.geometry("900x600")

        # set grid layout 1x2
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # load images with light and dark mode image
        image_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "assets")
        # You can add icons here if you have them, for now we use text

        # create navigation frame
        self.navigation_frame = ctk.CTkFrame(self, corner_radius=0)
        self.navigation_frame.grid(row=0, column=0, sticky="nsew")
        self.navigation_frame.grid_rowconfigure(4, weight=1)

        self.navigation_frame_label = ctk.CTkLabel(self.navigation_frame, text="  Real Estate Tools",
                                                             compound="left", font=ctk.CTkFont(size=15, weight="bold"))
        self.navigation_frame_label.grid(row=0, column=0, padx=20, pady=20)

        self.home_button = ctk.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Residential Script",
                                                   fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                   anchor="w", command=self.home_button_event)
        self.home_button.grid(row=1, column=0, sticky="ew")

        self.frame_2_button = ctk.CTkButton(self.navigation_frame, corner_radius=0, height=40, border_spacing=10, text="Manage Builtup Area",
                                                      fg_color="transparent", text_color=("gray10", "gray90"), hover_color=("gray70", "gray30"),
                                                      anchor="w", command=self.frame_2_button_event)
        self.frame_2_button.grid(row=2, column=0, sticky="ew")

        # create home frame (Residential Script)
        self.home_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.home_frame.grid_columnconfigure(0, weight=1)

        self.home_label = ctk.CTkLabel(self.home_frame, text="Residential Data Cleaning", font=ctk.CTkFont(size=20, weight="bold"))
        self.home_label.grid(row=0, column=0, padx=20, pady=10, sticky="w")

        self.res_file_entry = ctk.CTkEntry(self.home_frame, placeholder_text="Select Input Excel File")
        self.res_file_entry.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        
        self.res_browse_btn = ctk.CTkButton(self.home_frame, text="Browse", command=self.browse_res_file)
        self.res_browse_btn.grid(row=1, column=1, padx=20, pady=10)

        self.res_run_btn = ctk.CTkButton(self.home_frame, text="Run Process", command=self.run_residential_script)
        self.res_run_btn.grid(row=2, column=0, padx=20, pady=10, sticky="ew")

        self.res_open_btn = ctk.CTkButton(self.home_frame, text="Open Output Folder", command=self.open_res_output, state="disabled", fg_color="green")
        self.res_open_btn.grid(row=2, column=1, padx=20, pady=10, sticky="ew")

        self.res_log_box = ctk.CTkTextbox(self.home_frame, width=400, height=300)
        self.res_log_box.grid(row=3, column=0, padx=20, pady=10, sticky="nsew", columnspan=2)
        self.home_frame.grid_rowconfigure(3, weight=1)

        # create second frame (Manage Builtup)
        self.second_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.second_frame.grid_columnconfigure(0, weight=1)

        self.second_label = ctk.CTkLabel(self.second_frame, text="Manage Builtup Area", font=ctk.CTkFont(size=20, weight="bold"))
        self.second_label.grid(row=0, column=0, padx=20, pady=10, sticky="w")

        self.area_file_entry = ctk.CTkEntry(self.second_frame, placeholder_text="Select Area File (Excel)")
        self.area_file_entry.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        self.area_browse_btn = ctk.CTkButton(self.second_frame, text="Browse", command=self.browse_area_file)
        self.area_browse_btn.grid(row=1, column=1, padx=20, pady=10)

        self.floor_file_entry = ctk.CTkEntry(self.second_frame, placeholder_text="Select Floor File (Excel)")
        self.floor_file_entry.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        self.floor_browse_btn = ctk.CTkButton(self.second_frame, text="Browse", command=self.browse_floor_file)
        self.floor_browse_btn.grid(row=2, column=1, padx=20, pady=10)

        self.manage_run_btn = ctk.CTkButton(self.second_frame, text="Run Process", command=self.run_manage_script)
        self.manage_run_btn.grid(row=3, column=0, padx=20, pady=10, sticky="ew")

        self.manage_open_btn = ctk.CTkButton(self.second_frame, text="Open Output Folder", command=self.open_manage_output, state="disabled", fg_color="green")
        self.manage_open_btn.grid(row=3, column=1, padx=20, pady=10, sticky="ew")

        self.manage_log_box = ctk.CTkTextbox(self.second_frame, width=400, height=300)
        self.manage_log_box.grid(row=4, column=0, padx=20, pady=10, sticky="nsew", columnspan=2)
        self.second_frame.grid_rowconfigure(4, weight=1)

        # select default frame
        self.select_frame_by_name("home")
        
        self.res_output_path = None
        self.manage_output_path = None

    def select_frame_by_name(self, name):
        # set button color for selected button
        self.home_button.configure(fg_color=("gray75", "gray25") if name == "home" else "transparent")
        self.frame_2_button.configure(fg_color=("gray75", "gray25") if name == "frame_2" else "transparent")

        # show selected frame
        if name == "home":
            self.home_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.home_frame.grid_forget()
        if name == "frame_2":
            self.second_frame.grid(row=0, column=1, sticky="nsew")
        else:
            self.second_frame.grid_forget()

    def home_button_event(self):
        self.select_frame_by_name("home")

    def frame_2_button_event(self):
        self.select_frame_by_name("frame_2")

    def browse_res_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if filename:
            self.res_file_entry.delete(0, "end")
            self.res_file_entry.insert(0, filename)

    def browse_area_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if filename:
            self.area_file_entry.delete(0, "end")
            self.area_file_entry.insert(0, filename)

    def browse_floor_file(self):
        filename = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if filename:
            self.floor_file_entry.delete(0, "end")
            self.floor_file_entry.insert(0, filename)

    def log_res(self, message):
        self.res_log_box.insert("end", str(message) + "\n")
        self.res_log_box.see("end")

    def log_manage(self, message):
        self.manage_log_box.insert("end", str(message) + "\n")
        self.manage_log_box.see("end")
        
    def open_file_or_folder(self, path):
        if not path or not os.path.exists(path):
            messagebox.showerror("Error", "Output file not found.")
            return
        try:
            # Open the file itself
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file: {e}")

    def open_res_output(self):
        if self.res_output_path:
            self.open_file_or_folder(os.path.dirname(self.res_output_path))

    def open_manage_output(self):
        if self.manage_output_path:
            self.open_file_or_folder(os.path.dirname(self.manage_output_path))

    def run_residential_script(self):
        file_path = self.res_file_entry.get()
        if not file_path:
            messagebox.showerror("Error", "Please select an input file.")
            return
        
        self.res_log_box.delete("1.0", "end")
        self.res_run_btn.configure(state="disabled")
        self.res_open_btn.configure(state="disabled")
        
        def task():
            try:
                if residentialscript:
                    output = residentialscript.process_residential_data(file_path, log_callback=self.log_res)
                    if output and os.path.exists(output):
                        self.res_output_path = output
                        self.res_open_btn.configure(state="normal")
                        # Auto open
                        try:
                            os.startfile(output)
                        except:
                            pass
                else:
                    self.log_res(f"Error: residentialscript module not loaded.\nDetails: {residential_error}")
            except Exception as e:
                self.log_res(f"Critical Error: {e}")
            finally:
                self.res_run_btn.configure(state="normal")
        
        threading.Thread(target=task, daemon=True).start()

    def run_manage_script(self):
        area_file = self.area_file_entry.get()
        floor_file = self.floor_file_entry.get()

        if not area_file or not floor_file:
            messagebox.showerror("Error", "Please select both Area and Floor files.")
            return

        self.manage_log_box.delete("1.0", "end")
        self.manage_run_btn.configure(state="disabled")
        self.manage_open_btn.configure(state="disabled")

        def task():
            try:
                if manage_builtup_area:
                    output = manage_builtup_area.main(area_file, floor_file, log_callback=self.log_manage)
                    if output and os.path.exists(output):
                        self.manage_output_path = output
                        self.manage_open_btn.configure(state="normal")
                        # Auto open
                        try:
                            os.startfile(output)
                        except:
                            pass
                else:
                    self.log_manage(f"Error: manage_builtup_area module not loaded.\nDetails: {manage_error}")
            except Exception as e:
                self.log_manage(f"Critical Error: {e}")
            finally:
                self.manage_run_btn.configure(state="normal")

        threading.Thread(target=task, daemon=True).start()

if __name__ == "__main__":
    app = App()
    app.mainloop()

import tkinter as tk
from PIL import Image, ImageTk
from tkinter import ttk
from tkinter import filedialog
from tkinter import simpledialog
import subprocess
import os
import random
import shutil
from tkinter import messagebox
from docx import Document
from PIL import Image, ImageTk
from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont
from tkinter.scrolledtext import ScrolledText
import xlwings as xw
import tempfile

libreoffice_path = r"C:\Program Files\LibreOffice\program"

# ---------- LIBREOFFICE FILE OPENERS ----------

# Opens a single file with the appropriate LibreOffice application based on its extension
def open_with_LibreOffice():
    file_path = filedialog.askopenfilename(title="Choose a file to open")
    if file_path:
        ext = os.path.splitext(file_path)[1].lower()
        if ext in ['.ods', '.xls', '.xlsx']:  # Spreadsheet files
            subprocess.Popen(["C:\\Program Files\\LibreOffice\\program\\scalc.exe", file_path])
        elif ext in ['.odt', '.doc', '.docx']:  # Word processing files
            subprocess.Popen(["C:\\Program Files\\LibreOffice\\program\\swriter.exe", file_path])
        else:  # Fallback for other types
            subprocess.Popen(["C:\\Program Files\\LibreOffice\\program\\soffice.exe", file_path])

# Opens all supported files within a selected folder using the appropriate LibreOffice application
def open_folder_with_LibreOffice():
    folder_path = filedialog.askdirectory(title="Choose a folder to open files from")
    if folder_path:
        for filename in os.listdir(folder_path):
            file_path = os.path.join(folder_path, filename)
            if os.path.isfile(file_path):
                ext = os.path.splitext(file_path)[1].lower()
                if ext in ['.ods', '.xls', '.xlsx']:
                    subprocess.Popen([os.path.join(libreoffice_path, "scalc.exe"), file_path])
                elif ext in ['.odt', '.doc', '.docx']:
                    subprocess.Popen([os.path.join(libreoffice_path, "swriter.exe"), file_path])
                else:
                    subprocess.Popen([os.path.join(libreoffice_path, "soffice.exe"), file_path])

# Lets the user copy and rename a file, then opens it with LibreOffice
def copy_and_rename_and_open_with_LibreOffice():
    source_file_path = filedialog.askopenfilename(title="Choose a file to copy")
    if source_file_path:
        dest_folder = filedialog.askdirectory(title="Choose a folder to copy the file to")
        if dest_folder:
            new_file_name = simpledialog.askstring("Rename File", "Enter a new name for the file (with extension):")
            if new_file_name:
                try:
                    # Ensure new file name has the same extension
                    ext = os.path.splitext(source_file_path)[1]
                    if not new_file_name.endswith(ext):
                        new_file_name += ext

                    dest_file_path = os.path.join(dest_folder, new_file_name)
                    shutil.copy(source_file_path, dest_file_path)  # Perform the file copy

                    # Open copied file with the appropriate LibreOffice program
                    ext = os.path.splitext(dest_file_path)[1].lower()
                    if ext in ['.ods', '.xls', '.xlsx']:
                        subprocess.Popen([os.path.join(libreoffice_path, "scalc.exe"), dest_file_path])
                    elif ext in ['.odt', '.doc', '.docx']:
                        subprocess.Popen([os.path.join(libreoffice_path, "swriter.exe"), dest_file_path])
                    else:
                        subprocess.Popen([os.path.join(libreoffice_path, "soffice.exe"), dest_file_path])

                    messagebox.showinfo("Success", f"File copied and opened with LibreOffice.")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to copy the file: {e}")
            else:
                messagebox.showwarning("No Name", "You must provide a new name for the file.")

# ---------- CAMPAIGN CREATION ----------

# Allows user to create a new campaign folder with subfolders
def make_new_campaign():
    campaign_location = filedialog.askdirectory(title="Choose a location for the new campaign folder")
    if campaign_location:
        campaign_name = simpledialog.askstring("Campaign Name", "Enter the name for the new campaign:")
        if campaign_name:
            try:
                new_campaign_path = os.path.join(campaign_location, campaign_name)

                # Create folders only if they don't already exist
                if not os.path.exists(new_campaign_path):
                    os.makedirs(new_campaign_path)
                    os.makedirs(os.path.join(new_campaign_path, "DM Only"))
                    os.makedirs(os.path.join(new_campaign_path, "Players"))
                    os.makedirs(os.path.join(new_campaign_path, "General"))

                    messagebox.showinfo("Success", f"New campaign folder '{campaign_name}' created with subfolders.")
                else:
                    messagebox.showwarning("Folder Exists", "A folder with that name already exists.")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create the campaign folder: {e}")
        else:
            messagebox.showwarning("No Name", "You must provide a name for the campaign folder.")

class QuickSwitchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DiceScribe - QuickSwitch")

        # Create a notebook widget to act as tabbed interface
        self.notebook = tk.Notebook(root)
        self.notebook.pack(expand=True, fill='both')

        self.pages = {}       # Keeps track of open pages
        self.history = []     # Tracks tab switch history for Ctrl+Tab navigation

        # Add tabs/pages for different tools
        self.add_page("Notes", "View and write notes.") 
        self.add_page("Character Sheets", "View and create Character Sheets.")
        self.add_page("Campaigns", "View and create Campiagns.")
        self.add_page("Damage Calculator", "Calculate Damage And Saving Rolls.")

        # Bind mouse and keyboard events for tab switching
        self.notebook.bind("<Button-1>", self.track_page)
        self.notebook.bind_all("<Control-Tab>", self.next_tab)
        self.notebook.bind_all("<Control-Shift-Tab>", self.prev_tab)

    def add_page(self, title, text=""):
        frame = tk.Frame(self.notebook)
        label = tk.Label(frame, text=text, padding=10)
        label.pack(expand=True)

        # Assign specific buttons depending on tab title
        if title == "Character Sheets":
            ttk.Button(frame, text="Open Character Sheet with LibreOffice", command=open_with_LibreOffice).pack(pady=10)
            ttk.Button(frame, text="Use a Character Sheet Template", command=copy_and_rename_and_open_with_LibreOffice).pack(pady=10)

        if title == "Notes":
            ttk.Button(frame, text="Open Notes with LibreOffice", command=open_with_LibreOffice).pack(pady=10)

        if title == "Campaigns":
            ttk.Button(frame, text="Open all Campaign files", command=open_folder_with_LibreOffice).pack(pady=10)
            ttk.Button(frame, text="Create New Campaign", command=make_new_campaign).pack(pady=10)

        self.notebook.add(frame, text=title)
        self.pages[title] = frame

    # Keeps track of last visited tab for tab-switching
    def track_page(self, event=None):
        current_tab = self.notebook.tab(self.notebook.select(), "text")
        if current_tab in self.history:
            self.history.remove(current_tab)
        self.history.insert(0, current_tab)

    # Move to the next tab using Ctrl+Tab
    def next_tab(self, event=None):
        if self.history:
            current_tab = self.notebook.tab(self.notebook.select(), "text")
            try:
                index = self.history.index(current_tab)
                next_index = (index + 1) % len(self.history)
                self.select_tab(self.history[next_index])
            except ValueError:
                pass

    # Move to the previous tab using Ctrl+Shift+Tab
    def prev_tab(self, event=None):
        if self.history:
            current_tab = self.notebook.tab(self.notebook.select(), "text")
            try:
                index = self.history.index(current_tab)
                prev_index = (index - 1) % len(self.history)
                self.select_tab(self.history[prev_index])
            except ValueError:
                pass

    # Programmatically select a tab by name
    def select_tab(self, title):
        for tab_id in self.notebook.tabs():
            if self.notebook.tab(tab_id, "text") == title:
                self.notebook.select(tab_id)
                self.track_page()
                break

class DiceScribeApp:
    import tempfile

    def make_background_frame(self, parent):
        frame = tk.Frame(parent)
        bg_label = tk.Label(frame, image=self.bg_photo)
        bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        return frame

    def __init__(self, root):
        self.root = root
        self.root.title("DiceScribe")
        self.root.geometry("900x600")

        # Load and resize the background image for aesthetic
        bg_image = Image.open("paper-1914901_1280.jpg")
        bg_image = bg_image.resize((900, 600))  # Match window size
        self.bg_photo = ImageTk.PhotoImage(bg_image)

        # Create tabbed layout
        self.page_control = ttk.Notebook(root)
        self.page_control.pack(expand=True, fill='both')

        # Pages with background image applied
        self.home_page = self.make_background_frame(self.page_control)
        self.directory_page = self.make_background_frame(self.page_control)
        self.editor_page = self.make_background_frame(self.page_control)
        self.damage_page = self.make_background_frame(self.page_control)

        # Add pages to notebook
        self.page_control.add(self.home_page, text="\U0001F3E0 Home")
        self.page_control.add(self.directory_page, text="\U0001F4D8 Directory")
        self.page_control.add(self.editor_page, text="\U0001F5C2\uFE0F Editor")
        self.page_control.add(self.damage_page, text="\U0001F3AF Damage Calculator")

        self.pages = {}  # Tab name → widget (editor or image)
        self.history = []  # Track recently visited tabs

        # Initialize individual page setups
        self.build_home()
        self.build_directory()
        self.build_editor()
        self.build_damage_calc()

    def build_home(self):
        # Top navigation frame
        top_frame = tk.Frame(self.home_page)
        top_frame.pack(fill='x', pady=10)

        # Buttons to jump to Directory or Editor
        ttk.Button(top_frame, text="\U0001F4D8 Go to Directory", command=lambda: self.page_control.select(self.directory_page)).pack(side='left', padx=5)
        ttk.Button(top_frame, text="\U0001F5C2\uFE0F Go to Editor", command=lambda: self.page_control.select(self.editor_page)).pack(side='left', padx=5)

        # Welcome message
        desc = ("Welcome to DiceScribe — your all-in-one template and file editor!\n\n"
                "DiceScribe is a simple, intuitive platform that allows you to import and edit your gameplay files,\n"
                "and utilize one of our many customizable templates to speed up your note-taking and campaign management.\n\n"
                "To get started, just head to the Directory and import your first file.")

        label = tk.Label(self.home_page, text=desc, font=("Arial", 12), bg="#f6e2b0", wraplength=800, justify="center")
        label.pack(pady=20)

    def build_damage_calc(self):
        container = ttk.Frame(self.damage_page, padding=10)
        container.pack(expand=True, fill='both')

         # Set the background image
        bg_label = tk.Label(container, image=self.bg_photo)
        bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        # Overlay frame to hold widgets, transparent or matched color
        content = tk.Frame(container, bg='#c1a37e')  # Optional: match background image
        content.place(relx=0.5, rely=0.5, anchor='center')  # Center the content on the page

        # UI labels and entry fields
        tk.Label(container, text="Damage Calculator", font=("Arial", 16, "bold"), bg="#f6e2b0").pack(pady=(0, 10))
        tk.Label(container, text="Enter Number of Dice:", bg="#f6e2b0").pack(pady=5)
        self.dice_count_entry = tk.Entry(container, width=20, bg="#f6e2b0")
        self.dice_count_entry.pack(pady=5)

        tk.Label(container, text="Enter Sides of Die:", bg="#f6e2b0").pack(pady=5)
        self.dice_sides_entry = tk.Entry(container, width=20, bg="#f6e2b0")
        self.dice_sides_entry.pack(pady=5)

        tk.Label(container, text="Enter Modifier(+ , -):", bg="#f6e2b0").pack(pady=5)
        self.modifier_entry = tk.Entry(container, width=20, bg="#f6e2b0")
        self.modifier_entry.pack(pady=5)

        # Calculate damage button
        ttk.Button(container, text="Calculate Damage", command=self.calculate_damage).pack(pady=10)
        self.result_label = tk.Label(container, text="Result: ", font=("Arial", 12), bg="#f6e2b0")
        self.result_label.pack(pady=20)

    # Dice roll logic
    def calculate_damage(self):
        try:
            dice_count = int(self.dice_count_entry.get())
            dice_sides = int(self.dice_sides_entry.get())
            modifier = int(self.modifier_entry.get()) if self.modifier_entry.get() else 0

            # Roll all the dice
            dice_rolls = [random.randint(1, dice_sides) for _ in range(dice_count)]
            total_roll = sum(dice_rolls)
            total_damage = total_roll + modifier

            # Display results
            self.result_label.config(text=f"Result: {total_damage} (You Rolled: {', '.join(map(str, dice_rolls))}, Modifier: {modifier})")

        except ValueError:
            self.result_label.config(text="Error: Please enter valid integers for all fields.")

    def build_directory(self):
        container = ttk.Frame(self.directory_page, padding=10)
        container.pack(expand=True, fill='both')

        tk.Label(container, text="Directory", font=("Arial", 16, "bold"), bg="#f6e2b0").pack(pady=(0, 10))

        # Predefined sample scenes
        self.scenes = [
            ("Session 12: Betrayal at the Keep", "The party is betrayed by a trusted ally inside the fortress."),
            ("Session 13: The Oracle's Warning", "The oracle delivers a cryptic warning about an approaching threat."),
            ("Session 14: Final Battle Under the Moon", "The final showdown takes place beneath a blood-red moon."),
        ]

        self.scene_button_container = ttk.Frame(container)
        self.scene_button_container.pack(expand=True, fill='both')

        # Create a button for each predefined scene
        for title, summary in self.scenes:
            self.create_scene_button(title, summary)

        # Button to allow importing new notes from file
        ttk.Button(container, text="➕ Add Note from File", command=self.add_note_from_file).pack(pady=10)

    def create_scene_button(self, title, summary):
        btn = ttk.Button(self.scene_button_container, text=title,
                         command=lambda: self.open_scene_in_editor(title, summary))
        btn.pack(fill='x', pady=4)

    def add_note_from_file(self):
        file_path = filedialog.askopenfilename(
            title="Select a Note File",
            filetypes=[
                ("Text Files", "*.txt"),
                ("Word Documents", "*.docx"),
                ("Excel Files", "*.xlsx"),
                ("Image Files", "*.jpg; *.jpeg; *.png; *.gif")
            ]
        )

        if file_path:
            title = os.path.basename(file_path)
            ext = os.path.splitext(file_path)[1].lower()
            content = ""

            try:
                # Read file based on type
                if ext == ".txt":
                    with open(file_path, "r", encoding="utf-8") as file:
                        content = file.read()
                elif ext == ".docx":
                    doc = Document(file_path)
                    content = "\n".join([para.text for para in doc.paragraphs])
                elif ext == ".xlsx":
                    wb = load_workbook(file_path, data_only=True)
                    sheet = wb.active
                    content = "\n".join([
                        "\t".join([str(cell.value) if cell.value else "" for cell in row])
                        for row in sheet.iter_rows()
                    ])
                elif ext in [".jpg", ".jpeg", ".png", ".gif"]:
                    self.add_image_page(title, file_path)
                    self.page_control.select(self.editor_page)
                    return
                else:
                    print("Unsupported file type.")
                    return

                # Add new scene button and save content
                self.scenes.append((title, content))
                self.create_scene_button(title, content)

            except Exception as e:
                print(f"Error reading file: {e}")

    def add_image_page(self, title, image_path):
        try:
            from PIL import Image

            pil_img = Image.open(image_path)
            photo = ImageTk.PhotoImage(pil_img)

            frame = tk.Frame(self.notebook)

            # Canvas and scrollbars
            canvas = tk.Canvas(frame, width=800, height=600)
            h_scroll = ttk.Scrollbar(frame, orient="horizontal", command=canvas.xview)
            v_scroll = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
            canvas.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

            canvas.grid(row=0, column=0, sticky="nsew")
            v_scroll.grid(row=0, column=1, sticky="ns")
            h_scroll.grid(row=1, column=0, sticky="ew")

            frame.grid_rowconfigure(0, weight=1)
            frame.grid_columnconfigure(0, weight=1)

            # Frame inside canvas to hold the image
            image_frame = tk.Frame(canvas)
            image_label = tk.Label(image_frame, image=photo)
            image_label.image = photo  # Prevent garbage collection
            image_label.pil_image = pil_img  # Keep reference for printing
            image_label.pack()

            canvas.create_window((0, 0), window=image_frame, anchor="nw")
            canvas.config(scrollregion=(0, 0, photo.width(), photo.height()))

            # Add print button
            ttk.Button(frame, text="🖨️ Print", command=self.print_current_editor_page).grid(
                row=2, column=0, columnspan=2, pady=5
            )

            self.notebook.add(frame, text=title)
            self.pages[title] = image_label
            self.notebook.select(frame)

            # --- Mousewheel scroll support ---
            def _on_mousewheel(event):
                if event.state & 0x1:  # Shift is held → horizontal scroll
                    canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
                else:
                    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", lambda e: canvas.yview_scroll(-1, "units"))
            canvas.bind_all("<Button-5>", lambda e: canvas.yview_scroll(1, "units"))

        except Exception as e:
            messagebox.showerror("Error", f"Error loading image: {e}")

    def open_scene_in_editor(self, title, content):
        if title in self.pages:
            self.select_tab(title)
        else:
            self.add_editor_page(title, content)
            self.select_tab(title)
        self.page_control.select(self.editor_page)

    def build_editor(self):
        tk.Label(self.editor_page, text="Editor", font=("Arial", 16, "bold"), bg="#f6e2b0").pack(pady=(0, 10))
        self.build_search_bar()

        self.notebook = ttk.Notebook(self.editor_page)
        self.notebook.pack(expand=True, fill='both')

        # Preloaded pages
        self.add_editor_page("Notes Template", "...template text...")
        self.add_character_sheet_form()
        self.add_editor_page("Campaign Template", "...template text...")

        # Bind tab switching and tracking
        self.notebook.bind("<Button-1>", self.track_page)
        self.notebook.bind_all("<Control-Tab>", self.next_tab)
        self.notebook.bind_all("<Control-Shift-Tab>", self.prev_tab)
        self.notebook.bind_all("<Control-p>", lambda event: self.print_current_editor_page())

    def build_search_bar(self):
        search_frame = ttk.Frame(self.editor_page)
        search_frame.pack(pady=5, fill='x')

        ttk.Label(search_frame, text="Search:").pack(side="left", padx=(10, 5))
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40)
        search_entry.pack(side="left", padx=(0, 10))
        search_entry.bind("<KeyRelease>", self.highlight_search)

    def highlight_search(self, event=None):
        if self.is_search_empty():
            current_tab_id = self.notebook.select()
            if current_tab_id:
                current_tab_title = self.notebook.tab(current_tab_id, "text")
                text_widget = self.pages.get(current_tab_title)
                if text_widget:
                    text_widget.tag_remove("highlight", "1.0", tk.END)
            return

        search_term = self.search_var.get()
        if len(search_term.strip()) < 4:
            self.highlight_search()  # Clear highlights if too short
            return

        current_tab_id = self.notebook.select()
        if not current_tab_id:
            return

        current_tab_title = self.notebook.tab(current_tab_id, "text")
        text_widget = self.pages.get(current_tab_title)
        if not text_widget:
            print(f"No text widget found for tab {current_tab_title}.")
            return

        # Highlight search matches
        text_widget.tag_remove("highlight", "1.0", tk.END)
        start = "1.0"
        while True:
            start = text_widget.search(search_term, start, nocase=True, stopindex=tk.END)
            if not start:
                break
            end = f"{start}+{len(search_term)}c"
            text_widget.tag_add("highlight", start, end)
            start = end
        text_widget.tag_config("highlight", background="yellow", foreground="black")

        def is_search_empty(self):
            return self.search_var.get().strip() == ""

    def add_editor_page(self, title, content):
        frame = tk.Frame(self.notebook)
        text_widget = ScrolledText(frame, wrap="word", font=("Arial", 11))
        text_widget.insert("1.0", content)
        text_widget.pack(expand=True, fill='both')

        ttk.Button(frame, text="💾 Save", command=lambda: self.save_content(title, text_widget)).pack(pady=5)
        ttk.Button(frame, text="🖨️ Print", command=self.print_current_editor_page).pack(pady=5)
        self.notebook.add(frame, text=title)
        self.pages[title] = text_widget

    def save_content(self, title, text_widget):
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", initialfile=title,
                                                 filetypes=[("Text Files", "*.txt")])
        if file_path:
            try:
                content = text_widget.get("1.0", tk.END)
                with open(file_path, "w", encoding="utf-8") as file:
                    file.write(content)
            except Exception as e:
                print(f"Error saving file: {e}")

    def track_page(self, event=None):
        current_tab = self.notebook.tab(self.notebook.select(), "text")
        if current_tab in self.history:
            self.history.remove(current_tab)
        self.history.insert(0, current_tab)

    def next_tab(self, event=None):
        if self.history:
            current_tab = self.notebook.tab(self.notebook.select(), "text")
            try:
                index = self.history.index(current_tab)
                next_index = (index + 1) % len(self.history)
                self.select_tab(self.history[next_index])
            except ValueError:
                pass

    def prev_tab(self, event=None):
        if self.history:
            current_tab = self.notebook.tab(self.notebook.select(), "text")
            try:
                index = self.history.index(current_tab)
                prev_index = (index - 1) % len(self.history)
                self.select_tab(self.history[prev_index])
            except ValueError:
                pass

    def select_tab(self, title):
        for tab_id in self.notebook.tabs():
            if self.notebook.tab(tab_id, "text") == title:
                self.notebook.select(tab_id)
                self.track_page()
                break

    def add_character_sheet_form(self):
        frame = tk.Frame(self.notebook)
        self.notebook.add(frame, text="Character Sheet Form")

        # --- Scrollable canvas inside the frame ---
        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Define fields for form
        fields = {
            "Character Name": tk.StringVar(),
            "Class": tk.StringVar(),
            "Subclass": tk.StringVar(),
            "Level": tk.StringVar(),
            "Race": tk.StringVar(),
            "Base Strength": tk.StringVar(),
            "Added Strength": tk.StringVar(),
            "Base Dexterity": tk.StringVar(),
            "Added Dexterity": tk.StringVar(),
            "Base Constitution": tk.StringVar(),
            "Added Constitution": tk.StringVar(),
            "Base Intelligence": tk.StringVar(),
            "Added Intelligence": tk.StringVar(),
            "Base Wisdom": tk.StringVar(),
            "Added Wisdom": tk.StringVar(),
            "Base Charisma": tk.StringVar(),
            "Added Charisma": tk.StringVar(),
            "Rolled Health": tk.StringVar(),
            "Armor Worn": tk.StringVar(),
        }

        self.form_fields = fields
        row = 0

        # First section (Character Info)
        for label, var in fields.items():
            tk.Label(scrollable_frame, text=label).grid(row=row, column=0, sticky="e", padx=10, pady=4)
            tk.Entry(scrollable_frame, textvariable=var, width=30).grid(row=row, column=1, padx=10, pady=4)
            row += 1

        # Submit button
        ttk.Button(scrollable_frame, text="Generate Character Sheet", command=self.submit_character_form).grid(row=row, column=0, columnspan=2, pady=15)

    def convert_excel_to_image(self, excel_path):
        
        from openpyxl import load_workbook
        from PIL import Image, ImageDraw, ImageFont

        try:
           wb = load_workbook(excel_path, data_only=True)
           sheet = wb.active

           #Set dimensions and font
           cell_width = 140
           cell_height = 30
           font = ImageFont.load_default()

           max_row = sheet.max_row
           max_col = sheet.max_column

           #Create a blank white image
           img = Image.new("RGB", (cell_width * max_col, cell_height * max_row), "white")
           draw = ImageDraw.Draw(img)

           for row in range(1, max_row + 1):
               for col in range(1, max_col + 1):
                   cell = sheet.cell(row=row, column=col)
                   value = str(cell.value) if cell.value is not None else ""

                   x0 = (col - 1) * cell_width
                   y0 = (row - 1) * cell_height
                   x1 = x0 + cell_width
                   y1 = y0 + cell_height

                   draw.rectangle([x0, y0, x1, y1], outline="black")
                   draw.text((x0 + 5, y0 + 5), value, fill="black", font=font)

           png_path = excel_path.replace(".xlsx", ".png")
           img.save(png_path)
           self.display_image(img)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert Excel to image:\n\n{e}")

    def submit_character_form(self):
        from openpyxl import load_workbook

        import shutil

        input_path = "DND5e_Character_Sheet.xlsx"

        f = self.form_fields  # shortcut
        character_name = f["Character Name"].get().strip()
        if not character_name:
            messagebox.showwarning("Missing Data", "Please enter a character name.")
            return

        safe_name = "".join(c for c in character_name if c.isalnum() or c in (' ', '_')).rstrip()
        folder = os.path.dirname(os.path.abspath(input_path))
        output_path = os.path.join(folder, f"{safe_name}_CharacterSheet.xlsx")

        try:
            shutil.copy(input_path, output_path)

            wb = load_workbook(output_path)
        
            sheet0 = wb["Sheet0"]
            sheet1 = wb["Sheet1"]

            f = self.form_fields  # shortcut

            # Fill values into specified cells
            sheet0["B3"] = f["Character Name"].get()
            sheet0["D3"] = f["Class"].get() 
            sheet0["E3"] = f["Subclass"].get() 
            sheet0["C5"] = f["Level"].get()
            sheet0["F3"] = f["Race"].get()

            sheet1["B4"] = f["Base Strength"].get()
            sheet1["C4"] = f["Added Strength"].get()

            sheet1["B5"] = f["Base Dexterity"].get()
            sheet1["C5"] = f["Added Dexterity"].get()

            sheet1["B6"] = f["Base Constitution"].get()
            sheet1["C6"] = f["Added Constitution"].get()

            sheet1["B7"] = f["Base Intelligence"].get()
            sheet1["C7"] = f["Added Intelligence"].get()

            sheet1["B8"] = f["Base Wisdom"].get()
            sheet1["C8"] = f["Added Wisdom"].get()

            sheet1["B9"] = f["Base Charisma"].get()
            sheet1["C9"] = f["Added Charisma"].get()

            sheet1["U10"] = f["Rolled Health"].get()
            sheet1["B26"] = f["Armor Worn"].get()

            wb.save(output_path)


            self.convert_excel_to_image(output_path)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to fill sheet: {e}")

    def render_excel_as_image(excel_path, sheet_name="Sheet1", output_image="output.png"):
        wb = load_workbook(excel_path, data_only=True)
        sheet = wb[sheet_name]

        # Define basic styles
        cell_width = 140
        cell_height = 30
        font = ImageFont.load_default()

        max_row = sheet.max_row
        max_col = sheet.max_column

        img = Image.new("RGB", (cell_width * max_col, cell_height * max_row), "white")
        draw = ImageDraw.Draw(img)

        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                cell = sheet.cell(row=row, column=col)
                value = str(cell.value) if cell.value is not None else ""

                # Draw border
                x0 = (col - 1) * cell_width
                y0 = (row - 1) * cell_height
                x1 = x0 + cell_width
                y1 = y0 + cell_height
                draw.rectangle([x0, y0, x1, y1], outline="black")

                # Draw text
                draw.text((x0 + 5, y0 + 5), value, fill="black", font=font)

        img.save(output_image)
        print(f"Saved to {output_image}")


    def display_image(self, path):
        try:
            from PIL import Image

            pil_img = Image.open(path)
            img = pil_img  # Don't resize here—keep full resolution for scrolling
            photo = ImageTk.PhotoImage(img)

            frame = tk.Frame(self.notebook)

            # Create canvas and scrollbars
            canvas = tk.Canvas(frame, width=800, height=600)
            h_scroll = ttk.Scrollbar(frame, orient="horizontal", command=canvas.xview)
            v_scroll = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
            canvas.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

            # Create inner frame to hold the image
            image_frame = tk.Frame(canvas)
            image_label = tk.Label(image_frame, image=photo)
            image_label.image = photo
            image_label.pil_image = pil_img
            image_label.pack()

            # Create window on canvas
            canvas.create_window((0, 0), window=image_frame, anchor="nw")

            # Update scroll region after image loads
            image_frame.update_idletasks()
            canvas.config(scrollregion=(0, 0, photo.width(), photo.height()))

            # Layout
            canvas.grid(row=0, column=0, sticky="nsew")
            v_scroll.grid(row=0, column=1, sticky="ns")
            h_scroll.grid(row=1, column=0, sticky="ew")

            frame.grid_rowconfigure(0, weight=1)
            frame.grid_columnconfigure(0, weight=1)

            # Add Print Button
            ttk.Button(frame, text="🖨️ Print", command=self.print_current_editor_page).grid(row=2, column=0, columnspan=2, pady=5)

            self.notebook.add(frame, text="Generated Sheet")
            self.pages["Generated Sheet"] = image_label  # Keep track of the label (with .pil_image)
            self.notebook.select(frame)
            def _on_mousewheel(event):
                if event.state & 0x1:  # Shift is held → horizontal scroll
                    canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
                else:
                    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

            # Windows bindings
            canvas.bind_all("<MouseWheel>", _on_mousewheel)

        except Exception as e:
            messagebox.showerror("Error", f"Could not display image: {e}")

    def print_current_editor_page(self):
        import tempfile
        from PIL import Image
        import platform

        current_tab_id = self.notebook.select()
        current_tab_title = self.notebook.tab(current_tab_id, "text")
        widget = self.pages.get(current_tab_title)

        if isinstance(widget, ScrolledText):
            # --- TEXT CONTENT PRINT ---
            content = widget.get("1.0", tk.END).strip()
            if not content:
                messagebox.showinfo("Empty", "There is no content to print.")
                return

            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".txt", mode="w", encoding="utf-8") as tmp_file:
                    tmp_file.write(content)
                    tmp_path = tmp_file.name

                subprocess.run(["notepad.exe", "/p", tmp_path], check=True)

            except Exception as e:
                messagebox.showerror("Print Error", f"Failed to print text:\n{e}")

        elif isinstance(widget, tk.Label) and hasattr(widget, "image"):
            # --- IMAGE PRINT ---
            try:
                pil_image = widget.image._PhotoImage__photo  # Access Tkinter image object
                # Convert back to PIL image using the original PIL reference if you kept one
                if hasattr(widget, "pil_image"):
                    pil_image = widget.pil_image
                else:
                    messagebox.showerror("Print Error", "Original image not available for printing.")
                    return

                with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp_img_file:
                    tmp_img_path = tmp_img_file.name
                    pil_image.save(tmp_img_path)

                # Platform-specific print handling
                if platform.system() == "Windows":
                    os.startfile(tmp_img_path, "print")
                else:
                    subprocess.run(["lpr", tmp_img_path])  # macOS/Linux alternative

            except Exception as e:
                messagebox.showerror("Print Error", f"Failed to print image:\n{e}")

        else:
            messagebox.showinfo("Not Printable", "The current tab cannot be printed.")




if __name__ == "__main__":
    root = tk.Tk()
app = DiceScribeApp(root)
root.mainloop()

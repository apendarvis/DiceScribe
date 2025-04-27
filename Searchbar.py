import tkinter as tk
from tkinter import ttk

def build_search_bar(self):
       
        search_frame = ttk.Frame(self.editor_page)
        search_frame.pack(pady=5, fill='x')

        ttk.Label(search_frame, text="Search:").pack(side="left", padx=(10, 5))
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40)
        search_entry.pack(side="left", padx=(0, 10))
        search_entry.bind("<KeyRelease>", self.highlight_search)
        
def build_search_bar(self):
       
        search_frame = ttk.Frame(self.editor_page)
        search_frame.pack(pady=5, fill='x')

        ttk.Label(search_frame, text="Search:").pack(side="left", padx=(10, 5))
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40)
        search_entry.pack(side="left", padx=(0, 10))
        search_entry.bind("<KeyRelease>", self.highlight_search)
        
def highlight_search(self, event=None):
        # If the search bar is empty, remove all highlights and return
        if self.is_search_empty():
            current_tab_id = self.notebook.select()
            if not current_tab_id:
                return

            current_tab_title = self.notebook.tab(current_tab_id, "text")
            text_widget = self.pages.get(current_tab_title)

            if text_widget:
                text_widget.tag_remove("highlight", "1.0", tk.END)  # Remove all highlights
                return

        search_term = self.search_var.get()
        
        if len(search_term.strip()) < 4:
            # Also clear existing highlights
            current_tab_id = self.notebook.select()
            if current_tab_id:
                current_tab_title = self.notebook.tab(current_tab_id, "text")
                text_widget = self.pages.get(current_tab_title)
                if text_widget:
                    text_widget.tag_remove("highlight", "1.0", tk.END)
            return
        
        current_tab_id = self.notebook.select()
        if not current_tab_id:
            return

        current_tab_title = self.notebook.tab(current_tab_id, "text")
        text_widget = self.pages.get(current_tab_title)

        if not text_widget:
            print(f"No text widget found for tab {current_tab_title}.")
            return

        text_widget.tag_remove("highlight", "1.0", tk.END)
    
        startHighlight = "1.0"  #starts at line 1 character 0
        while True:
            #searches text widget 
            startHighlight = text_widget.search(search_term, startHighlight, nocase=True, stopindex=tk.END)
            if not startHighlight:
                break
            # finds start of found term and adds search term length to highlight full statement
            end = f"{startHighlight}+{len(search_term)}c"
            text_widget.tag_add("highlight", startHighlight, end)
            startHighlight = end

        # highlight style (yellow)
        text_widget.tag_config("highlight", background="yellow", foreground="black")
        
def is_search_empty(self):
        return self.search_var.get().strip() == ""
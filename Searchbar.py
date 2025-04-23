import tkinter as tk
from tkinter import ttk

def build_search_bar(self):
        # Create search bar frame
        search_frame = ttk.Frame(self.editor_page)
        search_frame.pack(pady=5, fill='x')

        ttk.Label(search_frame, text="Search:").pack(side="left", padx=(10, 5))
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40)
        search_entry.pack(side="left", padx=(0, 10))
        search_entry.bind("<KeyRelease>", self.highlight_search)
        
def highlight_search(self, event=None):
        search_term = self.search_var.get()
        if not search_term:
            return

    # Get the currently selected tab's title
        current_tab_id = self.notebook.select()
        if not current_tab_id:
            return

        current_tab_title = self.notebook.tab(current_tab_id, "text")

        # access text widget from self.pages using the tab's title
        text_widget = self.pages.get(current_tab_title)
        if not text_widget:
            print(f"No text widget found for tab {current_tab_title}.")
            return

         # remove old highlights
        text_widget.tag_remove("highlight", "1.0", tk.END)

        # search and highlight the search term in the text widget
        start_pos = "1.0" #method starts at line 1 character 0
        while True:
            #searches text widget from start to end
            start_pos = text_widget.search(search_term, start_pos, nocase=True, stopindex=tk.END)
            if not start_pos:
                break
            #finds start of found term and adds search term length to highlight full statement
            end_pos = f"{start_pos}+{len(search_term)}c"
            text_widget.tag_add("highlight", start_pos, end_pos)
            start_pos = end_pos

        # highlight style (yellow background)
        text_widget.tag_config("highlight", background="yellow", foreground="black")




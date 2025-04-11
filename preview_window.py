import tkinter as tk
from tkinter import ttk, messagebox, Toplevel
import os
from logger import logger
from xml_processor import XMLProcessor

class PreviewWindow:
    def __init__(self, parent, xml_files, processed_files, changes_made, backup_dir, output_dir):
        self.parent = parent
        self.xml_files = xml_files
        self.processed_files = processed_files
        self.changes_made = changes_made
        self.backup_dir = backup_dir
        self.output_dir = output_dir
        self.current_file_index = 0
        self.processor = XMLProcessor()
        
        # Create preview window
        self.window = Toplevel(parent)
        self.window.title("Preview Changes")
        self.window.geometry("1200x600")
        
        # Create a frame for the preview content
        preview_frame = ttk.Frame(self.window, padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Add a header
        header_frame = ttk.Frame(preview_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Create a frame for the text areas
        text_frame = ttk.Frame(preview_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create a common horizontal scrollbar at the bottom
        h_scrollbar = ttk.Scrollbar(text_frame, orient=tk.HORIZONTAL)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Original XML preview
        original_frame = ttk.LabelFrame(text_frame, text="Original XML")
        original_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        self.original_text = tk.Text(original_frame, wrap=tk.NONE, width=50, height=20)
        self.original_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add vertical scrollbar for original text (synchronized)
        original_y_scroll = ttk.Scrollbar(original_frame, orient=tk.VERTICAL, command=self.sync_scroll_y)
        original_y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.original_text.config(yscrollcommand=self.sync_scrollbar_y)
        
        # Modified XML preview
        modified_frame = ttk.LabelFrame(text_frame, text="Modified XML")
        modified_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        self.modified_text = tk.Text(modified_frame, wrap=tk.NONE, width=50, height=20)
        self.modified_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Add vertical scrollbar for modified text (synchronized)
        modified_y_scroll = ttk.Scrollbar(modified_frame, orient=tk.VERTICAL, command=self.sync_scroll_y)
        modified_y_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.modified_text.config(yscrollcommand=self.sync_scrollbar_y)
        
        # Configure horizontal scrollbar to control both text widgets
        h_scrollbar.config(command=self.sync_scroll_x)
        self.original_text.config(xscrollcommand=self.sync_scrollbar_x)
        self.modified_text.config(xscrollcommand=self.sync_scrollbar_x)
        
        # Bind additional events for better synchronization
        self.original_text.bind("<KeyRelease>", self.sync_from_original)
        self.modified_text.bind("<KeyRelease>", self.sync_from_modified)
        self.original_text.bind("<Button-1>", self.sync_from_original)
        self.modified_text.bind("<Button-1>", self.sync_from_modified)
        
        # Add pagination controls
        pagination_frame = ttk.Frame(preview_frame)
        pagination_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(pagination_frame, text="< Previous", command=self.prev_file).pack(side=tk.LEFT)
        self.file_label = ttk.Label(pagination_frame, text=f"File 1 of {len(self.processed_files)}")
        self.file_label.pack(side=tk.LEFT, padx=20)
        ttk.Button(pagination_frame, text="Next >", command=self.next_file).pack(side=tk.LEFT)
        
        # Add action buttons
        button_frame = ttk.Frame(preview_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(button_frame, text="History", command=self.show_history).pack(side=tk.LEFT)
        ttk.Button(button_frame, text="Cancel", command=self.window.destroy).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="Proceed to Process", command=self.confirm_process).pack(side=tk.RIGHT, padx=10)
        
        # Load the first file
        self.load_file_preview()
    
    def sync_scrollbar_x(self, *args):
        """Update both text widgets and scrollbar when one is scrolled horizontally"""
        self.original_text.xview_moveto(args[0])
        self.modified_text.xview_moveto(args[0])
        return args
    
    def sync_scrollbar_y(self, *args):
        """Update both text widgets when one is scrolled vertically"""
        self.original_text.yview_moveto(args[0])
        self.modified_text.yview_moveto(args[0])
        return args
    
    def sync_scroll_x(self, *args):
        """Synchronize horizontal scrolling between the two text widgets"""
        self.original_text.xview(*args)
        self.modified_text.xview(*args)
    
    def sync_scroll_y(self, *args):
        """Synchronize vertical scrolling between the two text widgets"""
        self.original_text.yview(*args)
        self.modified_text.yview(*args)
    
    def sync_from_original(self, event=None):
        """Synchronize modified text to match original text position"""
        self.modified_text.yview_moveto(self.original_text.yview()[0])
        self.modified_text.xview_moveto(self.original_text.xview()[0])
    
    def sync_from_modified(self, event=None):
        """Synchronize original text to match modified text position"""
        self.original_text.yview_moveto(self.modified_text.yview()[0])
        self.original_text.xview_moveto(self.modified_text.xview()[0])
    
    def load_file_preview(self):
        """Load the current file into the preview"""
        if not self.processed_files or self.current_file_index >= len(self.processed_files):
            return
            
        # Get the current processed file
        processed_file = self.processed_files[self.current_file_index]
        original_file = self.xml_files[self.current_file_index]
        
        # Update file label
        self.file_label.config(text=f"File {self.current_file_index + 1} of {len(self.processed_files)}")
        
        # Load original file content
        try:
            with open(original_file, 'r') as f:
                original_content = f.read()
                
            # Load processed file content
            with open(processed_file, 'r') as f:
                processed_content = f.read()
                
            # Update text widgets
            self.original_text.delete(1.0, tk.END)
            self.original_text.insert(tk.END, original_content)
            
            self.modified_text.delete(1.0, tk.END)
            self.modified_text.insert(tk.END, processed_content)
            
            # Highlight changes
            self.highlight_file_changes(original_file)
            
        except Exception as e:
            logger.error(f"Error loading file preview: {str(e)}", exc_info=True)
            messagebox.showerror("Error", f"Failed to load file preview: {str(e)}")
    
    def highlight_file_changes(self, original_file):
        """Highlight the changes in the current file"""
        try:
            # Configure tags for highlighting
            self.original_text.tag_configure("change", background="lightgreen")
            self.modified_text.tag_configure("change", background="yellow")
            self.original_text.tag_configure("line", background="#f0f0f0")  # Light gray for changed lines
            self.modified_text.tag_configure("line", background="#f0f0f0")  # Light gray for changed lines
            
            # Get changes for this file
            changes = self.changes_made.get(original_file, [])
            
            for change in changes:
                # Find the locations of the changes in both files
                original_pattern = f'startEffectiveDate="{change["old_value"]}"'
                modified_pattern = f'startEffectiveDate="{change["new_value"]}"'
                expression_pattern = f'expression="{change["expression"]}"'
                
                # Search in original text
                start_pos = "1.0"
                while True:
                    # Find the expression
                    expr_pos = self.original_text.search(expression_pattern, start_pos, tk.END)
                    if not expr_pos:
                        break
                    
                    # Get the line number
                    line_num = self.original_text.index(expr_pos).split('.')[0]
                    
                    # Highlight the whole line
                    line_start = f"{line_num}.0"
                    line_end = f"{line_num}.end"
                    self.original_text.tag_add("line", line_start, line_end)
                    
                    # Next, find the date pattern near this line
                    search_start = line_start
                    search_end = f"{int(line_num) + 5}.end"  # Look a few lines ahead
                    
                    try:
                        # Search for the date attribute
                        date_pos = self.original_text.search(original_pattern, search_start, search_end)
                        if date_pos:
                            # Highlight the specific date
                            date_end = f"{date_pos}+{len(original_pattern)}c"
                            self.original_text.tag_add("change", date_pos, date_end)
                            
                            # Also highlight the line containing the date
                            date_line = self.original_text.index(date_pos).split('.')[0]
                            self.original_text.tag_add("line", f"{date_line}.0", f"{date_line}.end")
                    except Exception as e:
                        logger.error(f"Error highlighting original text: {e}")
                    
                    # Move to next occurrence
                    start_pos = f"{expr_pos}+{len(expression_pattern)}c"
                
                # Search in modified text
                start_pos = "1.0"
                while True:
                    # Find the expression
                    expr_pos = self.modified_text.search(expression_pattern, start_pos, tk.END)
                    if not expr_pos:
                        break
                    
                    # Get the line number
                    line_num = self.modified_text.index(expr_pos).split('.')[0]
                    
                    # Highlight the whole line
                    line_start = f"{line_num}.0"
                    line_end = f"{line_num}.end"
                    self.modified_text.tag_add("line", line_start, line_end)
                    
                    # Next, find the date pattern near this line
                    search_start = line_start
                    search_end = f"{int(line_num) + 5}.end"  # Look a few lines ahead
                    
                    try:
                        # Search for the date attribute
                        date_pos = self.modified_text.search(modified_pattern, search_start, search_end)
                        if date_pos:
                            # Highlight the specific date
                            date_end = f"{date_pos}+{len(modified_pattern)}c"
                            self.modified_text.tag_add("change", date_pos, date_end)
                            
                            # Also highlight the line containing the date
                            date_line = self.modified_text.index(date_pos).split('.')[0]
                            self.modified_text.tag_add("line", f"{date_line}.0", f"{date_line}.end")
                    except Exception as e:
                        logger.error(f"Error highlighting modified text: {e}")
                    
                    # Move to next occurrence
                    start_pos = f"{expr_pos}+{len(expression_pattern)}c"
        
        except Exception as e:
            logger.error(f"Error in highlight_file_changes: {str(e)}", exc_info=True)
    
    def prev_file(self):
        """Show the previous file in the preview"""
        if self.current_file_index > 0:
            self.current_file_index -= 1
            self.load_file_preview()
        else:
            logger.debug("Already at the first file")
    
    def next_file(self):
        """Show the next file in the preview"""
        if self.current_file_index < len(self.processed_files) - 1:
            self.current_file_index += 1
            self.load_file_preview()
        else:
            logger.debug("Already at the last file")
    
    def show_history(self):
        """Show the history of changes"""
        history_window = Toplevel(self.window)
        history_window.title("Change History")
        history_window.geometry("600x400")
        
        # Create a frame for the history content
        history_frame = ttk.Frame(history_window, padding="10")
        history_frame.pack(fill=tk.BOTH, expand=True)
        
        # Add a header
        ttk.Label(history_frame, text="Changes to be Applied", font=("Arial", 12, "bold")).pack(anchor=tk.W, pady=(0, 10))
        
        # Create a text widget to display the history
        history_text = tk.Text(history_frame, wrap=tk.WORD)
        history_text.pack(fill=tk.BOTH, expand=True)
        
        # Add a scrollbar
        scrollbar = ttk.Scrollbar(history_text, orient=tk.VERTICAL, command=history_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        history_text.config(yscrollcommand=scrollbar.set)
        
        # Populate the history text
        total_changes = 0
        for i, original_file in enumerate(self.xml_files):
            if original_file in self.changes_made and self.changes_made[original_file]:
                file_name = os.path.basename(original_file)
                history_text.insert(tk.END, f"File: {file_name}\n", "file_header")
                
                for change in self.changes_made[original_file]:
                    history_text.insert(tk.END, f"  Expression: {change['expression']}\n")
                    history_text.insert(tk.END, f"  Changed {change['attribute']} from '{change['old_value']}' to '{change['new_value']}'\n")
                    history_text.insert(tk.END, f"  (Based on match with BaseFilename: {file_name})\n\n")
                    total_changes += 1
        
        # Add a summary at the top
        history_text.insert("1.0", f"Total files to be modified: {len(self.processed_files)}\n")
        history_text.insert("2.0", f"Total changes to be applied: {total_changes}\n\n")
        
        # Configure tags
        history_text.tag_configure("file_header", font=("Arial", 10, "bold"))
        
        # Make the text read-only
        history_text.config(state=tk.DISABLED)
    
    def confirm_process(self):
        """Confirm and proceed with processing the files"""
        # Ask for confirmation
        if not messagebox.askyesno("Confirm", "Are you sure you want to apply these changes?"):
            return
        
        try:
            # Apply the changes
            success, message = self.processor.apply_changes(
                self.xml_files, 
                self.processed_files, 
                self.changes_made, 
                self.backup_dir, 
                self.output_dir
            )
            
            if success:
                messagebox.showinfo("Success", message)
            else:
                messagebox.showerror("Error", message)
                
            # Close the preview window
            self.window.destroy()
            
        except Exception as e:
            logger.error(f"Error during final processing: {str(e)}", exc_info=True)
            messagebox.showerror("Error", f"An error occurred during processing: {str(e)}")
                    
#!/usr/bin/env python3
"""
EDI XML Data Extractor GUI

A user-friendly GUI tool to extract delivery ID, product service ID, and serial number
from XML-formatted EDI files with file selection dialogs.
"""

import os
import re
import csv
import sys
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext, messagebox
from datetime import datetime
from typing import Dict, List, Optional


class EDIExtractor:
    """Extract data from XML-formatted EDI files."""

    @staticmethod
    def extract_from_file(filename: str) -> List[Dict[str, str]]:
        """
        Extract data from an EDI file.
        
        Args:
            filename: Path to the EDI file
            
        Returns:
            List of dictionaries containing extracted data
        """
        try:
            with open(filename, 'r', encoding='utf-8') as file:
                content = file.read()
            return EDIExtractor.extract_data(content)
        except Exception as e:
            raise Exception(f"Error processing file {filename}: {e}")
            
    @staticmethod
    def extract_data(content: str) -> List[Dict[str, str]]:
        """
        Extract data from EDI content.
        
        Args:
            content: String content of the EDI file
            
        Returns:
            List of dictionaries containing extracted data
        """
        results = []
        
        # Find all ShipConfirmLine blocks
        ship_confirm_pattern = re.compile(r'<ShipConfirmLine>([\s\S]*?)</ShipConfirmLine>')
        for match in ship_confirm_pattern.finditer(content):
            line_block = match.group(1)
            
            # Extract delivery_id
            delivery_id_match = re.search(r'<delivery_id>(.*?)</delivery_id>', line_block)
            delivery_id = delivery_id_match.group(1) if delivery_id_match else ""
            
            # Extract ProductServiceId
            product_id_match = re.search(r'<ProductServiceId>(.*?)</ProductServiceId>', line_block)
            product_service_id = product_id_match.group(1) if product_id_match else ""
            
            # Find ShipConfirmSerials blocks
            serials_pattern = re.compile(r'<ShipConfirmSerials>([\s\S]*?)</ShipConfirmSerials>')
            serials_found = False
            
            for serials_match in serials_pattern.finditer(line_block):
                serials_found = True
                serials_block = serials_match.group(1)
                
                # Extract serial_number if it exists
                serial_match = re.search(r'<serial_number>(.*?)</serial_number>', serials_block)
                serial_number = serial_match.group(1) if serial_match else ""
                
                results.append({
                    'delivery_id': delivery_id,
                    'product_service_id': product_service_id,
                    'serial_number': serial_number
                })
            
            # If no serials block or no serial number found, still record the product
            if not serials_found:
                results.append({
                    'delivery_id': delivery_id,
                    'product_service_id': product_service_id,
                    'serial_number': ""
                })
                
        return results
    
    @staticmethod
    def save_to_csv(results: List[Dict[str, str]], output_file: str) -> None:
        """
        Save results to a CSV file.
        
        Args:
            results: List of dictionaries containing extracted data
            output_file: Path to output file
        """
        fieldnames = ['delivery_id', 'product_service_id', 'serial_number']
        
        with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter=',', 
                                   quotechar='"', quoting=csv.QUOTE_MINIMAL)
            writer.writeheader()
            writer.writerows(results)


class EDIExtractorApp:
    """GUI application for the EDI Extractor."""
    
    def __init__(self, root):
        """Initialize the GUI application."""
        self.root = root
        self.root.title("EDI XML Data Extractor")
        self.root.geometry("900x600")
        self.root.minsize(800, 500)
        
        # Set application icon if available
        try:
            self.root.iconbitmap("edi_icon.ico")  # You can add your own icon file
        except:
            pass
        
        # File path variables
        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()
        
        # Default output file path
        default_filename = f"edi_extract_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        default_output_path = os.path.join(os.path.expanduser("~"), "Documents", default_filename)
        self.output_file_path.set(default_output_path)
        
        # Create the main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create the input section
        input_frame = ttk.LabelFrame(main_frame, text="Input EDI File", padding="10")
        input_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Entry(input_frame, textvariable=self.input_file_path, width=70).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(input_frame, text="Browse...", command=self.browse_input_file).pack(side=tk.RIGHT, padx=5)
        
        # Create the output section
        output_frame = ttk.LabelFrame(main_frame, text="Output CSV File", padding="10")
        output_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Entry(output_frame, textvariable=self.output_file_path, width=70).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(output_frame, text="Browse...", command=self.browse_output_file).pack(side=tk.RIGHT, padx=5)
        
        # Create the action buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.extract_button = ttk.Button(button_frame, text="Extract Data", command=self.extract_data, state=tk.DISABLED)
        self.extract_button.pack(side=tk.LEFT, padx=5)
        
        self.save_button = ttk.Button(button_frame, text="Save to CSV", command=self.save_to_csv, state=tk.DISABLED)
        self.save_button.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(button_frame, text="Exit", command=self.root.destroy).pack(side=tk.RIGHT, padx=5)
        
        # Create the results area with tabs
        self.tabControl = ttk.Notebook(main_frame)
        self.tabControl.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Table view tab
        self.table_tab = ttk.Frame(self.tabControl)
        self.tabControl.add(self.table_tab, text="Table View")
        
        # Create the table
        table_frame = ttk.Frame(self.table_tab)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Scrollbars
        y_scrollbar = ttk.Scrollbar(table_frame)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        x_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Treeview for the table
        self.tree = ttk.Treeview(table_frame, 
                                yscrollcommand=y_scrollbar.set,
                                xscrollcommand=x_scrollbar.set)
        
        y_scrollbar.config(command=self.tree.yview)
        x_scrollbar.config(command=self.tree.xview)
        
        # Define columns
        self.tree["columns"] = ("delivery_id", "product_service_id", "serial_number")
        self.tree.column("#0", width=0, stretch=tk.NO)  # Hide the first column
        self.tree.column("delivery_id", width=200, anchor=tk.W)
        self.tree.column("product_service_id", width=250, anchor=tk.W)
        self.tree.column("serial_number", width=250, anchor=tk.W)
        
        # Define column headings
        self.tree.heading("#0", text="", anchor=tk.W)
        self.tree.heading("delivery_id", text="Delivery ID", anchor=tk.W)
        self.tree.heading("product_service_id", text="Product Service ID", anchor=tk.W)
        self.tree.heading("serial_number", text="Serial Number", anchor=tk.W)
        
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Raw data tab
        self.raw_tab = ttk.Frame(self.tabControl)
        self.tabControl.add(self.raw_tab, text="Raw Data")
        
        # Create a text area for raw data
        self.raw_text = scrolledtext.ScrolledText(self.raw_tab)
        self.raw_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready. Please select an input file.")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Store the extracted results
        self.results = []
        
        # Bind file path changes to update button states
        self.input_file_path.trace_add("write", self.update_button_states)
        
        # Center the window
        self.center_window()
    
    def center_window(self):
        """Center the application window on the screen."""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def update_button_states(self, *args):
        """Update the state of buttons based on inputs."""
        if self.input_file_path.get():
            self.extract_button.config(state=tk.NORMAL)
        else:
            self.extract_button.config(state=tk.DISABLED)
        
        if self.results:
            self.save_button.config(state=tk.NORMAL)
        else:
            self.save_button.config(state=tk.DISABLED)
    
    def browse_input_file(self):
        """Open file dialog to select input EDI file."""
        file_path = filedialog.askopenfilename(
            title="Select EDI File",
            filetypes=[("Text/XML files", "*.txt *.xml"), ("All files", "*.*")]
        )
        if file_path:
            self.input_file_path.set(file_path)
            self.status_var.set(f"Selected input file: {os.path.basename(file_path)}")
    
    def browse_output_file(self):
        """Open file dialog to select output CSV file."""
        file_path = filedialog.asksaveasfilename(
            title="Save CSV File",
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            initialfile=f"edi_extract_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        if file_path:
            self.output_file_path.set(file_path)
            self.status_var.set(f"Selected output file: {os.path.basename(file_path)}")
    
    def extract_data(self):
        """Extract data from the selected EDI file."""
        input_file = self.input_file_path.get()
        
        if not input_file:
            messagebox.showerror("Error", "Please select an input file.")
            return
        
        if not os.path.exists(input_file):
            messagebox.showerror("Error", f"File not found: {input_file}")
            return
        
        try:
            self.status_var.set("Extracting data...")
            self.root.update()
            
            # Clear previous results
            self.clear_results()
            
            # Extract data
            self.results = EDIExtractor.extract_from_file(input_file)
            
            if not self.results:
                messagebox.showinfo("Info", "No data found in the selected file.")
                self.status_var.set("No data found in the selected file.")
                return
            
            # Update the table view
            for i, row in enumerate(self.results):
                self.tree.insert("", tk.END, iid=i, values=(
                    row["delivery_id"],
                    row["product_service_id"],
                    row["serial_number"]
                ))
            
            # Update the raw view
            self.raw_text.delete(1.0, tk.END)
            for row in self.results:
                self.raw_text.insert(tk.END, 
                                    f"Delivery ID: {row['delivery_id']}\n"
                                    f"Product Service ID: {row['product_service_id']}\n"
                                    f"Serial Number: {row['serial_number']}\n\n")
            
            # Update button states and status
            self.update_button_states()
            self.status_var.set(f"Extracted {len(self.results)} records from {os.path.basename(input_file)}")
            
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.status_var.set(f"Error: {str(e)}")
    
    def save_to_csv(self):
        """Save extracted data to a CSV file."""
        if not self.results:
            messagebox.showerror("Error", "No data to save. Please extract data first.")
            return
        
        output_file = self.output_file_path.get()
        
        if not output_file:
            messagebox.showerror("Error", "Please select an output file location.")
            return
        
        try:
            # Make sure output directory exists
            output_dir = os.path.dirname(output_file)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
            
            # Save the data
            EDIExtractor.save_to_csv(self.results, output_file)
            
            # Update status
            self.status_var.set(f"Data saved to {os.path.basename(output_file)}")
            
            # Ask if user wants to open the file
            if messagebox.askyesno("Success", 
                                  f"Data saved to {output_file}.\n\nWould you like to open this file now?"):
                self.open_csv_file(output_file)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {str(e)}")
            self.status_var.set(f"Error: {str(e)}")
    
    def open_csv_file(self, file_path):
        """Open the CSV file with the default application."""
        try:
            if sys.platform == 'win32':
                os.startfile(file_path)
            elif sys.platform == 'darwin':  # macOS
                import subprocess
                subprocess.call(['open', file_path])
            else:  # Linux
                import subprocess
                subprocess.call(['xdg-open', file_path])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file: {str(e)}")
    
    def clear_results(self):
        """Clear all result displays."""
        # Clear the table
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Clear the raw text
        self.raw_text.delete(1.0, tk.END)
        
        # Clear results data
        self.results = []
        
        # Update button states
        self.update_button_states()


def main():
    """Main entry point for the application."""
    root = tk.Tk()
    app = EDIExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

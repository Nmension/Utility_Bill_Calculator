import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import numpy as np
import sv_ttk
import os
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, PageTemplate, Table, TableStyle, Frame
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus.paragraph import Paragraph
import sys


class UtilityBillCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("Calculateur de décompte énergétique")
        # Make application full screen by default
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight() -75
        self.root.geometry(f"{screen_width}x{screen_height}+0+0")
        
        # Alternative full screen option
        # self.root.attributes('-fullscreen', True)  # true fullscreen with no window decorations
    
        # Configure styles for larger text
        style = ttk.Style()
        style.configure("TLabel", font=('Arial', 16))
        style.configure("TButton", font=('Arial', 16))
        style.configure("TNotebook.Tab", font=('Arial', 16))

        
        # Data structure to hold all values
        self.data = {
            'columns': ['B1', 'B2', 'B3', 'B4', 'B5', 'R', 'Général'],
            'sections': ['Electricité', 'Eau', 'Consommation', 'A PAYER'],
            'rows': {
                'Electricité': ['Ancien Index', 'Nouvel Index'],
                'Eau': ['Ancien Index', 'Nouvel Index'],
                'Consommation': ['Electricité (en kWh)', 'Eau (en m3)'],
                'A PAYER': ['Electricité', 'Eau', 'Montant Total']
            }
        }
        
        # Initialize data matrix
        self.values = {}
        for section in self.data['sections']:
            self.values[section] = {}
            for row in self.data['rows'][section]:
                self.values[section][row] = {}
                for col in self.data['columns']:
                    self.values[section][row][col] = 0
        
        # Special values
        self.price_per_unit = {
            'Electricité': 0,
            'Eau': 0
        }
        
        self.total_net = {
            'Electricité': 0,
            'Eau': 0
        }
        #date variable 
        self.str_date:str

        # Create main container
        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create notebook for input sections
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Create input tabs for each column
        self.input_frames = {}
        for col in self.data['columns']:
            self.input_frames[col] = self.create_input_tab(col)
        
        # Create final values input tab
        self.create_final_values_tab()
        
        # Create results tab
        self.create_results_tab()
        
        # Navigation buttons with larger font
        self.nav_frame = ttk.Frame(self.main_frame, padding=10)
        self.nav_frame.pack(fill=tk.X, pady=15)
        
        button_font = ('Arial', 16)
        # Configure button style
        style.configure("Large.TButton", font=button_font)
        
        ttk.Button(self.nav_frame, text="Précédent", command=self.previous_tab, 
                style="Large.TButton").pack(side=tk.LEFT, padx=10)
        ttk.Button(self.nav_frame, text="Suivant", command=self.next_tab, 
                style="Large.TButton").pack(side=tk.LEFT, padx=10)
        ttk.Button(self.nav_frame, text="Calculer", command=self.calculate, 
                style="Large.TButton").pack(side=tk.RIGHT, padx=10)
        ttk.Button(self.nav_frame, text="Enregistrer les résultats", command=self.save_results, 
                style="Large.TButton").pack(side=tk.RIGHT, padx=10)
        
        
    
    def create_input_tab(self, column):
        frame = ttk.Frame(self.notebook, padding=30)  # Increased padding to 20
        self.notebook.add(frame, text=column)
        
        # Set a larger font for all elements
        large_font = ('Arial', 16)
        header_font = ('Arial', 16, 'bold')
        
        # Electricity section
        ttk.Label(frame, text="Electricité:", font=header_font).grid(column=0, row=0, columnspan=2, sticky=tk.W, pady=(15, 10))
        
        ttk.Label(frame, text="Ancien Index", font=large_font).grid(column=0, row=1, sticky=tk.W, padx=10, pady=5)
        ancien_elec = ttk.Entry(frame, width=20, font=large_font)  # Increased width from 15 to 20
        ancien_elec.grid(column=1, row=1, padx=10, pady=5)
        
        ttk.Label(frame, text="Nouvel Index", font=large_font).grid(column=0, row=2, sticky=tk.W, padx=10, pady=5)
        nouvel_elec = ttk.Entry(frame, width=20, font=large_font)  # Increased width from 15 to 20
        nouvel_elec.grid(column=1, row=2, padx=10, pady=5)
        
        # Water section
        ttk.Label(frame, text="Eau:", font=header_font).grid(column=0, row=3, columnspan=2, sticky=tk.W, pady=(15, 10))
        
        ttk.Label(frame, text="Ancien Index", font=large_font).grid(column=0, row=4, sticky=tk.W, padx=10, pady=5)
        ancien_eau = ttk.Entry(frame, width=20, font=large_font)  # Increased width from 15 to 20
        ancien_eau.grid(column=1, row=4, padx=10, pady=5)
        
        ttk.Label(frame, text="Nouvel Index", font=large_font).grid(column=0, row=5, sticky=tk.W, padx=10, pady=5)
        nouvel_eau = ttk.Entry(frame, width=20, font=large_font)  # Increased width from 15 to 20
        nouvel_eau.grid(column=1, row=5, padx=10, pady=5)
        
        # Store references to entry widgets
        entries = {
            'Electricité': {
                'Ancien Index': ancien_elec,
                'Nouvel Index': nouvel_elec
            },
            'Eau': {
                'Ancien Index': ancien_eau,
                'Nouvel Index': nouvel_eau
            }
        }
        
        return entries
    
    def create_final_values_tab(self):
        frame = ttk.Frame(self.notebook, padding=20)  #padding
        self.notebook.add(frame, text="Montants Jirama")
        
        # Set a larger font for all elements
        large_font = ('Arial', 14)
        header_font = ('Arial', 14, 'bold')
        
        ttk.Label(frame, text="Total Net Jirama:", font=header_font).grid(column=0, row=0, columnspan=2, sticky=tk.W, pady=(15, 10))
        
        ttk.Label(frame, text="Electricité", font=large_font).grid(column=0, row=1, sticky=tk.W, padx=10, pady=5)
        self.total_elec_entry = ttk.Entry(frame, width=20, font=large_font)  #width elec
        self.total_elec_entry.grid(column=1, row=1, padx=10, pady=5)
        
        ttk.Label(frame, text="Eau", font=large_font).grid(column=0, row=2, sticky=tk.W, padx=10, pady=5)
        self.total_eau_entry = ttk.Entry(frame, width=20, font=large_font)  #width eau
        self.total_eau_entry.grid(column=1, row=2, padx=10, pady=5)

        ttk.Label(frame, text="Dates ou période (optionnelle):", font=header_font).grid(column=0, row=4, columnspan=2, sticky=tk.W, pady=(15, 10))
        self.str_date_entry = ttk.Entry(frame, width=20, font=large_font)  #width date
        self.str_date_entry.grid(column=1, row=5, padx=10, pady=5)
    
    def create_results_tab(self):
        frame = ttk.Frame(self.notebook, padding=20)  #padding res
        self.notebook.add(frame, text="Résultats")
        
        # Create treeview to display results in a table
        columns = ['Category'] + self.data['columns'] + ['Somme(Villas+R)', 'Prix unitaire', 'Total Net Jirama']
        
        # Configure style for the treeview
        style = ttk.Style()
        style.configure("Treeview", font=('Arial', 14))  # Set font size for treeview items
        style.configure("Treeview.Heading", font=('Arial', 14))  # Set font size for treeview headings
        
        self.tree = ttk.Treeview(frame, columns=columns, show='headings', height=15)  # Adjusted height for larger font
        
        # Configure columns
        self.tree.column('Category', width=180, anchor='w')  # Increased width
        for col in columns[1:]:
            self.tree.column(col, width=120, anchor='center')  # Increased width
        
        # Add headings
        for col in columns:
            self.tree.heading(col, text=col)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack widgets
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def next_tab(self):
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab < self.notebook.index('end') - 1:
            self.notebook.select(current_tab + 1)
    
    def previous_tab(self):
        current_tab = self.notebook.index(self.notebook.select())
        if current_tab > 0:
            self.notebook.select(current_tab - 1)
    
    def get_input_values(self):
        # Get values from input tabs
        for col in self.data['columns']:
            for section in ['Electricité', 'Eau']:
                for row in self.data['rows'][section]:
                    entry = self.input_frames[col][section][row]
                    try:
                        value = float(entry.get()) if entry.get() else 0
                        self.values[section][row][col] = value
                    except ValueError:
                        messagebox.showerror("Donnée entrée érronée", f"Valeur invalide au niveau de: {col} - {section} - {row}")
                        return False
        
        # Get final values
        try:
            self.total_net['Electricité'] = float(self.total_elec_entry.get().replace(',', '.')) if self.total_elec_entry.get() else 0
            self.total_net['Eau'] = float(self.total_eau_entry.get().replace(',', '.')) if self.total_eau_entry.get() else 0
            self.str_date = self.str_date_entry.get() if self.str_date_entry.get() else " "
        except ValueError:
            messagebox.showerror("Donnée entrée érronée", "Veuillez entrer le bon type de valeur")
            return False
        
        return True
    
    def calculate(self):
        if not self.get_input_values():
            return
        
        # Calculate consumption (Nouvel Index - Ancien Index)
        for col in self.data['columns']:
            # Electricity consumption
            self.values['Consommation']['Electricité (en kWh)'][col] = (
                self.values['Electricité']['Nouvel Index'][col] - 
                self.values['Electricité']['Ancien Index'][col]
            )
            
            # Water consumption
            if col != 'R':
                self.values['Consommation']['Eau (en m3)'][col] = (
                    (self.values['Eau']['Nouvel Index'][col] - 
                    self.values['Eau']['Ancien Index'][col])) /1000
                
            else: 
                self.values['Consommation']['Eau (en m3)'][col] = (
                    (self.values['Eau']['Nouvel Index'][col] - 
                    self.values['Eau']['Ancien Index'][col]) 
                )
        
        # Calculate sums for I10 and I11 (sum of consumption except 'Général')
        sum_elec = sum(self.values['Consommation']['Electricité (en kWh)'][col] for col in self.data['columns'] if col != 'Général')
        sum_eau = sum(self.values['Consommation']['Eau (en m3)'][col] for col in self.data['columns'] if col != 'Général')
        # Calculate unit prices (K14 = L14/I10, K15 = L15/I11)
        if sum_elec > 0:
            self.price_per_unit['Electricité'] = round(self.total_net['Electricité'] / sum_elec, 2)
        else:
            self.price_per_unit['Electricité'] = 0
            
        if sum_eau > 0:
            self.price_per_unit['Eau'] = round(self.total_net['Eau'] / sum_eau, 2)
        else:
            self.price_per_unit['Eau'] = 0
        
        # Calculate amounts to pay
        for col in self.data['columns']:

            # Electricity payment
            self.values['A PAYER']['Electricité'][col] = (
                    self.values['Consommation']['Electricité (en kWh)'][col] * 
                    self.price_per_unit['Electricité']
            )
                
            # Water payment
            self.values['A PAYER']['Eau'][col] = (
                    self.values['Consommation']['Eau (en m3)'][col] * 
                    self.price_per_unit['Eau'] 
            ) 
            # Total amount
            self.values['A PAYER']['Montant Total'][col] = (
                    self.values['A PAYER']['Electricité'][col] + 
                    self.values['A PAYER']['Eau'][col]
            )
        
        # Update results display
        self.update_results()
    
    def update_results(self):
        # Clear previous results
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add date row
        self.tree.insert('', 'end', values=["Dates ou période:", "", "", "", "", f"{self.str_date}", "", "", "", "", "", ""])
        
        # Calculate sums for each section
        sums = {}
        for section in self.data['sections']:
            sums[section] = {}
            for row in self.data['rows'][section]:
                regular_sum = sum(self.values[section][row][col] for col in self.data['columns'] if col != 'Général')
                sums[section][row] = {'sum': regular_sum}
        
        # Add data rows by section
        for section in self.data['sections']:   
            self.tree.insert('', 'end', values=[section] + [""] * 10)
            
            for row in self.data['rows'][section]:
                values = [row]
                
                # Add values for each column
                for col in self.data['columns']:
                    #add empty strings instead of values for the General rows in the A PAYER section
                    if col == 'Général':
                        if section == 'A PAYER':
                            values.append(f"")
                        else:
                            values.append(f"{self.values[section][row][col]}".rstrip('0').rstrip('.'))
                    else:                    
                        values.append(f"{self.values[section][row][col]:.3f}".rstrip('0').rstrip('.') if isinstance(self.values[section][row][col], float) else self.values[section][row][col])
                
                
                # Add sum, difference, etc.
                if section == 'A PAYER':
                    # For A PAYER rows, add Ar to the values
                    values_with_ar = [('{:_}'.format(round(float(value), 2))).replace('_', ' ') + " Ar" if (i > 0 and i < 7) else value for i, value in enumerate(values)]
                    
                    # Add empty Somme(Villas+R) in the cells
                    values_with_ar.append(f"")
                    
                    
                    
                    # Add Prix unitaire (Ar)
                    if row in ['Electricité', 'Eau']:
                        values_with_ar.append(f"{self.price_per_unit[row]} Ar".rstrip('0').rstrip('.'))
                    else:
                        values_with_ar.append(f"")
                    
                    # Add Total Net Jirama (Ar)
                    if row in ['Electricité', 'Eau']:
                        values_with_ar.append('{:_}'.format(float(f"{self.total_net[row]:.2f}")).replace('_', ' ') + " Ar")
                    else:
                        total = self.total_net['Electricité'] + self.total_net['Eau']
                        values_with_ar.append(('{:_}'.format(float(f"{total:.2f}"))).replace('_', ' ') + " Ar")
                    
                    self.tree.insert('', 'end', values=values_with_ar)
                
                else:
                    # Add Somme(Villas+R) for other sections
                    if section == self.data['sections'][2]:
                        regular_sum = sum(self.values[section][row][col] for col in self.data['columns'] if col != 'Général')
                        values.append(f"{regular_sum:.2f}".rstrip('0').rstrip('.'))
                    
                    # Default empty values for other columns
                    values.extend([""] * 2)
                    
                    self.tree.insert('', 'end', values=values)
    
    def save_results(self):
        file_types = [('PDF files', '*.pdf'), ('Excel files', '*.xlsx')]
        file_path = filedialog.asksaveasfilename(filetypes=file_types, defaultextension=".pdf")
        
        if not file_path:
            return
        
        if file_path.endswith('.pdf'):
            self.save_as_pdf(file_path)
        elif file_path.endswith('.xlsx'):
            self.save_as_excel(file_path)
        
    def save_as_pdf(self, file_path):
        # Create data for PDF
        data = []
        headers = ['Category'] + self.data['columns'] + ['Somme(Villas+R)', 'Prix unitaire', 'Total Net Jirama']
        data.append(headers)
        
        # Add date row
        data.append(["Dates ou période:", "", "", "", "", "", "", f"{self.str_date}", "", "", ""])
        
        # Add data by section
        for section in self.data['sections']:
            # Add section header if needed
        
            data.append([section] + [""] * 10)
            
            for row in self.data['rows'][section]:
                values = [row]
                
                # Add values for each column
                for col in self.data['columns']:
                    #add empty strings instead of values for the General rows in the A PAYER section
                    if col == 'Général':
                        if section == 'A PAYER':
                            values.append(f"")
                        else:
                            values.append(f"{self.values[section][row][col]}".rstrip('0').rstrip('.'))
                    else:                    
                        values.append(f"{self.values[section][row][col]:.3f}".rstrip('0').rstrip('.') if isinstance(self.values[section][row][col], float) else self.values[section][row][col])
                
                
                # Add sum, difference, etc.
                if section == 'A PAYER':
                    # For A PAYER rows, add Ar to the values
                    values_with_ar = [('{:_}'.format(round(float(value), 2))).replace('_', ' ') + " Ar" if (i > 0 and i < 7) else value for i, value in enumerate(values)]
                    
                    # Add empty Somme(Villas+R) in the cells
                    values_with_ar.append(f"")
                    
                    
                    
                    # Add Prix unitaire (Ar)
                    if row in ['Electricité', 'Eau']:
                        values_with_ar.append(f"{self.price_per_unit[row]} Ar".rstrip('0').rstrip('.'))
                    else:
                        values_with_ar.append(f"")
                    
                    # Add Total Net Jirama (Ar)
                    if row in ['Electricité', 'Eau']:
                        values_with_ar.append('{:_}'.format(float(f"{self.total_net[row]:.2f}")).replace('_', ' ') + " Ar")
                    else:
                        total = self.total_net['Electricité'] + self.total_net['Eau']
                        values_with_ar.append(('{:_}'.format(float(f"{total:.2f}"))).replace('_', ' ') + " Ar")
                    
                    data.append(values_with_ar)
                else:
                    # Add Somme(Villas+R) for other sections
                    if section == self.data['sections'][2]:
                        regular_sum = sum(self.values[section][row][col] for col in self.data['columns'] if col != 'Général')
                        values.append(f"{regular_sum:.2f}".rstrip('0').rstrip('.'))
                    
                    # Default empty values for other columns
                    values.extend([""] * 2)
                    
                    data.append(values)
                
            
        # Create PDF
        doc = SimpleDocTemplate(file_path, pagesize=landscape(A4))

        # Define a paragraph style for wrapping header text
        header_style = ParagraphStyle(
            name='HeaderStyle',
            fontName='Helvetica-Bold',
            fontSize=12,
            textColor=colors.whitesmoke,
            alignment=1,  # Center alignment
            leading=14    # Line spacing
        )

        # Process the data to wrap ONLY header text
        processed_data = []
        for row_idx, row in enumerate(data):
            processed_row = []
            for col_idx, cell in enumerate(row):
                # Only apply wrapping to the header row
                if row_idx == 0:  # Header row
                    processed_row.append(Paragraph(str(cell), header_style))
                else:
                    # Keep original data for non-headers
                    processed_row.append(cell)
            processed_data.append(processed_row)

        # Calculate column widths based on content
        col_count = len(processed_data[0])
        col_widths = [0] * col_count
        
        # Find the maximum width needed for each column
        for row in processed_data:
            for i, cell in enumerate(row):
                if isinstance(cell, str):
                    # For string cells, estimate width based on character count
                    width = len(cell) * 9  # Approximate 9 points per character
                elif isinstance(cell, Paragraph):
                    # For paragraphs, estimate the width based on text length
                    width = len(cell.text) * 9
                else:
                    # Default width for other types
                    width = 80
                    
                col_widths[i] = max(col_widths[i], width)
        
        # Add padding to each column width
        for i in range(col_count):
            col_widths[i] += 20
        
        # Ensure minimum widths for each column
        col_widths = [max(width, 80) for width in col_widths]
        
        # Calculate total width of all columns
        total_width = sum(col_widths)
        
        # Scale if needed to fit page width
        available_width = landscape(A4)[0] - ((doc.leftMargin + doc.rightMargin)/3) #part of what determines the max width of cols
        if total_width > available_width:
            scale_factor = available_width / total_width
            col_widths = [width * scale_factor for width in col_widths]

        # Create the table with the processed data and calculated column widths
        table = Table(processed_data, colWidths=col_widths)

        # Create the frame
        frame = Frame(
            doc.leftMargin, 
            doc.bottomMargin,
            doc.width, 
            doc.height,
            id='normal'
        )

        # Style definition
        style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),  # Vertical alignment for headers only
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Center alignment for headers
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),     # Left alignment for first column
            ('ALIGN', (1, 1), (-1, -1), 'RIGHT'),   # Right alignment for all other cells
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('BACKGROUND', (0, 1), (-1, 1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('SPAN', (7, 1), (10, 1)),  # Updated span for date (reduced by 1)
        ])

        # Add styles for section headers
        row_idx = 2
        for section in self.data['sections']:
            style.add('BACKGROUND', (0, row_idx), (-1, row_idx), colors.lightblue)
            style.add('SPAN', (0, row_idx), (-1, row_idx))  # Span for section header
            style.add('ALIGN', (0, row_idx), (-1, row_idx), 'CENTER')  # Center alignment for section headers
            row_idx += 1
            
            row_idx += len(self.data['rows'][section])

        # Apply the style to the table
        table.setStyle(style)

        # Calculate required table dimensions
        table_width, table_height = table.wrap(0, 0)
        available_width, available_height = landscape(A4)

        # Scale if needed
        if table_width > available_width or table_height > available_height:
            # Calculate scaling factor
            width_ratio = available_width / table_width if table_width > available_width else 1
            height_ratio = available_height / table_height if table_height > available_height else 1
            scale_factor = min(width_ratio, height_ratio) * 0.95  # 5% margin for safety
            
            # Reduce font size and cell padding proportionally
            for i, command in enumerate(style._cmds):
                if command[0] == 'FONTSIZE':
                    new_size = int(command[3] * scale_factor)
                    style._cmds[i] = (command[0], command[1], command[2], max(new_size, 6))  # Minimum 6pt font
                elif command[0] in ('TOPPADDING', 'BOTTOMPADDING', 'LEFTPADDING', 'RIGHTPADDING'):
                    new_padding = command[3] * scale_factor
                    style._cmds[i] = (command[0], command[1], command[2], new_padding)

        elements = [table]
        doc.addPageTemplates([PageTemplate(id='OneCol', frames=frame)])
        doc.build(elements)
        messagebox.showinfo("Enregistrement réussi", f"Décompte enregistré dans: {file_path}")

    def save_as_excel(self, file_path):
        # Create a DataFrame to save
        # Create data for PDF
        data = []
        headers = ['Category'] + self.data['columns'] + ['Somme(Villas+R)', 'Prix unitaire', 'Total Net Jirama']
        data.append(headers)
        
        # Add date row
        data.append(["Dates ou période:", "", "", "", "", "", "", f"{self.str_date}", "", "", ""])
        
        # Add data by section
        for section in self.data['sections']:
            # Add section header if needed
            
            data.append([section] + [""] * 10)
            
            for row in self.data['rows'][section]:
                values = [row]
                
                # Add values for each column
                for col in self.data['columns']:
                    #add empty strings instead of values for the General rows in the A PAYER section
                    if col == 'Général':
                        if section == 'A PAYER':
                            values.append(f"")
                        else:
                            values.append(f"{self.values[section][row][col]}".rstrip('0').rstrip('.'))
                    else:                    
                        values.append(f"{self.values[section][row][col]:.3f}".rstrip('0').rstrip('.') if isinstance(self.values[section][row][col], float) else self.values[section][row][col])
                
                # Add sum, difference, etc.
                if section == 'A PAYER':
                    # For A PAYER rows, add Ar to the values
                    values_with_ar = [('{:_}'.format(round(float(value), 2))).replace('_', ' ') + " Ar" if (i > 0 and i < 7) else value for i, value in enumerate(values)]
                    
                    # Add empty Somme(Villas+R) in the cells
                    values_with_ar.append(f"")
                    
                    
                    
                    # Add Prix unitaire (Ar)
                    if row in ['Electricité', 'Eau']:
                        values_with_ar.append(f"{self.price_per_unit[row]} Ar".rstrip('0').rstrip('.'))
                    else:
                        values_with_ar.append(f"")
                    
                    # Add Total Net Jirama (Ar)
                    if row in ['Electricité', 'Eau']:
                        values_with_ar.append('{:_}'.format(float(f"{self.total_net[row]:.2f}")).replace('_', ' ') + " Ar")
                    else:
                        total = self.total_net['Electricité'] + self.total_net['Eau']
                        values_with_ar.append(('{:_}'.format(float(f"{total:.2f}"))).replace('_', ' ') + " Ar")
                    
                    data.append(values_with_ar)
                else:
                    # Add Somme(Villas+R) for other sections
                    if section == self.data['sections'][2]:
                        regular_sum = sum(self.values[section][row][col] for col in self.data['columns'] if col != 'Général')
                        values.append(f"{regular_sum:.2f}".rstrip('0').rstrip('.'))
                    
                    # Default empty values for other columns
                    values.extend(["", ""])
                    
                    data.append(values)
        
        df = pd.DataFrame(data)
        # Write DataFrame to Excel
        df.to_excel(file_path, header=headers, index=False)
        messagebox.showinfo("Export Successful", f"Data exported to {file_path}")
    
    

def resource_path(relative_path):
    """Get absolute path to resource, for PyInstaller"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def main():
    root = tk.Tk()
    sv_ttk.set_theme("light")
    app = UtilityBillCalculator(root)
    root.mainloop()

if __name__ == "__main__":
    main()

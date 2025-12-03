import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import duckdb
import os
from openai import OpenAI

class ExcelSQLApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel SQL Transformer")
        self.root.geometry("1000x700")

        # State
        self.df = None
        self.original_df = None
        self.table_name = "data"
        self.con = duckdb.connect(database=':memory:')
        
        # OpenAI State
        self.openai_api_key = None
        self.openai_model = None
        self.openai_client = None

        # UI Setup
        self.setup_menu()
        self.setup_layout()

    def setup_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Load Excel File", command=self.load_file)
        file_menu.add_command(label="Export Results", command=self.export_data)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)

    def setup_layout(self):
        # Main split container
        self.paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Left Panel: Data Display
        self.left_frame = ttk.Frame(self.paned_window)
        self.paned_window.add(self.left_frame, weight=1)

        # Treeview for data
        self.tree_frame = ttk.Frame(self.left_frame)
        self.tree_frame.pack(fill=tk.BOTH, expand=True)
        
        self.tree_scroll_y = ttk.Scrollbar(self.tree_frame)
        self.tree_scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_scroll_x = ttk.Scrollbar(self.tree_frame, orient=tk.HORIZONTAL)
        self.tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree = ttk.Treeview(self.tree_frame, yscrollcommand=self.tree_scroll_y.set, xscrollcommand=self.tree_scroll_x.set)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        self.tree_scroll_y.config(command=self.tree.yview)
        self.tree_scroll_x.config(command=self.tree.xview)

        self.status_label = ttk.Label(self.left_frame, text="No file loaded.")
        self.status_label.pack(fill=tk.X, pady=2)

        # Right Panel: SQL Editor
        self.right_frame = ttk.Frame(self.paned_window)
        self.paned_window.add(self.right_frame, weight=1)

        # OpenAI Configuration Panel
        self.setup_openai_panel()
        
        # Natural Language Query Panel
        self.setup_nl_query_panel()

        lbl_instruction = ttk.Label(self.right_frame, text=f"Write SQL Query (Table name: '{self.table_name}'):")
        lbl_instruction.pack(anchor=tk.W, pady=5)

        self.txt_sql = tk.Text(self.right_frame, height=10, width=40)
        self.txt_sql.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
        self.txt_sql.insert(tk.END, f"SELECT * FROM {self.table_name}")

        btn_frame = ttk.Frame(self.right_frame)
        btn_frame.pack(fill=tk.X, pady=5)

        self.btn_transform = ttk.Button(btn_frame, text="Transform", command=self.transform_data)
        self.btn_transform.pack(side=tk.LEFT, padx=5)
        
        self.btn_reset = ttk.Button(btn_frame, text="Reset Data", command=self.reset_data)
        self.btn_reset.pack(side=tk.LEFT, padx=5)

        self.btn_export = ttk.Button(btn_frame, text="Export Results", command=self.export_data, state=tk.DISABLED)
        self.btn_export.pack(side=tk.RIGHT, padx=5)

    def load_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        try:
            self.status_label.config(text="Loading file...")
            self.root.update()
            
            # Read Excel
            self.original_df = pd.read_excel(file_path)
            self.df = self.original_df.copy()
            
            # Register with DuckDB
            self.con.register(self.table_name, self.df)

            self.display_dataframe(self.df)
            self.status_label.config(text=f"Loaded {len(self.df)} rows from {os.path.basename(file_path)}")
            self.btn_export.config(state=tk.NORMAL)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file:\n{str(e)}")
            self.status_label.config(text="Error loading file.")

    def display_dataframe(self, dataframe):
        # Clear current view
        self.tree.delete(*self.tree.get_children())
        
        # Set columns
        cols = list(dataframe.columns)
        self.tree["columns"] = cols
        self.tree["show"] = "headings"

        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        # Add data (limit to first 1000 rows for performance if needed, but requirements said 'all loaded data')
        # For very large datasets, Treeview can be slow. Let's show up to 2000 rows for responsiveness.
        # If the user wants to see more, they should filter.
        
        max_display_rows = 2000
        rows_to_show = dataframe.head(max_display_rows)
        
        for index, row in rows_to_show.iterrows():
            self.tree.insert("", "end", values=list(row))
            
        if len(dataframe) > max_display_rows:
            self.status_label.config(text=f"Showing first {max_display_rows} of {len(dataframe)} rows.")
        else:
            self.status_label.config(text=f"Showing {len(dataframe)} rows.")

    def transform_data(self):
        if self.original_df is None:
            messagebox.showwarning("Warning", "Please load an Excel file first.")
            return

        query = self.txt_sql.get("1.0", tk.END).strip()
        if not query:
            return

        try:
            self.con.register(self.table_name, self.original_df)
            
            result_df = self.con.execute(query).df()
            self.df = result_df # Update current view df
            self.display_dataframe(self.df)
            
        except Exception as e:
            messagebox.showerror("SQL Error", f"Query failed:\n{str(e)}")

    def reset_data(self):
        if self.original_df is not None:
            self.df = self.original_df.copy()
            self.display_dataframe(self.df)
            self.txt_sql.delete("1.0", tk.END)
            self.txt_sql.insert(tk.END, f"SELECT * FROM {self.table_name}")

    def export_data(self):
        if self.df is None:
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel File", "*.xlsx"), ("CSV File", "*.csv")]
        )
        
        if not file_path:
            return
            
        try:
            if file_path.endswith('.csv'):
                self.df.to_csv(file_path, index=False)
            else:
                self.df.to_excel(file_path, index=False)
            messagebox.showinfo("Success", "Data exported successfully!")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export:\n{str(e)}")
    
    def setup_openai_panel(self):
        """Setup OpenAI configuration panel"""
        openai_frame = ttk.LabelFrame(self.right_frame, text="OpenAI Configuration", padding=10)
        openai_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # API Key
        api_key_frame = ttk.Frame(openai_frame)
        api_key_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(api_key_frame, text="API Key:").pack(side=tk.LEFT, padx=(0, 5))
        self.txt_api_key = ttk.Entry(api_key_frame, show="*", width=30)
        self.txt_api_key.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        
        self.show_key_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(api_key_frame, text="Show", variable=self.show_key_var, 
                       command=lambda: self.txt_api_key.config(show="" if self.show_key_var.get() else "*")).pack(side=tk.LEFT)
        
        # Model Selection
        model_frame = ttk.Frame(openai_frame)
        model_frame.pack(fill=tk.X, pady=2)
        
        ttk.Label(model_frame, text="Model:").pack(side=tk.LEFT, padx=(0, 5))
        self.cmb_model = ttk.Combobox(model_frame, values=[
            "gpt-5-mini",
            "gpt-5-nano",
            "gpt-4.1-mini"
        ], state="readonly", width=28)
        self.cmb_model.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.cmb_model.current(0)  # Default to gpt-5-mini
        
        # Save Button
        self.btn_save_config = ttk.Button(openai_frame, text="Save Configuration", command=self.save_openai_config)
        self.btn_save_config.pack(pady=(5, 0))
        
        self.lbl_config_status = ttk.Label(openai_frame, text="", foreground="gray")
        self.lbl_config_status.pack()
    
    def setup_nl_query_panel(self):
        """Setup natural language query panel"""
        nl_frame = ttk.LabelFrame(self.right_frame, text="Generate SQL from Question", padding=10)
        nl_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(nl_frame, text="Ask a question about your data:").pack(anchor=tk.W, pady=(0, 5))
        
        self.txt_question = tk.Text(nl_frame, height=3, width=40)
        self.txt_question.pack(fill=tk.X, pady=(0, 5))
        
        self.btn_generate = ttk.Button(nl_frame, text="Generate SQL", command=self.generate_sql_from_question)
        self.btn_generate.pack()
        
        self.lbl_generate_status = ttk.Label(nl_frame, text="", foreground="gray")
        self.lbl_generate_status.pack()
    
    def save_openai_config(self):
        """Save and validate OpenAI configuration"""
        api_key = self.txt_api_key.get().strip()
        model = self.cmb_model.get()
        
        if not api_key:
            messagebox.showwarning("Warning", "Please enter an API key.")
            return
        
        if not model:
            messagebox.showwarning("Warning", "Please select a model.")
            return
        
        try:
            # Test the API key by creating a client
            self.openai_client = OpenAI(api_key=api_key)
            self.openai_api_key = api_key
            self.openai_model = model
            
            self.lbl_config_status.config(text=f"✓ Configured with {model}", foreground="green")
            messagebox.showinfo("Success", "OpenAI configuration saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to configure OpenAI:\n{str(e)}")
            self.lbl_config_status.config(text="✗ Configuration failed", foreground="red")
    
    def get_table_metadata(self):
        """Extract table metadata (column names and types)"""
        if self.original_df is None:
            return None
        
        metadata = []
        for col in self.original_df.columns:
            dtype = str(self.original_df[col].dtype)
            metadata.append(f"{col} ({dtype})")
        
        return metadata
    
    def create_openai_prompt(self, question, metadata):
        """Create prompt for OpenAI API"""
        columns_str = ", ".join(metadata)
        
        prompt = f"""You are a SQL expert. Convert the following natural language question into a DuckDB SQL query.

Table name: {self.table_name}
Columns: {columns_str}

Question: {question}

IMPORTANT:
- Return ONLY the SQL query, no explanations or markdown formatting
- Use DuckDB SQL syntax
- The table name is '{self.table_name}'
- Do not include any text before or after the SQL query

SQL Query:"""
        
        return prompt
    
    def generate_sql_from_question(self):
        """Generate SQL from natural language question using OpenAI"""
        # Validate configuration
        if not self.openai_client or not self.openai_model:
            messagebox.showwarning("Warning", "Please configure OpenAI API key and model first.")
            return
        
        # Validate data is loaded
        if self.original_df is None:
            messagebox.showwarning("Warning", "Please load an Excel file first.")
            return
        
        # Get question
        question = self.txt_question.get("1.0", tk.END).strip()
        if not question:
            messagebox.showwarning("Warning", "Please enter a question.")
            return
        
        try:
            self.lbl_generate_status.config(text="Generating SQL...", foreground="blue")
            self.root.update()
            
            # Get table metadata
            metadata = self.get_table_metadata()
            
            # Create prompt
            prompt = self.create_openai_prompt(question, metadata)
            
            # Call OpenAI API
            response = self.openai_client.chat.completions.create(
                model=self.openai_model,
                messages=[
                    {"role": "system", "content": "You are a SQL expert that converts natural language to DuckDB SQL queries."},
                    {"role": "user", "content": prompt}
                ]
            )
            
            # Extract SQL from response
            sql_query = response.choices[0].message.content.strip()
            
            # Clean up the response (remove markdown code blocks if present)
            if sql_query.startswith("```"):
                lines = sql_query.split("\n")
                sql_query = "\n".join(lines[1:-1]) if len(lines) > 2 else sql_query
                sql_query = sql_query.replace("```sql", "").replace("```", "").strip()
            
            # Insert into SQL text area
            self.txt_sql.delete("1.0", tk.END)
            self.txt_sql.insert(tk.END, sql_query)
            
            self.lbl_generate_status.config(text="✓ SQL generated successfully", foreground="green")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate SQL:\n{str(e)}")
            self.lbl_generate_status.config(text="✗ Generation failed", foreground="red")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSQLApp(root)
    root.mainloop()

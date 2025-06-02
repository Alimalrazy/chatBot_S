import os
import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog, ttk
from tkinter import font as tkFont
from processor import ExcelRAGProcessor
import threading

class ModernExcelChatApp:
    def __init__(self, root):
        self.root = root
        self.processor = ExcelRAGProcessor()
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the modern UI."""
        self.root.title("📊 Smart Excel RAG Chatbot")
        self.root.geometry("900x700")
        self.root.configure(bg='#f0f2f5')
        
        # Configure styles
        style = ttk.Style()
        style.theme_use('clam')
        
        # Main container
        main_frame = tk.Frame(self.root, bg='#f0f2f5')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header
        header_frame = tk.Frame(main_frame, bg='#2c3e50', height=80)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        header_frame.pack_propagate(False)
        
        header_label = tk.Label(
            header_frame, 
            text="📊 Smart Excel RAG Chatbot",
            font=tkFont.Font(family="Arial", size=18, weight="bold"),
            fg='white', 
            bg='#2c3e50'
        )
        header_label.pack(expand=True)
        
        subtitle_label = tk.Label(
            header_frame,
            text="Powered by Custom Header Detection + Gemini AI",
            font=tkFont.Font(family="Arial", size=10),
            fg='#bdc3c7',
            bg='#2c3e50'
        )
        subtitle_label.pack()
        
        # Control panel
        control_frame = tk.Frame(main_frame, bg='#f0f2f5')
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # API Key frame
        api_frame = tk.Frame(control_frame, bg='#f0f2f5')
        api_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Label(api_frame, text="Gemini API Key:", font=tkFont.Font(family="Arial", size=10), bg='#f0f2f5').pack(side=tk.LEFT)
        self.api_key_var = tk.StringVar(value=os.getenv('GEMINI_API_KEY', '')[:20] + '...' if os.getenv('GEMINI_API_KEY') else '')
        self.api_entry = tk.Entry(api_frame, textvariable=self.api_key_var, width=30, show='*')
        self.api_entry.pack(side=tk.LEFT, padx=(5, 10))
        
        tk.Button(
            api_frame,
            text="Set API Key",
            command=self.update_api_key,
            bg='#3498db',
            fg='white',
            font=tkFont.Font(family="Arial", size=9),
            relief=tk.FLAT,
            padx=10
        ).pack(side=tk.LEFT)
        
        # File operations frame
        file_frame = tk.Frame(control_frame, bg='#f0f2f5')
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        tk.Button(
            file_frame,
            text="📁 Load Files from Directory",
            command=self.load_files_from_directory,
            bg='#27ae60',
            fg='white',
            font=tkFont.Font(family="Arial", size=10),
            relief=tk.FLAT,
            padx=15,
            pady=5
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        tk.Button(
            file_frame,
            text="📂 Select Specific Files",
            command=self.select_files,
            bg='#e67e22',
            fg='white',
            font=tkFont.Font(family="Arial", size=10),
            relief=tk.FLAT,
            padx=15,
            pady=5
        ).pack(side=tk.LEFT, padx=(0, 10))
        
        tk.Button(
            file_frame,
            text="🔄 Clear All",
            command=self.clear_all,
            bg='#e74c3c',
            fg='white',
            font=tkFont.Font(family="Arial", size=10),
            relief=tk.FLAT,
            padx=15,
            pady=5
        ).pack(side=tk.RIGHT)
        
        # Chat area with better styling
        chat_frame = tk.Frame(main_frame, bg='white', relief=tk.RAISED, bd=1)
        chat_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Chat header
        chat_header = tk.Frame(chat_frame, bg='#34495e', height=30)
        chat_header.pack(fill=tk.X)
        chat_header.pack_propagate(False)
        
        tk.Label(
            chat_header,
            text="💬 Chat Area",
            font=tkFont.Font(family="Arial", size=10, weight="bold"),
            fg='white',
            bg='#34495e'
        ).pack(side=tk.LEFT, padx=10, pady=5)
        
        # Scrolled text area
        self.chat_area = scrolledtext.ScrolledText(
            chat_frame,
            wrap=tk.WORD,
            width=80,
            height=20,
            font=tkFont.Font(family="Consolas", size=10),
            bg='#ffffff',
            fg='#2c3e50',
            relief=tk.FLAT,
            padx=10,
            pady=10
        )
        self.chat_area.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Input area
        input_frame = tk.Frame(main_frame, bg='#f0f2f5')
        input_frame.pack(fill=tk.X)
        
        self.input_box = tk.Entry(
            input_frame,
            font=tkFont.Font(family="Arial", size=11),
            relief=tk.RAISED,
            bd=1,
            bg='white'
        )
        self.input_box.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10), ipady=8)
        self.input_box.bind('<Return>', lambda e: self.ask_question())
        
        self.ask_button = tk.Button(
            input_frame,
            text="🚀 Ask",
            command=self.ask_question,
            bg='#9b59b6',
            fg='white',
            font=tkFont.Font(family="Arial", size=11, weight="bold"),
            relief=tk.FLAT,
            padx=20,
            pady=8
        )
        self.ask_button.pack(side=tk.RIGHT)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready - Load Excel files to start chatting!")
        status_bar = tk.Label(
            main_frame,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W,
            bg='#ecf0f1',
            font=tkFont.Font(family="Arial", size=9)
        )
        status_bar.pack(fill=tk.X, pady=(10, 0))
        
        # Welcome message
        self.add_message("system", "🎯 Welcome to Smart Excel RAG Chatbot!")
        self.add_message("system", "📋 Features:")
        self.add_message("system", "  • Custom header detection using 3-row sequential pattern")
        self.add_message("system", "  • Automatic removal of rows above detected header")
        self.add_message("system", "  • AI-powered chat using Gemini 2.0 Flash")
        self.add_message("system", "  • Semantic search across all your Excel data")
        
        # Auto-load files from current directory
        self.auto_load_files()
    
    def auto_load_files(self):
        """Automatically find and load Excel files from current directory."""
        current_dir = os.getcwd()
        excel_files = []
        
        for file_name in os.listdir(current_dir):
            if file_name.lower().endswith(('.xlsx', '.xls', '.csv')):
                excel_files.append(os.path.join(current_dir, file_name))
        
        if excel_files:
            self.add_message("system", f"\n🔍 Auto-detected {len(excel_files)} Excel files in current directory:")
            for file_path in excel_files:
                self.add_message("system", f"  • {os.path.basename(file_path)}")
            
            self.add_message("system", "\n🚀 Auto-loading files...")
            self.process_files(excel_files)
        else:
            self.add_message("system", "\n❌ No Excel files found in current directory.")
            self.add_message("system", "🚀 Use the buttons above to load Excel files manually!")
    
    def add_message(self, sender, message):
        """Add a message to the chat area."""
        if sender == "system":
            self.chat_area.insert(tk.END, f"🤖 System: {message}\n\n")
        elif sender == "user":
            self.chat_area.insert(tk.END, f"👤 You: {message}\n")
        elif sender == "bot":
            self.chat_area.insert(tk.END, f"🤖 Assistant: {message}\n\n")
        
        self.chat_area.see(tk.END)
        self.chat_area.update()
    
    def update_status(self, message):
        """Update status bar."""
        self.status_var.set(message)
        self.root.update()
    
    def update_api_key(self):
        """Update the API key."""
        api_key = self.api_entry.get()
        if api_key and not api_key.endswith('...'):
            self.processor.gemini_key = api_key
            os.environ['GEMINI_API_KEY'] = api_key
            self.add_message("system", "✅ API key updated successfully!")
            # Update display to show masked version
            self.api_key_var.set(api_key[:20] + '...' if len(api_key) > 20 else api_key)
        else:
            messagebox.showwarning("Warning", "Please enter a valid API key.")
    
    def load_files_from_directory(self):
        """Load all Excel files from selected directory."""
        directory = filedialog.askdirectory(title="Select Directory with Excel Files")
        if not directory:
            return
        
        excel_files = []
        for file_name in os.listdir(directory):
            if file_name.lower().endswith(('.xlsx', '.xls', '.csv')):
                excel_files.append(os.path.join(directory, file_name))
        
        if not excel_files:
            messagebox.showwarning("Warning", "No Excel or CSV files found in the selected directory.")
            return
        
        self.process_files(excel_files)
    
    def select_files(self):
        """Select specific Excel files."""
        files = filedialog.askopenfilenames(
            title="Select Excel Files",
            filetypes=[
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("All supported", "*.xlsx *.xls *.csv")
            ]
        )
        
        if files:
            self.process_files(list(files))
    
    def process_files(self, file_paths):
        """Process selected files in a separate thread."""
        def process_thread():
            try:
                self.update_status("🔄 Processing files...")
                
                if not hasattr(self, '_files_logged'):
                    self.add_message("system", f"📁 Processing {len(file_paths)} files:")
                    for file_path in file_paths:
                        self.add_message("system", f"  • {os.path.basename(file_path)}")
                    self._files_logged = True
                
                # Open files and process
                uploaded_files = []
                for file_path in file_paths:
                    # Create a simple file-like object
                    class FileWrapper:
                        def __init__(self, path):
                            self.name = os.path.basename(path)
                            self.path = path
                        
                        def read(self):
                            with open(self.path, 'rb') as f:
                                return f.read()
                    
                    uploaded_files.append(FileWrapper(file_path))
                
                # Process files
                self.add_message("system", "\n🔄 Processing with custom header detection...")
                results = self.processor.process_files(uploaded_files)
                
                # Display results
                self.add_message("system", "\n📊 Processing Results:")
                success_count = 0
                
                for file_id, result in results.items():
                    if result['status'] == 'success':
                        success_count += 1
                        self.add_message("system", 
                            f"✅ {result['filename']}:\n"
                            f"   • Header found at row: {result['header_row'] + 1}\n"
                            f"   • Rows deleted above header: {result['rows_deleted_above_header']}\n"
                            f"   • Data rows: {result['row_count']}\n"
                            f"   • Columns: {result['column_count']}\n"
                            f"   • Sample columns: {', '.join(result['columns'][:3])}...\n"
                        )
                    else:
                        self.add_message("system", f"❌ {result['filename']}: {result.get('error', 'Unknown error')}")
                
                if success_count > 0:
                    self.add_message("system", f"\n🎉 Successfully processed {success_count}/{len(results)} files!")
                    self.add_message("system", "💬 You can now start asking questions about your data!")
                    self.add_message("system", "\n💡 Try asking: 'What files do I have?' or 'Show me data for ID 157408'")
                    
                    # Add data summary for debugging
                    summary = self.processor.get_all_data_summary()
                    self.add_message("system", f"\n📋 Data loaded successfully! Use 'show all data' to see complete summary.")
                    
                    self.update_status(f"✅ Ready - {success_count} files loaded")
                else:
                    self.add_message("system", "❌ No files were processed successfully.")
                    self.update_status("❌ Processing failed")
                
            except Exception as e:
                error_msg = f"Processing error: {str(e)}"
                self.add_message("system", f"❌ {error_msg}")
                self.update_status("❌ Error occurred")
        
        # Run processing in separate thread to avoid UI freezing
        thread = threading.Thread(target=process_thread)
        thread.daemon = True
        thread.start()
    
    def ask_question(self):
        """Process user question."""
        question = self.input_box.get().strip()
        if not question:
            messagebox.showwarning("Warning", "Please enter a question.")
            return
        
        # Handle special commands
        if question.lower() in ['show all data', 'show data summary', 'debug data']:
            summary = self.processor.get_all_data_summary()
            self.add_message("user", question)
            self.add_message("bot", summary)
            self.input_box.delete(0, tk.END)
            return
        
        self.add_message("user", question)
        self.input_box.delete(0, tk.END)
        
        if not self.processor.processed_files:
            self.add_message("bot", "No data loaded. Please process Excel files first.")
            return
        
        def query_thread():
            try:
                self.update_status("🤔 AI is thinking...")
                self.ask_button.config(state='disabled', text='🤔 Thinking...')
                
                answer = self.processor.query(question)
                self.add_message("bot", answer)
                
                self.update_status("✅ Response generated")
                self.ask_button.config(state='normal', text='🚀 Ask')
                
            except Exception as e:
                self.add_message("bot", f"Sorry, I encountered an error: {str(e)}")
                self.update_status("❌ Error in response")
                self.ask_button.config(state='normal', text='🚀 Ask')
        
        # Run query in separate thread
        thread = threading.Thread(target=query_thread)
        thread.daemon = True
        thread.start()
    
    def clear_all(self):
        """Clear all data and reset."""
        if messagebox.askyesno("Confirm", "Are you sure you want to clear all data?"):
            self.processor = ExcelRAGProcessor(self.processor.gemini_key)
            self.chat_area.delete(1.0, tk.END)
            self.add_message("system", "🔄 All data cleared. Ready for new files!")
            self.update_status("Ready - Load Excel files to start chatting!")

def main():
    """Run the application."""
    root = tk.Tk()
    app = ModernExcelChatApp(root)
    
    try:
        root.mainloop()
    except KeyboardInterrupt:
        print("\nShutting down...")

if __name__ == "__main__":
    main()
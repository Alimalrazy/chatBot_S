import streamlit as st
import os
from dotenv import load_dotenv
import logging
from processor import ExcelRAGProcessor
import time
import glob

# Load environment variables
load_dotenv()

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# App configuration
st.set_page_config(
    page_title="📊 Smart Excel RAG Chatbot",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
.main-header {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white;
    padding: 2rem;
    border-radius: 10px;
    text-align: center;
    margin-bottom: 2rem;
}
.chat-message {
    padding: 1rem;
    margin: 0.5rem 0;
    border-radius: 10px;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}
.user-message {
    background: #e3f2fd;
    margin-left: 2rem;
}
.assistant-message {
    background: #f1f8e9;
    margin-right: 2rem;
}
.file-status {
    padding: 0.5rem;
    margin: 0.5rem 0;
    border-radius: 5px;
    border-left: 4px solid #28a745;
    background: #d4edda;
    color: #155724;
}
.gemini-badge {
    background: linear-gradient(135deg, #4285f4 0%, #34a853 50%, #fbbc05 75%, #ea4335 100%);
    color: white;
    padding: 0.3rem 0.8rem;
    border-radius: 15px;
    font-size: 0.8rem;
    font-weight: bold;
    display: inline-block;
    margin-left: 1rem;
}
.api-input {
    background: #fff3cd;
    border: 1px solid #ffeaa7;
    border-radius: 8px;
    padding: 1rem;
    margin: 1rem 0;
}
.stButton>button {
    width: 100%;
    border-radius: 5px;
    height: 3em;
}
</style>
""", unsafe_allow_html=True)

def initialize_session_state():
    """Initialize session state variables"""
    if 'processor' not in st.session_state:
        # Try to get API key from environment variables first
        gemini_key = os.getenv('GEMINI_API_KEY', '')
        st.session_state.processor = ExcelRAGProcessor(gemini_key)
    
    if 'chat_history' not in st.session_state:
        st.session_state.chat_history = []
    
    if 'processing_status' not in st.session_state:
        st.session_state.processing_status = None

def display_header():
    """Display the application header"""
    st.markdown("""
    <div class="main-header">
        <h1>🎯 Smart Excel RAG Chatbot</h1>
        <p>Powered by Google Gemini 2.0 Flash + LangChain RAG</p>
        <span class="gemini-badge">🤖 Gemini 2.0 Flash</span>
    </div>
    """, unsafe_allow_html=True)

def load_excel_files_from_folder():
    """Load Excel file paths from the current directory"""
    excel_files = []
    # Look for Excel files in the current directory
    for ext in ['*.xlsx', '*.xls', '*.csv']:
        excel_files.extend(glob.glob(ext))
    
    if not excel_files:
        st.error("❌ No Excel files found in the current directory!")
        return None
    
    return excel_files

def display_sidebar():
    """Display the sidebar with configuration options"""
    with st.sidebar:
        st.header("🔧 Configuration")
        
        # API Key configuration
        gemini_key = os.getenv('GEMINI_API_KEY')
        if gemini_key:
            st.success("✅ Gemini API Key configured from .env")
        else:
            st.markdown("""
            <div class="api-input">
                <strong>⚠️ No API key found in .env file</strong><br>
                Please enter your Gemini API key below or add it to your .env file
            </div>
            """, unsafe_allow_html=True)
            
            api_key_input = st.text_input(
                "Gemini API Key", 
                type="password",
                help="Get your API key from https://aistudio.google.com/app/apikey"
            )
            
            if api_key_input:
                st.session_state.processor = ExcelRAGProcessor(api_key_input)
                st.success("✅ API key configured!")
        
        st.divider()
        
        # Auto-load Excel files
        if st.button("📂 Load Excel Files from Folder", type="primary"):
            with st.spinner("🔄 Loading Excel files from folder..."):
                try:
                    file_paths = load_excel_files_from_folder()
                    if file_paths:
                        start_time = time.time()
                        results = st.session_state.processor.process_files(file_paths)
                        processing_time = time.time() - start_time
                        
                        st.session_state.results = results
                        st.session_state.processing_status = {
                            'success': True,
                            'time': processing_time,
                            'file_count': len(results)
                        }
                        
                        success_count = sum(1 for r in results.values() if r['status'] == 'success')
                        if success_count > 0:
                            st.success(f"✅ Successfully processed {success_count}/{len(results)} files in {processing_time:.2f}s!")
                        else:
                            st.error("❌ No files were processed successfully")
                    
                except Exception as e:
                    logger.error(f"Error processing files: {e}")
                    st.error(f"❌ Processing Error: {str(e)}")
                    st.session_state.processing_status = {
                        'success': False,
                        'error': str(e)
                    }
        
        # Display file status
        if hasattr(st.session_state, 'results'):
            st.header("📊 File Status")
            for file_data in st.session_state.results.values():
                if file_data['status'] == 'success':
                    st.markdown(f"""
                    <div class="file-status">
                        <strong>✅ {file_data['filename']}</strong><br>
                        📊 {file_data['row_count']} rows, {file_data['column_count']} columns<br>
                        🎯 Header detected at row {file_data['header_row'] + 1}
                    </div>
                    """, unsafe_allow_html=True)
                else:
                    st.error(f"❌ {file_data['filename']}: {file_data.get('error', 'Unknown error')}")
        
        st.divider()
        
        # Actions section
        st.header("🔄 Actions")
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("🗑️ Clear All", help="Clear all data and reset"):
                for key in list(st.session_state.keys()):
                    del st.session_state[key]
                st.rerun()
        
        with col2:
            if st.button("💬 Reset Chat", help="Clear chat history"):
                if 'chat_history' in st.session_state:
                    st.session_state.chat_history = []
                st.success("Chat reset!")

def display_chat_interface():
    """Display the main chat interface"""
    if not hasattr(st.session_state, 'results') or not st.session_state.results:
        display_welcome_screen()
        return
    
    # Chat input
    user_input = st.chat_input("Ask a question about your Excel data...")
    
    if user_input:
        # Add user message to chat history
        st.session_state.chat_history.append({"role": "user", "content": user_input})
        
        # Get AI response
        with st.spinner("🤔 Thinking..."):
            try:
                response = st.session_state.processor.query(user_input)
                st.session_state.chat_history.append({"role": "assistant", "content": response})
            except Exception as e:
                logger.error(f"Error getting response: {e}")
                error_msg = f"Sorry, I encountered an error: {str(e)}"
                st.session_state.chat_history.append({"role": "assistant", "content": error_msg})
    
    # Display chat history
    for message in st.session_state.chat_history:
        if message["role"] == "user":
            st.markdown(f"""
            <div class="chat-message user-message">
                <strong>You:</strong><br>{message["content"]}
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class="chat-message assistant-message">
                <strong>Assistant:</strong><br>{message["content"]}
            </div>
            """, unsafe_allow_html=True)

def display_welcome_screen():
    """Display the welcome screen with instructions"""
    st.markdown("""
    ## 🚀 Welcome to Smart Excel RAG Chatbot!
    
    This application uses **Google Gemini 2.0 Flash** with **RAG (Retrieval-Augmented Generation)** for intelligent Excel analysis.
    
    ### ✨ Key Features:
    - 📊 **Multi-file processing** with automatic header detection
    - 🔍 **Semantic search** across all your Excel data  
    - 🤖 **Natural language queries** powered by Gemini 2.0 Flash
    - 📈 **Smart calculations** and data aggregations
    - 💡 **Context-aware responses** using vector similarity
    - ⚡ **Lightning-fast** responses with Gemini 2.0 Flash
    
    ### 🏗️ Technical Stack:
    - **AI Model:** Google Gemini 2.0 Flash (latest & fastest)
    - **RAG System:** ChromaDB + HuggingFace Embeddings
    - **Data Processing:** Pandas + Auto Header Detection
    - **UI Framework:** Streamlit
    
    ### 📊 Supported Data:
    - Insurance agent data with commissions
    - Business metrics and KPIs  
    - Financial records and calculations
    - Any structured Excel/CSV data
    
    **👈 Start by configuring your Gemini API key and uploading Excel files!**
    
    ---
    
    ### 🔑 Quick Setup:
    1. Get your free API key: [Google AI Studio](https://aistudio.google.com/app/apikey)
    2. Add to `.env` file: `GEMINI_API_KEY=your_key_here`
    3. Upload your Excel files
    4. Start chatting with your data!
    """)
    
    # Show example queries
    st.markdown("### 💡 Example Queries You Can Try:")
    example_queries = [
        "📄 What files do I have loaded?",
        "👥 Show me all agent names",
        "📊 Calculate average Total Premium",
        "🔍 Find agent with Sl number 5",
        "👤 Tell me about Mohammad Zahedul Islam",  
        "📈 Compare commission data across files",
        "💰 What's the highest commission earned?",
        "📋 List all columns in my data"
    ]
    
    for query in example_queries:
        st.markdown(f"- {query}")

def main():
    """Main application function"""
    initialize_session_state()
    display_header()
    display_sidebar()
    display_chat_interface()

if __name__ == "__main__":
    main()
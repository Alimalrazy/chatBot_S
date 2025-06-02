# Smart Excel RAG Chatbot

A powerful chatbot application that uses Google's Gemini 2.0 Flash model and RAG (Retrieval-Augmented Generation) to analyze and answer questions about Excel data.

## 🌟 Features

- 📊 Multi-file Excel processing with automatic header detection
- 🔍 Semantic search across all your Excel data
- 🤖 Natural language queries powered by Gemini 2.0 Flash
- 📈 Smart calculations and data aggregations
- 💡 Context-aware responses using vector similarity
- ⚡ Lightning-fast responses with Gemini 2.0 Flash

## 🚀 Quick Start

1. Clone the repository:
```bash
git clone <your-repo-url>
cd <your-repo-name>
```

2. Create a virtual environment and activate it:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Create a `.env` file in the root directory and add your Gemini API key:
```
GEMINI_API_KEY=your_api_key_here
```

5. Run the application:
```bash
streamlit run app.py
```

## 🏗️ Technical Stack

- **AI Model:** Google Gemini 2.0 Flash
- **RAG System:** ChromaDB + HuggingFace Embeddings
- **Data Processing:** Pandas + Auto Header Detection
- **UI Framework:** Streamlit

## 📦 Deployment

### Local Deployment

1. Follow the Quick Start instructions above
2. Access the application at `http://localhost:8501`

### Cloud Deployment

#### Streamlit Cloud

1. Push your code to a GitHub repository
2. Go to [Streamlit Cloud](https://streamlit.io/cloud)
3. Create a new app and connect your repository
4. Add your `GEMINI_API_KEY` to the secrets management
5. Deploy!

#### Other Cloud Platforms

The application can be deployed on any platform that supports Python applications:

- Heroku
- AWS
- Google Cloud Platform
- Azure

Make sure to:
1. Set the `GEMINI_API_KEY` environment variable
2. Install the required dependencies
3. Run the application using `streamlit run app.py`

## 🔧 Configuration

The application can be configured through:

1. `.env` file for environment variables
2. `.streamlit/config.toml` for Streamlit settings
3. `requirements.txt` for Python dependencies

## 📊 Supported Data

- Insurance agent data with commissions
- Business metrics and KPIs
- Financial records and calculations
- Any structured Excel/CSV data

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- Google Gemini 2.0 Flash for the AI model
- Streamlit for the web framework
- LangChain for the RAG implementation
- ChromaDB for vector storage
# processor.py - UPDATED VERSION
import pandas as pd
import numpy as np
from typing import Dict, List, Any, Optional
import logging
from langchain_community.vectorstores import Chroma
from langchain_community.embeddings import HuggingFaceEmbeddings
from langchain.docstore.document import Document
import requests
import json
import os
import re

logger = logging.getLogger(__name__)

class ExcelRAGProcessor:
    """Enhanced processor for Excel files with systematic data search."""
    
    def __init__(self, gemini_key: str = None):
        self.processed_files = {}
        self.all_records = []  # Centralized data storage
        self.embeddings = None  # Initialize as None
        self.vector_store = None
        self.conversation_history = []
        self.gemini_key = gemini_key or os.getenv('GEMINI_API_KEY')
        
        if not self.gemini_key:
            logger.warning("No Gemini API key provided. Query functionality will be limited.")
    
    def _initialize_embeddings(self):
        """Lazy initialization of embeddings to avoid startup errors."""
        if self.embeddings is None:
            try:
                self.embeddings = HuggingFaceEmbeddings(
                    model_name="sentence-transformers/all-MiniLM-L6-v2",
                    model_kwargs={'device': 'cpu'},
                    encode_kwargs={'normalize_embeddings': True}
                )
                logger.info("Embeddings initialized successfully")
            except Exception as e:
                logger.error(f"Failed to initialize embeddings: {e}")
                # Fallback: create a simple text-based search
                self.embeddings = None
    
    def find_header_with_sequential_pattern(self, data: List[List]) -> Optional[int]:
        """
        Find header by detecting 3 sequential rows with both null and non-null values.
        The first row among these 3 rows will be considered the header.
        """
        for i in range(len(data) - 2):
            row1 = data[i] if i < len(data) else []
            row2 = data[i + 1] if i + 1 < len(data) else []
            row3 = data[i + 2] if i + 2 < len(data) else []
            
            # Check if all three rows have both null and non-null values
            if (self._has_mixed_values(row1) and 
                self._has_mixed_values(row2) and 
                self._has_mixed_values(row3)):
                
                logger.info(f"Found header pattern at rows {i}, {i+1}, {i+2}")
                return i  # Return the first row of the pattern as header
        
        # Fallback: look for first row with mostly non-null values
        for i, row in enumerate(data[:10]):  # Check first 10 rows
            non_null_count = sum(1 for cell in row if cell is not None and str(cell).strip())
            if non_null_count >= len(row) * 0.5:  # At least 50% non-null
                logger.info(f"Using fallback header detection at row {i}")
                return i
        
        return None
    
    def _has_mixed_values(self, row: List) -> bool:
        """Check if a row has both null and non-null values."""
        if not row:
            return False
        
        has_null = False
        has_non_null = False
        
        for cell in row:
            if cell is None or str(cell).strip() == '' or str(cell).lower() in ['nan', 'none']:
                has_null = True
            else:
                has_non_null = True
            
            # Early return if we found both
            if has_null and has_non_null:
                return True
        
        return has_null and has_non_null
    
    def process_files(self, file_paths: List[str]) -> Dict[str, Any]:
        """Process multiple Excel files with systematic data storage."""
        results = {}
        self.all_records = []  # Reset centralized storage
        
        for i, file_path in enumerate(file_paths):
            try:
                file_id = f"file_{i+1}"
                # Pass the file path to the internal processing method
                result = self._process_single_file(file_path, file_id)
                results[file_id] = result
                
                # Add all records to centralized storage
                if result['status'] == 'success':
                    self.all_records.extend(result['data'])
                
                logger.info(f"Processed {file_path}")
            except Exception as e:
                logger.error(f"Error processing {file_path}: {e}")
                results[f"file_{i+1}"] = {
                    'status': 'error',
                    'filename': os.path.basename(file_path),
                    'error': str(e)
                }
        
        self.processed_files = results
        self._build_rag_system()
        
        logger.info(f"Total records loaded: {len(self.all_records)}")
        return results
    
    def _process_single_file(self, file_path: str, file_id: str) -> Dict[str, Any]:
        """Process a single Excel file with enhanced header detection."""
        filename = os.path.basename(file_path)
        
        try:
            # Open and read the file using pandas
            if file_path.lower().endswith('.csv'):
                df_raw = pd.read_csv(file_path, header=None)
            else:
                df_raw = pd.read_excel(file_path, header=None, engine='openpyxl')
        except FileNotFoundError:
            raise FileNotFoundError(f"File not found: {file_path}")
        except Exception as e:
            raise IOError(f"Error reading file {file_path}: {e}")
        
        raw_data = df_raw.values.tolist()
        
        # Find header using sequential pattern detection
        header_row = self.find_header_with_sequential_pattern(raw_data)
        if header_row is None:
            raise ValueError(f"Could not detect header row in {filename}")
        
        # Extract headers
        headers = []
        header_data = raw_data[header_row]
        for cell in header_data:
            if cell is not None and str(cell).strip():
                headers.append(str(cell).strip())
            else:
                headers.append(f"Column_{len(headers)+1}")
        
        # Process data rows (everything after header)
        data_rows = raw_data[header_row + 1:]
        structured_data = []
        
        for row_idx, row in enumerate(data_rows):
            # Skip completely empty rows
            if not any(cell and str(cell).strip() for cell in row):
                continue
            
            record = {
                '_file_id': file_id, 
                '_filename': filename, 
                '_row_index': row_idx,
                '_original_row': header_row + 1 + row_idx
            }
            
            # Map data to headers
            for col_idx, header in enumerate(headers):
                if col_idx < len(row):
                    value = row[col_idx]
                    # Normalize the value
                    if value is not None:
                        record[header] = str(value).strip()
                    else:
                        record[header] = ''
                else:
                    record[header] = ''
            
            # Only add records that have some meaningful data
            if any(record[key] and str(record[key]).strip() for key in record if not key.startswith('_')):
                structured_data.append(record)
        
        return {
            'status': 'success',
            'filename': filename,
            'file_id': file_id,
            'data': structured_data,
            'columns': headers,
            'header_row': header_row,
            'rows_deleted_above_header': header_row,
            'row_count': len(structured_data),
            'column_count': len(headers)
        }
    
    def _build_rag_system(self):
        """Build RAG system from processed data with error handling."""
        try:
            # Initialize embeddings lazily
            self._initialize_embeddings()
            
            if self.embeddings is None:
                logger.warning("Embeddings not available, using fallback search")
                return
            
            documents = []
            
            for record in self.all_records:
                content_parts = [f"Record from {record['_filename']}:"]
                
                for key, value in record.items():
                    if not key.startswith('_') and value and str(value).strip():
                        content_parts.append(f"{key}: {value}")
                
                content = "\n".join(content_parts)
                
                doc = Document(
                    page_content=content,
                    metadata={
                        'filename': record['_filename'], 
                        'file_id': record['_file_id'],
                        'row_index': record.get('_row_index', 0)
                    }
                )
                documents.append(doc)
            
            if documents:
                self.vector_store = Chroma.from_documents(
                    documents=documents,
                    embedding=self.embeddings,
                    persist_directory="./chroma_db"
                )
                logger.info(f"Built RAG system with {len(documents)} documents")
        except Exception as e:
            logger.error(f"Error building RAG system: {e}")
            self.vector_store = None
    
    def find_exact_record(self, search_id: str, search_name: str = None) -> Dict[str, Any]:
        """Find exact record by ID and optionally name."""
        logger.info(f"Searching for ID: {search_id}, Name: {search_name}")
        
        for record in self.all_records:
            # Check all fields for the ID
            id_found = False
            name_found = search_name is None  # If no name specified, consider it found
            
            for key, value in record.items():
                if key.startswith('_'):
                    continue
                    
                value_str = str(value).strip()
                
                # Check for ID match
                if search_id in value_str:
                    id_found = True
                
                # Check for name match if provided
                if search_name and search_name.lower() in value_str.lower():
                    name_found = True
            
            # If both ID and name found (or name not required), return this record
            if id_found and name_found:
                logger.info(f"Found matching record in {record['_filename']}")
                return record
        
        logger.info(f"No record found for ID: {search_id}")
        return None
    
    def _simple_text_search(self, question: str, k: int = 3) -> List[str]:
        """Fallback text search when vector search is not available."""
        question_lower = question.lower()
        results = []
        
        for record in self.all_records:
            content_parts = [f"Record from {record['_filename']}:"]
            match_score = 0
            
            for key, value in record.items():
                if not key.startswith('_') and value and str(value).strip():
                    value_str = str(value).strip()
                    content_parts.append(f"{key}: {value_str}")
                    
                    # Simple keyword matching
                    if any(word in value_str.lower() for word in question_lower.split()):
                        match_score += 1
            
            if match_score > 0:
                results.append((match_score, "\n".join(content_parts)))
        
        # Sort by match score and return top k
        results.sort(key=lambda x: x[0], reverse=True)
        return [content for _, content in results[:k]]
    
    def query(self, question: str) -> str:
        """Answer user queries with systematic search."""
        if not self.all_records:
            return "No data available. Please process files first."
        
        # Extract ID and name from question
        id_match = re.search(r'\b(\d{4,6})\b', question)
        
        if id_match:
            search_id = id_match.group(1)
            
            # Extract name (everything after the ID and dash)
            name_match = re.search(rf'{search_id}\s*-\s*([^?]+)', question)
            search_name = name_match.group(1).strip() if name_match else None
            
            # Find the exact record
            record = self.find_exact_record(search_id, search_name)
            
            if record:
                # Extract what information is being asked for
                question_lower = question.lower()
                
                # Map common questions to column names
                field_mappings = {
                    'ac no': ['Ac No', 'Account No', 'Account Number'],
                    'organization': ['Organization', 'Org', 'Company'],
                    'designation': ['Designation', 'Position', 'Title'],
                    'job duration': ['Job Duration', 'Duration', 'Service'],
                    'total business': ['Total Business', 'Total Premium', 'Business'],
                    'commission': ['Commission', 'Com'],
                    'pf': ['PF', 'Provident Fund'],
                    'allowance': ['Allowance', 'Allow'],
                    'net pay': ['Net Pay', 'Net', 'Pay'],
                    'tds': ['TDS', 'Tax'],
                    'total premium': ['Total Premium', 'Premium', 'Total PR'],
                    'agent name': ['Agent Name', 'Name'],
                    'sl': ['Sl', 'Serial', 'ID']
                }
                
                # Find what field is being asked about
                requested_field = None
                for key_phrase, possible_columns in field_mappings.items():
                    if key_phrase in question_lower:
                        # Look for matching column in the record
                        for col_name in possible_columns:
                            if col_name in record:
                                requested_field = col_name
                                break
                        if requested_field:
                            break
                
                if requested_field and record.get(requested_field):
                    answer = f"The {requested_field} of {search_id}"
                    if search_name:
                        answer += f" - {search_name}"
                    answer += f" is {record[requested_field]}."
                    
                    # Add source info
                    answer += f"\n\n(Source: {record['_filename']})"
                    return answer
                else:
                    # If specific field not found, return all available info
                    answer = f"Found record for {search_id}"
                    if search_name:
                        answer += f" - {search_name}"
                    answer += f":\n\n"
                    
                    for key, value in record.items():
                        if not key.startswith('_') and value and str(value).strip():
                            answer += f"• {key}: {value}\n"
                    
                    answer += f"\n(Source: {record['_filename']})"
                    
                    if requested_field:
                        answer += f"\n\nNote: The requested field '{requested_field}' was not found in this record."
                    
                    return answer
            else:
                return f"No record found for ID {search_id}" + (f" with name {search_name}" if search_name else "") + " in the loaded data."
        
        # If no ID pattern found, use vector search or fallback
        if self.vector_store:
            try:
                relevant_docs = self.vector_store.similarity_search(question, k=3)
                if relevant_docs:
                    context = "\n\n".join([doc.page_content for doc in relevant_docs])
                    return f"Based on the data, here's what I found:\n\n{context}"
                else:
                    return "No relevant information found for your query."
            except Exception as e:
                logger.error(f"Vector search error: {e}")
                # Fall back to simple text search
                results = self._simple_text_search(question)
                if results:
                    return f"Based on the data, here's what I found:\n\n{results[0]}"
                else:
                    return "No relevant information found for your query."
        else:
            # Use simple text search as fallback
            results = self._simple_text_search(question)
            if results:
                return f"Based on the data, here's what I found:\n\n{results[0]}"
            else:
                return "No relevant information found for your query."
    
    def get_all_data_summary(self) -> str:
        """Get a summary of all loaded records."""
        if not self.all_records:
            return "No data loaded."
        
        summary = f"📊 **Total Records Loaded: {len(self.all_records)}**\n\n"
        
        # Group by file
        file_groups = {}
        for record in self.all_records:
            filename = record['_filename']
            if filename not in file_groups:
                file_groups[filename] = []
            file_groups[filename].append(record)
        
        for filename, records in file_groups.items():
            summary += f"**File: {filename}** ({len(records)} records)\n"
            
            # Show column names
            if records:
                columns = [key for key in records[0].keys() if not key.startswith('_')]
                summary += f"Columns: {', '.join(columns)}\n"
                
                # Show first few IDs
                sample_ids = []
                for record in records[:5]:
                    for key, value in record.items():
                        if not key.startswith('_') and value and str(value).strip().isdigit() and len(str(value).strip()) >= 4:
                            sample_ids.append(str(value).strip())
                            break
                
                if sample_ids:
                    summary += f"Sample IDs: {', '.join(sample_ids)}\n"
            
            summary += "\n"
        
        return summary
    
    def _get_summary(self) -> str:
        """Get detailed summary for debugging."""
        return self.get_all_data_summary()

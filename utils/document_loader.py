from langchain.document_loaders import DirectoryLoader, PyPDFLoader
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain.embeddings import OpenAIEmbeddings
from langchain.vectorstores import Chroma
import os

class DocumentIndexer:
    def __init__(self):
        self.embeddings = OpenAIEmbeddings()
        self.persist_directory = os.getenv("CHROMA_PERSIST_DIRECTORY", "./chroma_db")
        self.vectorstore = None
        self._initialize_vectorstore()
    
    def _initialize_vectorstore(self):
        """Load or create vector store"""
        if os.path.exists(self.persist_directory):
            # Load existing
            self.vectorstore = Chroma(
                persist_directory=self.persist_directory,
                embedding_function=self.embeddings
            )
        else:
            # Create new and index documents
            self.index_documents()
    
    def index_documents(self):
        """Index all ESG documents"""
        # Load regulations
        reg_loader = DirectoryLoader(
            './data/regulations',
            glob="**/*.pdf",
            loader_cls=PyPDFLoader
        )
        
        # Load investor frameworks
        framework_loader = DirectoryLoader(
            './data/frameworks',
            glob="**/*.pdf",
            loader_cls=PyPDFLoader
        )
        
        all_docs = reg_loader.load() + framework_loader.load()
        
        # Split documents
        text_splitter = RecursiveCharacterTextSplitter(
            chunk_size=1000,
            chunk_overlap=200,
            separators=["\n\n", "\n", " ", ""]
        )
        
        texts = text_splitter.split_documents(all_docs)
        
        # Create vector store
        self.vectorstore = Chroma.from_documents(
            documents=texts,
            embedding=self.embeddings,
            persist_directory=self.persist_directory
        )
        
        self.vectorstore.persist()
    
    def search(self, query: str, k: int = 5) -> str:
        """Search for relevant documents"""
        results = self.vectorstore.similarity_search(query, k=k)
        return "\n\n".join([doc.page_content for doc in results])
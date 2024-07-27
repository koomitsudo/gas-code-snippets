import os
import openai
from langchain_community.document_loaders import UnstructuredMarkdownLoader
from langchain.text_splitter import CharacterTextSplitter
from langchain_openai import OpenAIEmbeddings
from langchain_community.vectorstores import FAISS

openai.api_key = os.getenv("OPENAI_API_KEY")

def create_vectorstore():
    loader = UnstructuredMarkdownLoader('service_text.md')
    documents = loader.load()
    text_splitter = CharacterTextSplitter(chunk_size=500, chunk_overlap=100)
    chunks = text_splitter.split_documents(documents)
    embeddings = OpenAIEmbeddings()
    vectorstore = FAISS.from_documents(chunks, embeddings)
    vectorstore.save_local("vectorstore")
    print("Vectorstore が作成されました\n")

if __name__ == "__main__":
    create_vectorstore()
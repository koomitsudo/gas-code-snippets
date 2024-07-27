from langchain_community.document_loaders import UnstructuredMarkdownLoader
from langchain.text_splitter import CharacterTextSplitter

# マークダウンファイルを読み込む
loader = UnstructuredMarkdownLoader("sample.md", encoding="utf-8")

# テキストを指定されたchunk_sizeで分割する
text_splitter = CharacterTextSplitter(chunk_size=10, chunk_overlap=0)
documents = loader.load_and_split(text_splitter=text_splitter)
for doc in documents:
    print(f"Metadata:{doc.metadata['source']} - Content: {doc.page_content}")
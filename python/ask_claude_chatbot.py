import os
import re
import anthropic
from langchain_anthropic import ChatAnthropic
from langchain.prompts import PromptTemplate
from langchain.schema import HumanMessage
from langchain_openai import OpenAIEmbeddings
from langchain_community.vectorstores import FAISS
from langchain_community.document_loaders import UnstructuredMarkdownLoader

anthropic.api_key = os.getenv("ANTHROPIC_API_KEY")

llm = ChatAnthropic(temperature=0, model_name="claude-3-haiku-20240307", anthropic_api_key=os.getenv("ANTHROPIC_API_KEY"), max_tokens_to_sample=400)

pre_template = """
以下の指示を参考に、質問に対して適切に回答してください。
- 回答は、サービス情報に基づいて正確に行うこと。
- 回答は、サービス情報と矛盾しないように注意すること。回答とサービス情報に齟齬がないかを再帰的に自己検証して最も確からしい解釈を提示してください。
- 回答は前置きや余計な言い回しは避けて簡潔に短くまとめてください。

サービス情報:
{processed_description}

上記はサービス情報です。最初にサービスの通常の料金と超過料金、技術的な内容、追加オプション機能の補足事項を書き出してください。その後にオプション費用を確認してください。その他に考慮すべき事項がないかを確認してください。

ユーザーからの質問:
{question}

サービス情報と矛盾が発生しないように質問に対して、明確に回答をしてください。
"""

check_template = """
以下の指示を参考に、回答を検証してください。
- サービス情報を参照して、回答に矛盾や齟齬がないかを徹底的に検証すること。
- 問題がある場合は「問題あり」と返し、修正すべき点を洗い出して修正した回答を提示すること。その際、質問に対する回答部分のみを抜粋して回答すること。
- 問題がない場合は「問題なし」と返し、最初の回答から質問に対する回答部分のみを抜粋して回答すること。

サービス情報:
{processed_description}

質問:
{question}

最初の回答:
{answer}
"""

# ローカルに保存されたベクトルストアを読み込む
def load_vectorstore():
    embeddings = OpenAIEmbeddings(openai_api_key=os.getenv("OPENAI_API_KEY"))
    vectorstore = FAISS.load_local("vectorstore", embeddings, allow_dangerous_deserialization=True)
    return vectorstore

def generate_answer(processed_description, question, pre_template):
    # PromptTemplateを使用して、pre_templateとprocessed_description、questionを組み合わせてプロンプトを作成
    prompt = PromptTemplate(
        input_variables=["processed_description", "question"],
        template=pre_template,
    )
    # LLMへ質問などを渡して回答を生成
    human_message = HumanMessage(content=prompt.format(processed_description=processed_description, question=question))
    answer = llm.invoke([human_message])
    print(f"\n''''''''''\ngenerate_answer関数の出力 : {answer.content}\n\n'''''''''\n")
    return answer.content

# 1度目のLLMの回答を再検証
def verify_answer(answer, processed_description, question, check_template):
    # PromptTemplateを使用して、check_templateとanswer、processed_description、questionを組み合わせてプロンプトを作成
    prompt = PromptTemplate(
        input_variables=["answer", "processed_description", "question"],
        template=check_template,
    )    
    # LLMへ質問などを渡して回答を生成
    verification_result = llm.invoke([HumanMessage(content=prompt.format(answer=answer, processed_description=processed_description, question=question))])
    print(f"\n''''''''''\nverify_answer関数の出力 : {verification_result.content}\n\n'''''''''\n")
    
    # LLMの回答に余計な文字列(ここでは修正*など)が記載されるため正規表現で検知して除去
    pattern = r"(修正した回答:|修正後の回答:|修正版:)\n(.*)"
    try:
        match = re.search(pattern, verification_result.content, re.DOTALL)
        if match:
            return match.group(2).strip()
        else:
            return answer
    except AttributeError as e:
        # match オブジェクトがNoneである場合
        print(f"AttributeErrorが発生: {e}")
        # 元の回答をそのまま返す。
        return answer

def get_answer(question):
    vectorstore = load_vectorstore()
    docs = vectorstore.similarity_search(question)
    processed_description = "\n".join([doc.page_content for doc in docs])
    initial_answer = generate_answer(processed_description, question, pre_template)
    answer = verify_answer(initial_answer, processed_description, question, check_template)
    formatted_answer = f"お問い合わせありがとうございます。\n\n{answer}\n\n何卒よろしくお願いいたします。"
    return formatted_answer

if __name__ == "__main__":
    question = "私は今月からエントリープランを契約しており、追加有償機能に契約している。利用回数は89回を使用している。私の今月の料金はいくらだろうか？"
    answer = get_answer(question)
    print(question)
    print(f"\n''''''''''\n[回答]\n {answer}\n\n'''''''''\n")	
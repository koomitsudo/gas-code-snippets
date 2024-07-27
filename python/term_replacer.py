import os
import re
import json
import pandas as pd

PATH = os.getcwd()
CONFIG = os.path.join(PATH, "config.json") 

def main():
    # configファイルを読み込む
    with open(CONFIG, 'r', encoding='utf-8') as f:
        json_load = json.load(f)
    READ_FILE = json_load['path']['readFile']
    WRITE_FILE = json_load['path']['writeFile']
    PATTERN_FILE = json_load['path']['patternFile']

    txt = read_text(READ_FILE)
    create_text(WRITE_FILE, txt)
    load_csv(WRITE_FILE, PATTERN_FILE)
    result = read_text(WRITE_FILE)
    print(result)

def read_text(path):
    # ファイルからテキストを読み込む
    with open(path, 'r', encoding="utf-8") as f:  
        txt = f.read()
        return txt

def create_text(write_file, txt):
    with open(write_file, mode='w', encoding="utf-8") as ct:
        ct.write(txt)

# 置換されたテキストでファイルを上書きする
def overwrite_text(write_file, replaced_text):
    with open(write_file, mode='w', encoding="utf-8") as ot:
        ot.write(replaced_text)

# CSVファイルから置換パターンを読み込み、テキストを置換する
def load_csv(write_file, pattern_file):
    df = pd.read_csv(pattern_file)
    for i in range(len(df)):
        new_txt = read_text(write_file)
        before = df['before'][i]
        after = df['after'][i]
        replaced_text = re.sub(before, after, new_txt)
        overwrite_text(write_file, replaced_text)

if __name__ == "__main__":
    main()

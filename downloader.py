import subprocess
import sys

# 必要なライブラリのリスト
libraries = [
    'python-Levenshtein',   # Levenshtein
    'numpy',                # numpy
    'pandas',               # pandas
    'janome',               # janome (jenomeの代替として)
    'pykakasi',             # pykakasi
    'tkinter',              # tkinter (画面生成)
    'mecab-python3',        # MeCab
    'openpyxl',             # Excelファイルの読み込みに必要
    "unidic-lite",
]

# 各ライブラリを順次インストール
for library in libraries:
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", library])
        print(f"{library} のインストールが成功しました")
    except subprocess.CalledProcessError:
        print(f"{library} のインストールに失敗しました")

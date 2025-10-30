
"""
1. 「送る」メニューのフォルダを開く
ファイル名を指定して実行を開きます。
Winキー + Rキー を押します。
以下のパスを入力してEnterを押します。
code
Code
shell:sendto
これにより、「送る」メニューに表示されるショートカットが格納されているフォルダが開きます。
2. exeプログラムのショートカットを作成する
先ほど開いた「SendTo」フォルダの中に、あなたのPython exeプログラムのショートカットを作成します。
あなたのexeファイルがある場所を開きます。
exeファイルを右クリックし、「送る」→「デスクトップ (ショートカットを作成)」を選択します。
作成されたデスクトップ上のショートカットを、「SendTo」フォルダにドラッグ＆ドロップで移動します。
または、exeファイルを右クリックし、Ctrlキーを押しながら「SendTo」フォルダにドラッグ＆ドロップすると、ショートカットが作成されます。

3. ショートカットの動作確認
適当なテキストファイルを用意します。
そのテキストファイルを右クリックし、「送る」メニューにあなたのexeプログラムの名前が表示されていることを確認します。
表示されたプログラム名をクリックして、正しくexeプログラムにファイルが渡されるか確認します。

デスクトップに作成したショートカットにファイルをD＆DしてもOKです。 


Pythonプログラム側でのファイルの受け取り方
あなたのPythonプログラム（exe化したもの）では、コマンドライン引数として渡されたファイルパスを受け取る必要があります。sysモジュールを使用します。
"""

import sys
import os

def process_file(file_path):
    if not os.path.exists(file_path):
        print(f"Error: File not found at {file_path}")
        return

    print(f"Processing file: {file_path}")
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
            print("File content (first 200 chars):")
            print(content[:200]) # 例として、最初の200文字を表示
            # ここにファイルの処理ロジックを記述します
            # 例: ファイルの内容を別のファイルに書き出す、解析する など
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")

if __name__ == "__main__":
    if len(sys.argv) > 1:
        # sys.argv[0] はスクリプト自身のパスなので、
        # sys.argv[1] 以降が渡されたファイルのパスになります。
        for file_path in sys.argv[1:]:
            print(file_path)
            process_file(file_path)
    else:
        print("Usage: Drag and drop a file onto this executable, or send it via 'Send To' menu.")
        print("No file path provided as an argument.")

    # プログラムがすぐに閉じないように、ユーザーがキーを押すまで待つ
    input("Press Enter to exit...")
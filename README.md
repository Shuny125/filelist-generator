# filelist-generator
ローカルファイルを元にファイルリスト（URL リスト）を作成する機能です。

## Dependency
動作確認済み環境  
macOS Sierra 10.13.4  
Python 3.5.0  

## Setup
### Python
https://www.python.org/downloads/windows/  
3系の最新バージョンをインストール  
→「Download Windows x86-64 executable installer」

ターミナルを開いて、以下のコマンドを実行。
```
$ python —V
```
バージョンが表示されたらOK。

### pip
ターミナルを開いて、以下のコマンドを実行。
```
$ easy_insatll pip
```

### Python ライブラリインストール
ターミナルを開いて「requirements.txt」と同じディレクトリにいる状態で、以下のコマンドを実行。
```
$ pip install -r requirements.txt
```

## Usage
### filelist.xlsx
基本的には何も編集しないで OK です。  
行や列が足りない場合は、編集してください。  
※列を増減する場合は、コードの編集も必要になります

出力される情報は以下の通りです。
- ページID（独自に生成）
- タイトル
- パス
- keywords
- discription

### filelist.py
サイトのルートディレクトリに配置してください。  
※filelist.xlsx も併せて同階層に配置

### 実行
ターミナルを開いて、以下のコマンドを実行。
```
$ python filelist.py
```

## Licence
This software is released under the MIT License, see LICENSE.

## Author
[Shuny125](https://github.com/Shuny125)

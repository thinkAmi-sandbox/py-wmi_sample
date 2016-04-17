# py-wmi_sample

## セットアップ
```
# 任意のGit用ディレクトリへ移動
>cd path\to\dir

# GitHubからカレントディレクトリへclone
path\to\dir>git clone https://github.com/thinkAmi-sandbox/py-wmi_sample.git

# virtualenv環境の作成とactivate
# *Python3.5は、`c:\python35-32\`の下にインストール
path\to\dir>virtualenv -p c:\python35-32\python.exe env
path\to\dir>env\Scripts\activate

# requirements.txtよりインストール
(env)path\to\dir>pip install -r requirements.txt

# 実行
(env)path\to\dir>python blog_example.py
# 実行 => 結果イメージは、wmi_result.txt
(env)path\to\dir>python runner.py
```

　  
## テスト環境

- Windows10
- Python 3.5.1
- pywin32 220
- WIM 1.4.9

　  
## 関係するブログ
[Python3 + WMI + ospp.vbsで、Windows端末やMS Office情報を取得してみた - メモ的な思考的な](http://thinkami.hatenablog.com/entry/2016/04/07/233504)

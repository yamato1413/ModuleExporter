# ModuleExporter



ドラッグアンドドロップでエクセルファイルのモジュールがエクスポートできるVBScript。
![image](https://user-images.githubusercontent.com/69558300/145819386-9bf8d06b-e476-40fd-b6f4-bf2c89a3f388.png)

動画のサンプルは以下より
<blockquote class="twitter-tweet"><p lang="ja" dir="ltr">エクセルのモジュールをエクスポートするVBScript<br>コマンドラインだとオプション引数で種類を選べる <a href="https://t.co/Oi7trumFkT">pic.twitter.com/Oi7trumFkT</a></p>&mdash; やまと💻アイコン変えましたん (@yamato_1413) <a href="https://twitter.com/yamato_1413/status/1470354255190294529?ref_src=twsrc%5Etfw">December 13, 2021</a></blockquote>

うっかり関係ないファイルが混ざっても大丈夫。

コマンドラインで実行するときには，オプション引数を渡せばエクスポートするモジュールの種類も選べます。

```
-a 全部
-s 標準モジュール
-c クラスモジュール
-u フォームモジュール
-o オブジェクトモジュール
```

```
export.vbs sample.xlsm                #省略すると全部
export.vbs -s sample.xlsm
export.vbs -sc sample.xlsm            #複数種類を同時に選択可
export.vbs sample.xlsm -o             #オプション引数の順番はどこでもOK
export.vbs sample.xlsm sample2.xlsm   #エクセルファイルはいくつ渡してもOK
```

# 注意点
「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」を強制的にONにしているので，異常終了するとONのままになってしまいます。
お気を付けて。

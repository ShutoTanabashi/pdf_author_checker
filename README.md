# PDFの製作者を確認するプログラム

製作者：棚橋秀斗[tanabashi@s.okayama-u.ac.jp]

## 注意事項

本プログラムを実行あるいは流用したことにより生じる一切の事象について製作者は責任を負いません。
自己責任でお願いいたします。

## 実行前の注意事項

本プログラムは以下の外部パッケージに依存しています。
適宜インストールしてから実行してください。

* numpy
* pandas
* pypdf
* pdf2image

特に`pdf2image`パッケージの使用にはpoppler（PDF関連のライブラリ）を別途インストールする必要があります。
popplerのインストール方法の詳細は
[https://github.com/Belval/pdf2image](https://github.com/Belval/pdf2image)
を参照してください

## 実行方法

本リポジトリにはテスト用のデータが`testdata`ディレクトリに格納されています。
ターミナルエミュレータで

```zsh
python checkauthor.py testdata
```

と実行してください。
製作者情報をまとめた`metadata.xlsx`というファイルが出力されます。

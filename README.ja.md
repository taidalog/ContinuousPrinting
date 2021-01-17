# ContinuousPrinting

[English README](README.md)

1. [概要](#概要)
1. [インストール](#インストール)
1. [使い方](#使い方)
1. [今後の機能](#今後の機能)

## 概要
Excelで連続印刷を行うためのVBAスクリプト（マクロ）です。  
連続印刷のよくあるパターンとして、
- A1 セルに 1 と入力して、VLOOKUP()関数が動いて、シートを印刷する
- A1 セルに 2 と入力して、VLOOKUP()関数が動いて、シートを印刷する
- A1 セルに 3 と入力して、VLOOKUP()関数が動いて、シートを印刷する
- 以下繰り返し

というものがあると思います。このマクロを使うと、そのような印刷が簡単に行えます。1 から 3 まで印刷する場合、'1-3' と入力して、数値を入れるセルをクリックすることで連続印刷が行えます。

## インストール
以下の手順に従って操作してください。 
1. 'ContinuousPrinting.bas' をコンピュータに保存する。
1. 連続印刷を行いたい Excel ファイルに 'ContinuousPrinting.bas' をインポートする。  
拡張子が '**.xlsm**' か '**.xlam**' （または '.xls' ）の Excel ファイルにインポートしてください。
1. 'Alt + F8' を押して '**cp**' と入力して実行してください（**C**ontinuous**P**rinting の略です）。

このマクロをコンテキストメニュー（右クリックメニュー）に追加することもできます。
1. VBE を開く（'Alt + F11' を押します）。
1. 'VBAProject (目的の Excel ファイル名)' -> 'Misrosoft Excel Object' -> 'ThisWorkbook' を開く。
1. 以下のコードを追加する。  
既に `Private Sub Workbook_Open()` が存在している場合は、`End Sub`の前に `Call AddToContextMenu_ContinuousPrinting` だけを追加してください。
```VB
Private Sub Workbook_Open()
    Call AddToContextMenu_ContinuousPrinting
End Sub
```
4. Excel ファイルを保存して、もう一度開く。
4. そうすると、'ContinuousPrinting' がメニューに現れているはずです。  
'ContinuousPrinting' は、この Excel ファイルを開いている間のみ現れます。

## 使い方
1. 印刷する番号を入力する。
1. 番号を入れるセルをクリックする。**選択できるセルは1つのみです**。
1. 確認画面で「はい」をクリックする。

## 今後の機能
- [ ] 番号の代わりに文字を入力できるようにする
- [ ] セルに入力してある番号を使えるようにする

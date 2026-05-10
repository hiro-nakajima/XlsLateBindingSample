# ExcelLateBindingApp

Excel を COM の遅延バインディングで操作する Windows Forms サンプルアプリです。
元のコードは 2009 年に作成したものを、現在の .NET SDK / C# Dev Kit で動かせるように調整しています。

`Microsoft.Office.Interop.Excel` への参照を使わず、`Type.GetTypeFromProgID("Excel.Application")` とリフレクションで Excel を起動・操作します。

## 動作内容

- デスクトップの `test.xls` を開く
- ファイルが存在しない場合は新規ブックを作成する
- 1 シート目に九九表を書き込む
- セルの色や罫線、列幅、行高を設定する
- 「保存する」にチェックが入っている場合はブックを保存する

## 必要環境

- Windows
- .NET SDK 8 以降
- Microsoft Excel

## ビルド方法

```powershell
dotnet build WindowsApplication1.sln
```

## 実行方法

```powershell
dotnet run --project WindowsApplication1\WindowsApplication1.csproj
```

## 補足

このアプリは Excel がインストールされている Windows 環境での実行を前提にしています。
COM の遅延バインディングを使っているため、Excel Interop アセンブリへの参照は不要です。

## 参考

- ColorIndex property (Microsoft Learn): https://learn.microsoft.com/ja-jp/office/vba/api/excel.colorindex
- Excel 56 ColorIndex colors: https://www.excelsupersite.com/what-are-the-56-colorindex-colors-in-excel/
- Late binding example: http://d.hatena.ne.jp/matsumoto0325/20060822/1156245089

using System;
using System.IO;
using System.Text;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;

namespace WindowsApplication1
{
    /// <summary>
    /// LateBindingを使用したExcelクラス
    /// </summary>
    public class ExcelLateBindingUtility : IDisposable
    {
        object _apl = null;
        object _workbook = null;
        object _worksheet = null;
        List<object> _disposeList = new List<object>();
        bool _isQuit = true;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public ExcelLateBindingUtility()
        {
            _apl = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
            _disposeList.Add(_apl);
            _apl.GetType().InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, _apl, new object[] { false });
            _apl.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, _apl, new object[] { false });

        }

        /// <summary>
        /// Dispose/ReleaseComObjectを行う
        /// </summary>
        public void Dispose()
        {
            _apl.GetType().InvokeMember("CutCopyMode", BindingFlags.SetProperty, null, _apl, new object[] { 0 });
            if (_isQuit) { _apl.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, _apl, null); }
            foreach (object o in _disposeList)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
        }

        /// <summary>
        /// 終了設定
        /// </summary>
        public bool IsQuit
        {
            get { return _isQuit; }
            set { _isQuit = value; }
        }

        /// <summary>
        /// 表示設定
        /// </summary>
        public bool Visible
        {
            get { return (bool)_apl.GetType().InvokeMember("Visible", BindingFlags.GetProperty, null, _apl, null); }
            set { _apl.GetType().InvokeMember("Visible", BindingFlags.SetProperty, null, _apl, new object[] { value }); }
        }

        /// <summary>
        /// DisplayAlertsの設定
        /// </summary>
        public bool DisplayAlerts
        {
            get { return (bool)_apl.GetType().InvokeMember("DisplayAlerts", BindingFlags.GetProperty, null, _apl, null); }
            set { _apl.GetType().InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, _apl, new object[] { value }); }
        }

        /// <summary>
        /// デフォルトファイルパスの設定
        /// </summary>
        public string DefaultFilePath
        {
            set { _apl.GetType().InvokeMember("DefaultFilePath", BindingFlags.SetProperty, null, _apl, new object[] { value }); }
            get { return _apl.GetType().InvokeMember("DefaultFilePath", BindingFlags.GetProperty, null, _apl, null).ToString(); }
        }

        /// <summary>
        /// バージョンを取得する
        /// </summary>
        public string Version
        {
            get { return _apl.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, _apl, null).ToString(); }
        }

        /// <summary>
        /// シート名の設定
        /// </summary>
        public string SheetName
        {
            set { _worksheet.GetType().InvokeMember("Name", BindingFlags.SetProperty, null, _worksheet, new object[] { value }); }
            get { return _worksheet.GetType().InvokeMember("Name", BindingFlags.GetProperty, null, _worksheet, null).ToString(); }
        }

        /// <summary>
        /// ブックを開く
        /// </summary>
        /// <param name="filePath">ファイルパス</param>
        public void Open(string filePath)
        {
            object workbooks = _apl.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, _apl, null);
            _disposeList.Add(workbooks);
            _workbook = workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, workbooks, new object[] { filePath });
            _disposeList.Add(_workbook);
            object sheets = _workbook.GetType().InvokeMember("Sheets", BindingFlags.GetProperty, null, _workbook, null);
            _disposeList.Add(sheets);
            _worksheet = sheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, sheets, new object[] { 1 });
            _disposeList.Add(_worksheet);
        }

        /// <summary>
        /// テンプレートシートを追加する
        /// </summary>
        /// <param name="templateFilePath">テンプレートファイル</param>
        public void Add(string templateFilePath)
        {
            object workbooks = _apl.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, _apl, null);
            _disposeList.Add(workbooks);
            _workbook = workbooks.GetType().InvokeMember("Add", BindingFlags.InvokeMethod, null, workbooks, new object[] { templateFilePath });
            _disposeList.Add(_workbook);
            object sheets = _workbook.GetType().InvokeMember("Sheets", BindingFlags.GetProperty, null, _workbook, null);
            _disposeList.Add(sheets);
            _worksheet = sheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, sheets, new object[] { 1 });
            _disposeList.Add(_worksheet);
        }

        /// <summary>
        /// 現在のブックに指定したブックの１シート目をコピーする
        /// </summary>
        /// <param name="sourceFilePath">コピー元ブック</param>
        public void CopySheet(string sourceFilePath)
        {
            object workbooks = _apl.GetType().InvokeMember("Workbooks", BindingFlags.GetProperty, null, _apl, null);
            _disposeList.Add(workbooks);
            object _srcWorkbook = workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, workbooks, new object[] { sourceFilePath });
            _disposeList.Add(_srcWorkbook);
            object sheets = _workbook.GetType().InvokeMember("Sheets", BindingFlags.GetProperty, null, _srcWorkbook, null);
            _disposeList.Add(sheets);
            object _srcWorksheet = sheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, sheets, new object[] { 1 });
            _disposeList.Add(_srcWorksheet);

            _srcWorksheet.GetType().InvokeMember("Copy", BindingFlags.InvokeMethod, null, _srcWorksheet, new object[] { _worksheet });
            _srcWorkbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, _srcWorkbook, null);

            // コピーしたシートにItem設定
            object osheets = _workbook.GetType().InvokeMember("Sheets", BindingFlags.GetProperty, null, _workbook, null);
            _disposeList.Add(osheets);
            _worksheet = osheets.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, osheets, new object[] { 1 });
        }

        
        /// <summary>
        /// セルに値をセットする
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="col">列</param>
        /// <param name="str">値</param>
        public void SetCellValue(int row, int col, string str)
        {
            object cells = _worksheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, _worksheet, null);
            _disposeList.Add(cells);
            cells.GetType().InvokeMember("Item", BindingFlags.SetProperty, null, cells, new object[] { row, col, str });
        }

        /// <summary>
        /// セルの値を取得する
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="col">列</param>
        /// <returns>値（nullの場合はEmptyを返却）</returns>
        public string GetCellValue(int row, int col)
        {
            object cells = _worksheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, _worksheet, null);
            _disposeList.Add(cells);
            object range = cells.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, cells, new object[] { row, col });
            _disposeList.Add(range);

            object val = range.GetType().InvokeMember("Value", BindingFlags.GetProperty, null, range, null);
            return val != null ? val.ToString() : string.Empty;
        }

        /// <summary>
        /// セルオブジェクトを取得する
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public object GetCell(int row, int col)
        {
            object cells = _worksheet.GetType().InvokeMember("Cells", BindingFlags.GetProperty, null, _worksheet, null);
            _disposeList.Add(cells);
            object range = cells.GetType().InvokeMember("Item", BindingFlags.GetProperty, null, cells, new object[] { row, col });
            _disposeList.Add(range);
            return range;
        }

        /// <summary>
        /// 行オブジェクトを取得する
        /// </summary>
        /// <param name="row">行</param>
        /// <returns>行オブジェクト</returns>
        public object GetRow(int row)
        {
            object result = _worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null,
                _worksheet, new object[] { GetCell(row, 1), GetCell(row, 256) });
            _disposeList.Add(result);
            return result;
        }

        /// <summary>
        /// セルにカラーをセットする（Range指定可能）
        /// </summary>
        /// <param name="startRow">開始行</param>
        /// <param name="startCol">開始列</param>
        /// <param name="endRow">終了行</param>
        /// <param name="endCol">終了列</param>
        /// <param name="colorIndex">カラーインデックス</param>
        public void SetColor(int startRow, int startCol, int endRow, int endCol, int colorIndex)
        {
            object range = _worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null,
                _worksheet, new object[] { GetCell(startRow, startCol), GetCell(endRow, endCol) });
            _disposeList.Add(range);

            object interior = range.GetType().InvokeMember("Interior", BindingFlags.GetProperty, null, range, null);
            _disposeList.Add(interior);

            interior.GetType().InvokeMember("ColorIndex", BindingFlags.SetProperty, null, interior, new object[] { colorIndex });
        }

        /// <summary>
        /// セルのカラーを取得する
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="col">列</param>
        /// <returns>カラーインデックス</returns>
        public int GetColor(int row, int col)
        {
            object range = _worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null,
                _worksheet, new object[] { GetCell(row, col), GetCell(row, col) });
            _disposeList.Add(range);

            object interior = range.GetType().InvokeMember("Interior", BindingFlags.GetProperty, null, range, null);
            _disposeList.Add(interior);

            object color = interior.GetType().InvokeMember("ColorIndex", BindingFlags.GetProperty, null, interior, null);
            return (int)color;
        }

        /// <summary>
        /// Rangeを取得する
        /// </summary>
        /// <param name="startRow">開始行</param>
        /// <param name="startCol">開始列</param>
        /// <param name="endRow">終了行</param>
        /// <param name="endCol">終了列</param>
        /// <returns></returns>
        public object GetRange(int startRow, int startCol, int endRow, int endCol)
        {
            object range = _worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null,
                _worksheet, new object[] { GetCell(startRow, startCol), GetCell(endRow, endCol) });
            _disposeList.Add(range);
            return range;
        }

        /// <summary>
        /// マップ作成の為、列幅、行の高さ、囲み罫線を整える
        /// </summary>
        /// <param name="startRow">開始行</param>
        /// <param name="startCol">開始列</param>
        /// <param name="endRow">終了行</param>
        /// <param name="endCol">終了列</param>
        /// <param name="colWidth">列幅</param>
        /// <param name="rowHeight">行の高さ</param>
        public void PrepareMap(int startRow, int startCol, int endRow, int endCol, int colWidth, int rowHeight)
        {
            object range = _worksheet.GetType().InvokeMember("Range", BindingFlags.GetProperty, null,
                _worksheet, new object[] { GetCell(startRow, startCol), GetCell(endRow, endCol) });
            _disposeList.Add(range);

            range.GetType().InvokeMember("ColumnWidth", BindingFlags.SetProperty, null, range, new object[] { colWidth });
            range.GetType().InvokeMember("RowHeight", BindingFlags.SetProperty, null, range, new object[] { rowHeight });

            // 第4引数まで…LineStyle, XlBorderWeight Weight, XlColorIndex ColorIndex, Color、すべて省略可能
            range.GetType().InvokeMember("BorderAround", BindingFlags.InvokeMethod, null, range,
                new object[] { XlLineStyle.xlContinuous, XlBorderWeight.xlThick });
        }
        // 下記罫線の定数は、office xp にて確認
        // 罫線のスタイル
        public enum XlLineStyle : int
        {
            xlContinuous = 1,         // 実線
            xlDash = 2,               // 破線
            xlDot = 3,                // 一点鎖線
            xlDashDot = 4,            // 二点鎖線
            xlDashDotDot = 5,         // 点線
            //xlDouble = 6,           // 二重線
            //xlSlantDashDot = 7,     // 斜線
            //xlLineStyleNone = 8,    // 線なし
        }
        // 罫線の太さ
        public enum XlBorderWeight : int
        {
            xlHairline = 1,         // ほそい
            xlThin = 2,             // うすい
            xlMedium = 3,           // ふつう
            xlThick = 4,            // 太 
        }

        /// <summary>
        /// セルをセレクトする
        /// </summary>
        /// <param name="row"></param>
        /// <param name="col"></param>
        public void SelectCell(int row, int col)
        {
            object cell = GetCell(row, col);
            cell.GetType().InvokeMember("Select", BindingFlags.InvokeMethod, null, cell, null);
        }

        
        /// <summary>
        /// ブックのファイルパスを指定して保存する
        /// </summary>
        /// <param name="filePath"></param>
        public void SaveAs(string filePath)
        {
            int fileFormat = GetFileFormat(filePath);
            if (fileFormat == 0)
            {
                _workbook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, _workbook, new object[] { filePath });
            }
            else
            {
                _workbook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, _workbook, new object[] { filePath, fileFormat });
            }
        }

        private static int GetFileFormat(string filePath)
        {
            switch (Path.GetExtension(filePath).ToLowerInvariant())
            {
                case ".xls":
                    return 56; // xlExcel8
                case ".xlsx":
                    return 51; // xlOpenXMLWorkbook
                case ".xlsm":
                    return 52; // xlOpenXMLWorkbookMacroEnabled
                default:
                    return 0;
            }
        }

        /// <summary>
        /// ブックを保存する
        /// </summary>
        public void Save()
        {
            _workbook.GetType().InvokeMember("Save", BindingFlags.InvokeMethod, null, _workbook, null);
        }

        /// <summary>
        /// ブックを閉じる
        /// </summary>
        public void Close()
        {
            _workbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, _workbook, null);
        }

        /// <summary>
        /// Rangeをコピーするペーストタイプ
        /// </summary>
        public enum XlPasteType : int
        {
            xlPasteValues = -4163,
            xlPasteFormats = -4122,
            xlPasteAll = -4104,
        }

        /// <summary>
        /// range_fromからrange_toにコピーする
        /// </summary>
        /// <param name="range_from">コピー元Rangeオブジェクト</param>
        /// <param name="range_to">コピー先Rangeオブジェクト</param>
        /// <param name="pasteType">ペーストタイプ</param>
        public static void CopyRange(object range_from, object range_to, XlPasteType pasteType)
        {
            range_from.GetType().InvokeMember("Copy", BindingFlags.InvokeMethod, null, range_from, null);
            object[] args = new object[4];
            args[0] = (int)pasteType;
            args[1] = -4142;//Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone
            args[2] = false;
            args[3] = false;
            range_to.GetType().InvokeMember("PasteSpecial", BindingFlags.InvokeMethod, null, range_to, args);
        }
    }
}

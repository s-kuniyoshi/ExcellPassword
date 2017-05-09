using System;
using System.IO;

namespace ExcelPassword.Models
{
    class Excel : IDisposable
    {
        public string FileName { get; set; }
        public string Password { get; set; }
        public string ExcelPath { get; set; }
        public Microsoft.Office.Interop.Excel.Application ExcelFile { get; set; }
        public Microsoft.Office.Interop.Excel.Workbook Workbook { get; set; }

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="excelPath">Excelファイルのパス</param>
        public Excel(string excelPath)
        {
            ExcelPath = excelPath;
            FileName = Path.GetFileName(ExcelPath);
        }

        /// <summary>
        /// パスからExcelファイルを読み込む
        /// </summary>
        /// <returns>読み込み正常終了:True　読み込み失敗:False</returns>
        public void read()
        {
            if (ExcelPath.Length != 0)
            {

                ExcelFile = new Microsoft.Office.Interop.Excel.Application();

                //エクセルを非表示
                ExcelFile.Visible = false;

                //警告を非表示
                ExcelFile.DisplayAlerts = false;

                //エクセルファイルのオープンと
                //ワークブックをの作成
                Workbook = ExcelFile.Workbooks.Open(ExcelPath,
                    0,
                    Type.Missing, Type.Missing, "", "",//パスワードを空としておく
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }

        }

        /// <summary>
        /// Excelファイルをクローズ
        /// </summary>
        public void close()
        {
            //ワークブックを閉じる
            Workbook.Close();
            //エクセルを閉じる
            ExcelFile.Quit();
        }

        /// <summary>
        /// Excelを保存する
        /// </summary>
        public void save()
        {
            Workbook.SaveAs(ExcelPath);
        }

        /// <summary>
        /// Excelをパスワード付きで保存する
        /// </summary>
        /// <returns>保存成功:True　保存失敗:False</returns>
        public void saveWithPassword()
        {
            if (ExcelPath.Length != 0)
            {

                if (Password.Length != 0)
                {
                    Workbook.Password = Password;
                    Workbook.SaveAs(ExcelPath);
                }
            }
        }

        /// <summary>
        /// デストラクタ
        /// </summary>
        public void Dispose()
        {
            if (Workbook != null)
            {
                //ワークブックを閉じる
                Workbook.Close();
                //エクセルを閉じる
                ExcelFile.Quit();
            }
        }
    }
}

using ExcelPassword.Models;
using Livet;
using Livet.Commands;
using Livet.EventListeners;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelPassword.ViewModels
{
    public class MainWindowViewModel : ViewModel
    {
        //ProgressBarのデリゲートを定義
        private delegate void ProgressBarDelegate();
        private delegate void SetProgressBarDelegate(int min, int max);
        private delegate void ChangeExecuteDelegate();
        private Model model = null;
        private DailyLoggingTraceListener dailyLogging = null;
        private TextBoxTraceListener textboxLogging = null;

        public void Initialize()
        {
            var disp = DispatcherHelper.UIDispatcher;

            model = new Model();
            var eventListener = new PropertyChangedEventListener(model, (sender, e) => RaisePropertyChanged(e.PropertyName));
            CompositeDisposable.Add(eventListener);

            model.folderPath = "";
            model.selectFilePath = "";
            model.controllLock = true;
            model.doing = false;
            model.executeText = "実行";
            model.blnError = false;

            //ログを設定
            dailyLogging = new DailyLoggingTraceListener()
            {
                OutputDirectory = Application.StartupPath + @"\log\",   // 既定値 Application.StartupPath
                FileNameFormat = "{0:yyyyMMdd}_{1}.log",                // 既定値 {0:yyyyMMdd}_{1}.txt
                DatetimeFormat = "{0:MM/dd HH:mm}",                     // 既定値 {0:MM/dd HH:mm:ss}
            };
            Trace.Listeners.Add(dailyLogging);
            textboxLogging = new TextBoxTraceListener();
            Trace.Listeners.Add(textboxLogging);
            var eventListener2 = new PropertyChangedEventListener(textboxLogging, (sender, e) => RaisePropertyChanged(e.PropertyName));
            CompositeDisposable.Add(eventListener2);
        }

        public string logging
        {
            get { return textboxLogging.Trace; }
        }

        #region FolderPath変更通知プロパティ
        public string FolderPath
        {
            get { return model.folderPath; }
            set { model.folderPath = value; }
        }
        #endregion

        #region blnError変更通知プロパティ
        public Boolean blnError
        {
            get { return model.blnError; }
            set { model.blnError = value; }
        }
        #endregion

        #region progressValue変更通知プロパティ
        public int progressValue
        {
            get { return model.progressValue; }
            set { model.progressValue = value; }
        }
        #endregion

        #region progressMinimum変更通知プロパティ
        public int progressMinimum
        {
            get { return model.progressMinimum; }
            set { model.progressMinimum = value; }
        }
        #endregion

        #region progressMaximum変更通知プロパティ
        public int progressMaximum
        {
            get { return model.progressMaximum; }
            set { model.progressMaximum = value; }
        }
        #endregion

        #region controllLock変更通知プロパティ
        public bool controllLock
        {
            get { return model.controllLock; }
            set { model.controllLock = value; }
        }
        #endregion

        #region doing変更通知プロパティ
        public bool doing
        {
            get { return model.doing; }
            set { model.doing = value; }
        }
        #endregion

        #region executeText変更通知プロパティ
        public string executeText
        {
            get { return model.executeText; }
            set { model.executeText = value; }
        }
        #endregion

        #region FolderPath変更通知プロパティ
        public string SelectFilePath
        {
            get { return model.selectFilePath; }
            set { model.selectFilePath = value; }
        }
        #endregion

        /// <summary>
        /// ファイル選択ボタン押下時のCommandoを定義
        /// </summary>
        #region SelectFileCommand
        private ViewModelCommand _SelectFileCommand;

        public ViewModelCommand SelectFileCommand
        {
            get
            {
                if (_SelectFileCommand == null)
                {
                    _SelectFileCommand = new ViewModelCommand(SelectFile);
                }
                return _SelectFileCommand;
            }
        }

        #region ExecuteCommand
        private ViewModelCommand _ExecuteCommand;

        /// <summary>
        /// 実行選択ボタン押下時のCommandを定義 
        /// </summary>
        public ViewModelCommand ExecuteCommand
        {
            get
            {
                if (_ExecuteCommand == null)
                {
                    _ExecuteCommand = new ViewModelCommand(Execute_Button);
                }
                return _ExecuteCommand;
            }
        }
        #region OpenFolderDialogCommand
        private ViewModelCommand _OpenFolderDialogCommand;

        /// <summary>
        /// フォルダ選択ボタン押下時のCommandを定義 
        /// </summary>
        public ViewModelCommand OpenFolderDialogCommand
        {
            get
            {
                if (_OpenFolderDialogCommand == null)
                {
                    _OpenFolderDialogCommand = new ViewModelCommand(OpenFolderDialog);
                }
                return _OpenFolderDialogCommand;
            }
        }

        /// <summary>
        /// フォルダボタン選択時のイベント定義
        /// </summary>
        public void OpenFolderDialog()
        {
            //FolderBrowserDialogクラスのインスタンスを作成
            var fbd = new System.Windows.Forms.FolderBrowserDialog();

            //上部に表示する説明テキストを指定する
            fbd.Description = "フォルダを指定してください。";
            //ルートフォルダを指定する
            //デフォルトでDesktop
            fbd.RootFolder = Environment.SpecialFolder.Desktop;
            //最初に選択するフォルダを指定する
            //RootFolder以下にあるフォルダである必要がある
            fbd.SelectedPath = @"C:\";
            //ユーザーが新しいフォルダを作成できるようにする
            //デフォルトでTrue
            fbd.ShowNewFolderButton = true;

            //ダイアログを表示する
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                FolderPath = fbd.SelectedPath;
            }
        }
        #endregion

        /// <summary>
        /// CSV選択ボタン押下時処理
        /// </summary>
        private void SelectFile()
        {
            //OpenFileDialogクラスのインスタンスを作成
            OpenFileDialog ofd = new OpenFileDialog();

            //はじめのファイル名を指定する
            //はじめに「ファイル名」で表示される文字列を指定する
            ofd.FileName = "pass.csv";
            //はじめに表示されるフォルダを指定する
            //指定しない（空の文字列）の時は、現在のディレクトリが表示される
            ofd.InitialDirectory = @"C:\";
            //[ファイルの種類]に表示される選択肢を指定する
            //指定しないとすべてのファイルが表示される
            ofd.Filter =
                "CSVファイル(*.csv)|*.csv|すべてのファイル(*.*)|*.*";
            //[ファイルの種類]ではじめに
            //「すべてのファイル」が選択されているようにする
            ofd.FilterIndex = 1;
            //タイトルを設定する
            ofd.Title = "CSVファイルを選択してください";
            //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
            ofd.RestoreDirectory = true;
            //存在しないファイルの名前が指定されたとき警告を表示する
            //デフォルトでTrueなので指定する必要はない
            ofd.CheckFileExists = true;
            //存在しないパスが指定されたとき警告を表示する
            //デフォルトでTrueなので指定する必要はない
            ofd.CheckPathExists = true;

            //ダイアログを表示する
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //テキストボックスへパスを格納
                SelectFilePath = ofd.FileName;
            }
        }
        #endregion

        /// <summary>
        /// 実行ボタン押下時処理
        /// </summary>
        private void Execute_Button()
        {
            if (!doing)
            {
                //処理未実行の場合はパスワード設定処理を行う
                if (FolderPath.Length == 0 || SelectFilePath.Length == 0)
                {
                    MessageBox.Show("フォルダーとCSVを選択してください。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    SetStartText();
                    Task.Factory.StartNew(() =>
                    {
                        Execute();
                    });
                }
            }
            else
            {
                //処理実行中の場合はキャンセル処理を行う
                SetCancelText();
            }
        }
        #endregion

        /// <summary>
        /// Excelパスワード設定処理を実装
        /// </summary>
        private void Execute()
        {
            //初期化
            int success = 0;
            int csvCount = 0;

            //各種デリゲートを作成
            ProgressBarDelegate resetBar = new ProgressBarDelegate(ResetProgressBar);
            SetProgressBarDelegate setBar = new SetProgressBarDelegate(SetProgressBar);
            ProgressBarDelegate incBar = new ProgressBarDelegate(IncrementProgress);
            ChangeExecuteDelegate setStart = new ChangeExecuteDelegate(SetStartText);
            ChangeExecuteDelegate setCancel = new ChangeExecuteDelegate(SetCancelText);

            try
            {
                //CSV読み込み処理
                System.Collections.Generic.List<string[]> csvRecords = readCsvFile();
                csvCount = csvRecords.Count;

                //Excelファイル読み込み
                System.Collections.Generic.List<string> excelFiles = getExcelFiles();

                ////プログレスバーの初期化
                DispatcherHelper.UIDispatcher.BeginInvoke(resetBar);
                //this.Invoke(resetBar);

                ////プログレスバーの初期設定
                DispatcherHelper.UIDispatcher.BeginInvoke(setBar, new object[] { 0, csvRecords.Count });
                //this.Invoke(setBar, new object[] { 0, csvRecords.Count });

                Trace.WriteLine("パスワード設定の開始");

                //CSVのファイル名と、発見されたExcel名が一致した場合パスワードを設定して保存する
                Parallel.ForEach(csvRecords, target =>
                {
                    DispatcherHelper.UIDispatcher.BeginInvoke(incBar);
                    //this.Invoke(incBar);
                    if (doing)
                    {
                        Trace.WriteLine(target[0] + "のパスワード設定中");
                        Parallel.ForEach(excelFiles, excelFile =>
                        {
                            if (doing)
                            {
                                String fileName = Path.GetFileName(excelFile);

                                if (fileName.Equals(target[0]))
                                {
                                    using (Excel excel = new Excel(excelFile))
                                    {
                                        try
                                        {
                                            excel.read();
                                            excel.Password = target[1];
                                            excel.saveWithPassword();
                                            success++;
                                        }
                                        catch (COMException)
                                        {
                                            Trace.WriteLine(fileName + "はパスワードが設定されている為、パスワードを設定できません。");
                                            //チェックボックスにチェックが入っている場合はダイアログを表示しない
                                            if (blnError)
                                            {
                                                MessageBox.Show(fileName + "はパスワードが設定されている為、パスワードを設定できません。", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            }
                                        }
                                        catch (Exception err)
                                        {
                                            Trace.WriteLine(err.Message);
                                            MessageBox.Show(err.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                    }
                                }
                            }
                        });
                    }
                });
            }
            catch (Exception err)
            {
                Trace.WriteLine(err.Message);
                MessageBox.Show(err.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                //プログレスバーをリセットし、実行状態を「未実行」に設定する
                //this.Invoke(resetBar);
                DispatcherHelper.UIDispatcher.BeginInvoke(resetBar);
                //this.Invoke(setCancel);
                DispatcherHelper.UIDispatcher.BeginInvoke(setCancel);
                GC.Collect();
            }
            if (success > 0)
            {
                Trace.WriteLine("パスワード設定完了");

                MessageBox.Show("パスワード設定が完了しました。\r\n" +
                        "CSVに定義された" + csvCount +
                        "件のデータの内、" + success +
                        "件のパスワード設定に成功しました", "完了", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        /// <summary>
        /// CSV取得処理
        /// </summary>
        /// <returns>CSVから生成された２次元配列</returns>
        private System.Collections.Generic.List<string[]> readCsvFile()
        {
            //プログレスバーの初期化
            Trace.WriteLine("CSVファイルの取得開始");

            if (!File.Exists(SelectFilePath))
            {
                throw new System.IO.FileNotFoundException("CSVファイルが見つかりません");
            }

            System.Collections.Generic.List<string[]> csvRecords = new System.Collections.Generic.List<string[]>();

            //Shift JISで読み込む
            Microsoft.VisualBasic.FileIO.TextFieldParser tfp =
                new Microsoft.VisualBasic.FileIO.TextFieldParser(SelectFilePath, System.Text.Encoding.GetEncoding(932));

            //フィールドが文字で区切られているとする
            //デフォルトでDelimitedなので、必要なし
            tfp.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited;
            //区切り文字を,とする
            tfp.Delimiters = new string[] { "," };
            //フィールドを"で囲み、改行文字、区切り文字を含めることができるか

            while (!tfp.EndOfData)
            {
                //フィールドを読み込む
                string[] fields = tfp.ReadFields();

                if (fields.Length != 2)
                {
                    throw new System.IO.FileLoadException("CSV読み取りが失敗しました。CSVフォーマットが不正な可能性があります。");
                }

                //保存
                csvRecords.Add(fields);
            }
            //後始末
            tfp.Close();

            if (csvRecords.Count == 0)
            {
                throw new System.IO.FileLoadException("CSVファイルの中身が空です");
            }

            Trace.WriteLine("CSVファイルの取得終了");

            return csvRecords;
        }

        /// <summary>
        /// Excelファイル取得処理
        /// </summary>
        /// <returns>Excelファイルのパスリスト</returns>
        private System.Collections.Generic.List<string> getExcelFiles()
        {
            //抽出用拡張子定義リスト
            System.Collections.Generic.List<string> extentions = new System.Collections.Generic.List<string>();
            extentions.Add("*.xlt");
            extentions.Add("*.xls");
            extentions.Add("*.xml");
            extentions.Add("*.xla");
            extentions.Add("*.xlw");

            Trace.WriteLine("Excelファイルの取得開始");

            if (!Directory.Exists(FolderPath))
            {
                throw new System.IO.FileNotFoundException("Excelフォルダが見つかりません");
            }

            //エクセルファイル格納用リスト定義
            System.Collections.Generic.List<string> files = new System.Collections.Generic.List<string>();
            foreach (var ext in extentions)
            {
                //フォルダ内のエクセルファイルを取得する
                files.AddRange(System.IO.Directory.GetFiles(FolderPath, ext, System.IO.SearchOption.TopDirectoryOnly));
                Trace.WriteLine(ext + "ファイルの取得終了");
            }
            string[] strFile = files.ToArray();

            Trace.WriteLine("Excelファイル取得正常終了");

            return files;
        }

        /// <summary>
        /// プログレスバーを初期化する
        /// </summary>
        private void ResetProgressBar()
        {
            progressValue = 0;
            progressMinimum = 0;
            progressMaximum = 0;
        }

        /// <summary>
        /// プログレスバーの初期設定を行う
        /// </summary>
        private void SetProgressBar(int min, int max)
        {
            progressMinimum = min;
            progressMaximum = max;
        }

        /// <summary>
        /// プログレスの加算を行う
        /// </summary>
        private void IncrementProgress()
        {
            if (progressValue < progressMaximum)
            {
                progressValue += 1;
            }
            //progressUpdate();
        }

        /// <summary>
        /// Executeボタンのテキストを変更する
        /// </summary>
        private void SetStartText()
        {
            doing = true;
            controllLock = false;
            executeText = "キャンセル";
        }

        /// <summary>
        /// Executeボタンのテキストを変更する
        /// </summary>
        private void SetCancelText()
        {
            doing = false;
            controllLock = true;
            executeText = "実行";
        }
    }
}

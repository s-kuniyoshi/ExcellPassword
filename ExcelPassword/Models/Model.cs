
using Livet;
using System;

namespace ExcelPassword.Models
{
    public class Model : NotificationObject
    {

        #region FolderPath変更通知プロパティ
        private string _FolderPath;

        public string folderPath
        {
            get
            { return _FolderPath; }
            set
            {
                if (_FolderPath == value)
                    return;
                _FolderPath = value;
                RaisePropertyChanged("FolderPath");
            }
        }
        #endregion

        #region SelectFilePath変更通知プロパティ
        private string _SelectFilePath;

        public string selectFilePath
        {
            get
            { return _SelectFilePath; }
            set
            {
                if (_SelectFilePath == value)
                    return;
                _SelectFilePath = value;
                RaisePropertyChanged("SelectFilePath");
            }
        }
        #endregion


        #region blnError変更通知プロパティ
        private Boolean _blnError;

        public Boolean blnError
        {
            get
            { return _blnError; }
            set
            {
                if (_blnError == value)
                    return;
                _blnError = value;
                RaisePropertyChanged("blnError");
            }
        }
        #endregion


        #region progressValue変更通知プロパティ
        private int _progressValue;

        public int progressValue
        {
            get
            { return _progressValue; }
            set
            {
                if (_progressValue == value)
                    return;
                _progressValue = value;
                RaisePropertyChanged("progressValue");
            }
        }
        #endregion


        #region progressMinimum変更通知プロパティ
        private int _progressMinimum;

        public int progressMinimum
        {
            get
            { return _progressMinimum; }
            set
            {
                if (_progressMinimum == value)
                    return;
                _progressMinimum = value;
                RaisePropertyChanged("progressMinimum");
            }
        }
        #endregion


        #region progressMaximum変更通知プロパティ
        private int _progressMaximum;

        public int progressMaximum
        {
            get
            { return _progressMaximum; }
            set
            {
                if (_progressMaximum == value)
                    return;
                _progressMaximum = value;
                RaisePropertyChanged("progressMaximum");
            }
        }
        #endregion


        #region controllLock変更通知プロパティ
        private bool _controllLock;

        public bool controllLock
        {
            get
            { return _controllLock; }
            set
            {
                if (_controllLock == value)
                    return;
                _controllLock = value;
                RaisePropertyChanged("controllLock");
            }
        }
        #endregion


        #region doing変更通知プロパティ
        private bool _doing;

        public bool doing
        {
            get
            { return _doing; }
            set
            {
                if (_doing == value)
                    return;
                _doing = value;
                RaisePropertyChanged("doing");
            }
        }
        #endregion

        #region executeText変更通知プロパティ
        private string _executeText;

        public string executeText
        {
            get
            { return _executeText; }
            set
            {
                if (_executeText == value)
                    return;
                _executeText = value;
                RaisePropertyChanged("executeText");
            }
        }
        #endregion


    }
}

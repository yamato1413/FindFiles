using System;

namespace FindFiles
{
    class SearchCondition
    {
        public enum Wildcards
        {
            Contain,
            Match,
            Start,
            End,
        }

        public const string DEFAULT_BASEDIRECTORY = @"C:\";
        public const string DEFAULT_KEYWORD = "";
        public const int DEFAULT_DEPTH = 1;
        public const bool DEFAULT_SEARCHFOLDER = true;
        public const bool DEFAULT_SERACHFILE = true;
        public const bool DEFAULT_SORT = true;
        public const Wildcards DEFAULT_WILDCARD = Wildcards.Contain;

        private string _BaseDirectory = DEFAULT_BASEDIRECTORY;
        public string BaseDirectory
        {
            get { return _BaseDirectory == "" ? DEFAULT_BASEDIRECTORY : _BaseDirectory; }
            set { _BaseDirectory = value; }
        }

        private string _Keyword = DEFAULT_KEYWORD;
        public string Keyword
        {
            get { return _Keyword; }
            set { _Keyword = value; }
        }

        private int _Depth = DEFAULT_DEPTH;
        public int Depth
        {
            get { return _Depth; }
            set
            {
                if (value < 1) throw new Exception("1未満の値は代入できません");
                _Depth = value;
            }
        }

        private bool _SearchFolder = DEFAULT_SEARCHFOLDER;
        public bool SearchFolder
        {
            get { return _SearchFolder; }
            set { _SearchFolder = value; }
        }

        private bool _SearchFile = DEFAULT_SERACHFILE;
        public bool SearchFile
        {
            get { return _SearchFile; }
            set { _SearchFile = value; }
        }

        private bool _Sort = DEFAULT_SORT;
        public bool Sort
        {
            get { return _Sort; }
            set { _Sort = value; }
        }

        private Wildcards _Wildcard = DEFAULT_WILDCARD;
        public Wildcards Wildcard
        {
            get { return _Wildcard; }
            set { _Wildcard = value; }
        }
    }

}
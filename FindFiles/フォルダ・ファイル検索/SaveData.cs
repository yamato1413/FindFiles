using System;
using System.Linq;
using Microsoft.Win32;

namespace FindFiles
{
    class SaveData
    {
        private static RegistryKey key = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Yamato\FindFiles");
        public static WindowPos WindowPosition
        {
            get
            {
                int[] pos = key.GetValue("Position", WindowPos.DEFAULT_POSITION).ToString().Split(' ').Select(x => int.Parse(x)).ToArray();
                return new WindowPos(pos[0], pos[1], pos[2], pos[3]);
            }
            set
            {
                string pos = String.Join(" ", new int[] { value.Top, value.Left, value.Width, value.Height }.Select(x => x.ToString()));
                key.SetValue("Position", pos);
            }
        }
        public static SearchCondition SearchCondition
        {
            get
            {
                SearchCondition sc = new SearchCondition();
                try
                {
                    sc.BaseDirectory = key.GetValue("BaseDirectory", SearchCondition.DEFAULT_BASEDIRECTORY).ToString();
                    sc.Keyword = key.GetValue("Keyword", SearchCondition.DEFAULT_KEYWORD).ToString();
                    sc.Depth = (int)(key.GetValue("Depth", SearchCondition.DEFAULT_DEPTH));
                    sc.SearchFolder = Convert.ToBoolean(key.GetValue("SearchFolder", SearchCondition.DEFAULT_SEARCHFOLDER).ToString());
                    sc.SearchFile = Convert.ToBoolean(key.GetValue("SearchFile", SearchCondition.DEFAULT_SERACHFILE).ToString());
                    sc.Sort = Convert.ToBoolean(key.GetValue("Sort", SearchCondition.DEFAULT_SORT).ToString());
                    sc.Wildcard = (SearchCondition.Wildcards)(int)(key.GetValue("Wildcard", SearchCondition.DEFAULT_WILDCARD));
                }
                catch { }
                return sc;
            }
            set
            {
                key.SetValue("BaseDirectory", value.BaseDirectory, RegistryValueKind.String);
                key.SetValue("Keyword", value.Keyword, RegistryValueKind.String);
                key.SetValue("Depth", value.Depth, RegistryValueKind.DWord);
                key.SetValue("SearchFolder", value.SearchFolder.ToString(), RegistryValueKind.String);
                key.SetValue("SearchFile", value.SearchFile.ToString(), RegistryValueKind.String);
                key.SetValue("Sort", value.Sort.ToString(), RegistryValueKind.String);
                key.SetValue("Wildcard", (int)value.Wildcard, RegistryValueKind.DWord);
            }
        }
    }

}
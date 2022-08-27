using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace FindFiles
{
    class IndexMaker
    {
        private string appDir;
        private List<string[]> allItems;

        public IndexMaker()
        {
            appDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            allItems = new List<string[]>();
            MakeIndex();
        }
        private void MakeIndex()
        {
            SearchCondition sc = SaveData.SearchCondition;
            CollectItems(sc.BaseDirectory, 5, sc);

            FileStream fs = new FileStream(
                sc.BaseDirectory + @"\findfiles_index.txt",
                FileMode.Create,
                FileAccess.Write,
                FileShare.None);
            StreamWriter sw = new StreamWriter(fs);
            foreach (string[] item in allItems) sw.WriteLine(string.Join("|", item));
            sw.Close();
        }
        private void CollectItems(string searchDirectory, int depth, SearchCondition sc)
        {
            IEnumerable<string> subfolders = new List<string>();
            IEnumerable<string> items = new List<string>();

            // アクセス権限のないフォルダだとエラーが発生するのでエラーをもみ消す。
            try
            {
                // サブフォルダを取得
                subfolders = Directory.EnumerateDirectories(searchDirectory);
                items = Directory.EnumerateFiles(searchDirectory);
            }
            catch { }
            lock (allItems)
            {
                foreach (string item in subfolders)
                {
                    allItems.Add(new string[] {
                    item,
                    item.Replace(sc.BaseDirectory, "").Count(c => c == '\\').ToString(),
                    "folder" }
                    );

                }
                foreach (string item in items)
                {
                    allItems.Add(new string[] {
                    item,
                    item.Replace(sc.BaseDirectory, "").Count(c => c == '\\').ToString(),
                    "file" }
                    );
                }
            }
            if (depth > 1)
            {
                Parallel.ForEach(subfolders, nextSearchDirectory => CollectItems(nextSearchDirectory, depth - 1, sc));
            }
        }
    }
}

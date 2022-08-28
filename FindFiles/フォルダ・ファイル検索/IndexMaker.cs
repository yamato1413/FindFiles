using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace FindFiles
{
    class IndexMaker
    {
        public string IndexPath;
        private List<string[]> allItems;
        private SearchCondition sc;
        public IndexMaker()
        {
            this.sc = SaveData.SearchCondition;
            // IndexPath = sc.BaseDirectory + @"\findfiles_index.txt";
            IndexPath = Directory.GetParent(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\findfiles_index.txt";
            allItems = new List<string[]>();
        }
        public void MakeIndex()
        {
            CollectItems(sc.BaseDirectory, 5, sc);
            FileStream fs = new FileStream(
                IndexPath,
                FileMode.OpenOrCreate,
                FileAccess.Write,
                FileShare.None);
            StreamWriter sw = new StreamWriter(fs);
            foreach (string[] item in allItems)
            {
                try { sw.WriteLine(string.Join("|", item)); } catch { }
            }
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
                    string[] buf;
                    try
                    {
                        buf = new string[] {
                            item,
                            item.Replace(sc.BaseDirectory, "").Count(c => c == '\\').ToString(),
                            "folder"
                        };
                    }
                    catch
                    {
                        buf = new string[] { "error", "99999", "folder" };
                    }
                    allItems.Add(buf);

                }
                foreach (string item in items)
                {
                    string[] buf;
                    try
                    {
                        buf = new string[] {
                            item,
                            item.Replace(sc.BaseDirectory, "").Count(c => c == '\\').ToString(),
                            "file"
                        };
                    }
                    catch
                    {
                        buf = new string[] { "error", "99999", "file" };
                    }
                    allItems.Add(buf);
                }
            }
            if (depth > 1)
            {
                Parallel.ForEach(subfolders, nextSearchDirectory => CollectItems(nextSearchDirectory, depth - 1, sc));
            }
        }

        public DirectoryItem ParseIndex(StreamReader index)
        {
            List<DirectoryItem> items = new List<DirectoryItem>();
            while (!index.EndOfStream)
            {
                string line = index.ReadLine();
                string[] buf = line.Split('|');
                string[] splitpath = buf[0].Replace(@"\\", "??").Split('\\');
                splitpath[0] = splitpath[0].Replace("??", @"\\");
                DirectoryItem.ItemAttribute attribute =
                    buf[2] == "folder"
                    ? DirectoryItem.ItemAttribute.Folder
                    : DirectoryItem.ItemAttribute.File;
                items.Add(Parse(splitpath, attribute));
            }
            DirectoryItem root = items.ElementAt(0);
            MakeTree(items);
            return root;
        }
        private DirectoryItem Parse(string[] splitpath, DirectoryItem.ItemAttribute attribute)
        {
            DirectoryItem root, parent, child;
            string fullpath, name;

            // ドライブ
            name = splitpath[0];
            fullpath = splitpath[0];
            root = new DirectoryItem(DirectoryItem.ItemAttribute.Drive, name, fullpath);
            parent = root;

            // 途中のフォルダ
            for (int i = 1; i < splitpath.Length - 1; i++)
            {
                name = splitpath[i];
                fullpath = Path.Combine(parent.FullName, name);
                child = new DirectoryItem(DirectoryItem.ItemAttribute.Folder, name, fullpath);
                parent.Children.Add(child);
                parent = child;
            }

            // 末尾のフォルダorファイル
            name = splitpath[splitpath.Length - 1];
            fullpath = Path.Combine(parent.FullName, name);
            parent.Children.Add(new DirectoryItem(attribute, name, fullpath));

            return root;
        }
        private void MakeTree(List<DirectoryItem> items)
        {
            DirectoryItem root = items.ElementAt(0);
            items.RemoveAt(0);
            foreach (DirectoryItem item in items)
            {
                if (item.Name == root.Name)
                {
                    root.Children.AddRange(item.Children);
                }
                else
                {
                    root.Children.Add(item);
                }
            }
            foreach (DirectoryItem child in root.Children)
            {
                if (child.HasChild())
                {
                    MakeTree(child.Children);
                }
            }
        }
    }

    class DirectoryItem
    {
        public enum ItemAttribute
        {
            Drive,
            Folder,
            File,
        }
        public ItemAttribute Attribute;
        private string _Name;
        public string Name
        {
            get { return _Name; }
            set
            {
                _Name = value;
                FullName = Path.Combine(Path.GetDirectoryName(_FullName), _Name);
            }
        }

        private string _FullName;
        public string FullName
        {
            get { return _FullName; }
            set
            {
                _FullName = value;
                if (HasChild())
                {
                    foreach (DirectoryItem child in Children)
                    {
                        child.FullName = Path.Combine(this._FullName, child.Name);
                    }
                }
            }
        }

        public List<DirectoryItem> Children;
        public DirectoryItem(ItemAttribute attribute, string name, string fullname)
        {
            this.Attribute = attribute;
            this.Name = name;
            this.FullName = fullname;
            this.Children = new List<DirectoryItem>();
        }
        public bool HasChild()
        {
            return Children != null;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Microsoft.Win32;

namespace FindFiles
{
    public class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            if (args.Length != 0 && args[0] == "--makeIndex")
            {
                new IndexMaker();
            }
            else
            {
                Application.Run(new MainForm());
            }
        }
    }

    class MainForm : Form
    {

        private System.Windows.Forms.Label lblBaseDirectory = new System.Windows.Forms.Label();
        private System.Windows.Forms.Label lblKeyword = new System.Windows.Forms.Label();
        private System.Windows.Forms.Label lblDepth = new System.Windows.Forms.Label();
        private System.Windows.Forms.Label lblFinish = new System.Windows.Forms.Label();

        private TextBox txtBaseDirectory = new TextBox();
        private TextBox txtKeyword = new TextBox();
        private TextBox txtDepth = new TextBox();

        private Button btnDirectory = new Button();
        private Button btnSearch = new Button();

        private CheckBox chkFolder = new CheckBox();
        private CheckBox chkFile = new CheckBox();
        private CheckBox chkSort = new CheckBox();

        private ComboBox cmbWildcard = new ComboBox();
        private DataGridView lvResult = new DataGridView();
        private ProgressBar progress = new ProgressBar();

        private const int heightControls = 25;
        private const int leftControls = 20;
        private const int topBaseDirectory = 30;
        private const int topSearchCond = 90;
        private const int topResult = 150;

        private bool cancel = false;

        public MainForm()
        {
            init_FormWindow();

            init_lblBaseDirectory();
            init_txtBaseDirectory();
            init_btnDirectory();
            init_lblKeyword();
            init_chkFolder();
            init_chkFile();
            init_chkSort();
            init_txtKeyword();
            init_cmbWildcard();
            init_lblDepth();
            init_txtDepth();
            init_btnSearch();
            init_lvResult();
            init_lblFinish();
            init_progress();

            load_SaveData();

        }

        private void init_FormWindow()
        {
            WindowPos pos = SaveData.WindowPosition;
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new Point(pos.Left, pos.Top);
            this.Size = new Size(pos.Width, pos.Height);
            this.MinimumSize = new Size(400, 250);

            // フォームを閉じるときに位置とサイズを記録する
            this.Closing += ((object sender, System.ComponentModel.CancelEventArgs e) =>
                SaveData.WindowPosition = new WindowPos(this.Top, this.Left, this.Width, this.Height));

            this.Text = "フォルダ・ファイル検索 - Yamato_Tools";
            this.AllowDrop = true;
            this.Shown += ((object sender, EventArgs e) => txtKeyword.Focus());
        }

        private void init_lblBaseDirectory()
        {
            lblBaseDirectory.Location = new Point(leftControls, topBaseDirectory - 20);
            lblBaseDirectory.AutoSize = true;
            lblBaseDirectory.Text = "探索開始フォルダ(フォルダをドラッグ&ドロップで指定可)";
            lblBaseDirectory.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
            Controls.Add(lblBaseDirectory);
        }

        private void init_txtBaseDirectory()
        {
            txtBaseDirectory.AutoSize = false;
            txtBaseDirectory.Location = new Point(leftControls, topBaseDirectory);
            txtBaseDirectory.Size = new Size(Width - 100, heightControls);
            txtBaseDirectory.BorderStyle = BorderStyle.FixedSingle;
            txtBaseDirectory.AllowDrop = true;
            txtBaseDirectory.DragDrop += txtBaseDirectory_DragDrop;
            txtBaseDirectory.DragEnter += ((object sender, DragEventArgs e) => e.Effect = DragDropEffects.All);
            txtBaseDirectory.KeyDown += SelectAllText;
            txtBaseDirectory.Anchor = (AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right);
            txtBaseDirectory.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
            Controls.Add(txtBaseDirectory);
        }

        private void init_btnDirectory()
        {
            btnDirectory.Location = new Point(txtBaseDirectory.Right + 10, topBaseDirectory);
            btnDirectory.Size = new Size(30, heightControls);
            btnDirectory.Text = "..";
            btnDirectory.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
            btnDirectory.Click += SelectDirectory;
            btnDirectory.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
            Controls.Add(btnDirectory);
        }

        private void init_lblKeyword()
        {
            lblKeyword.Location = new Point(leftControls, topSearchCond - 20);
            lblKeyword.AutoSize = true;
            lblKeyword.Text = "検索条件";
            lblKeyword.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
            Controls.Add(lblKeyword);
        }

        private void init_chkFolder()
        {
            chkFolder.Location = new Point(lblKeyword.Right, topSearchCond - 20);
            chkFolder.AutoSize = true;
            chkFolder.Text = "フォルダ";
            chkFolder.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
            Controls.Add(chkFolder);
        }

        private void init_chkFile()
        {
            chkFile.Location = new Point(chkFolder.Right, topSearchCond - 20);
            chkFile.AutoSize = true;
            chkFile.Text = "ファイル";
            chkFile.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
            Controls.Add(chkFile);
        }

        private void init_chkSort()
        {
            chkSort.Location = new Point(chkFile.Right, topSearchCond - 20);
            chkSort.AutoSize = true;
            chkSort.Text = "検索終了後にソート";
            chkSort.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
            Controls.Add(chkSort);
        }

        private void init_txtKeyword()
        {
            txtKeyword.AutoSize = false;
            txtKeyword.Location = new Point(leftControls, topSearchCond);
            txtKeyword.Size = new Size(Width - 350, heightControls);
            txtKeyword.BorderStyle = BorderStyle.FixedSingle;
            txtKeyword.Anchor = (AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right);
            txtKeyword.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
            txtKeyword.KeyDown += SelectAllText;
            Controls.Add(txtKeyword);
        }

        private void init_cmbWildcard()
        {
            cmbWildcard.Location = new Point(txtKeyword.Right + 5, topSearchCond);
            cmbWildcard.Size = new Size(120, heightControls);
            cmbWildcard.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
            cmbWildcard.Items.AddRange(new string[] { "を含む", "と一致する", "から始まる", "で終わる" });
            cmbWildcard.IntegralHeight = true;
            cmbWildcard.FlatStyle = FlatStyle.Flat;
            cmbWildcard.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbWildcard.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
            Controls.Add(cmbWildcard);
        }

        private void init_lblDepth()
        {
            lblDepth.Location = new Point(cmbWildcard.Right + 5, topSearchCond - 20);
            lblDepth.AutoSize = true;
            lblDepth.Text = "探索する深さ";
            lblDepth.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
            lblDepth.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
            Controls.Add(lblDepth);
        }

        private void init_txtDepth()
        {
            txtDepth.AutoSize = false;
            txtDepth.Location = new Point(cmbWildcard.Right + 5, topSearchCond);
            txtDepth.Size = new Size(50, heightControls);
            txtDepth.BorderStyle = BorderStyle.FixedSingle;
            txtDepth.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
            txtDepth.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
            txtDepth.KeyDown += SelectAllText;
            txtDepth.TextChanged += ((object sender, EventArgs e) =>
                txtDepth.Text = Strings.StrConv(txtDepth.Text, VbStrConv.Narrow));
            Controls.Add(txtDepth);
        }

        private void init_btnSearch()
        {
            btnSearch.Location = new Point(txtDepth.Right + 10, topSearchCond - 1);
            btnSearch.Size = new Size(100, heightControls);
            btnSearch.Text = "検索";
            btnSearch.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
            btnSearch.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
            btnSearch.Click += btnSearch_Click;
            Controls.Add(btnSearch);
        }

        private void init_lvResult()
        {
            lvResult.Location = new Point(leftControls, topResult);
            lvResult.Size = new Size(Width - 60, Height - 200);
            lvResult.BorderStyle = BorderStyle.FixedSingle;
            lvResult.Anchor = (AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right);
            lvResult.ReadOnly = true;   // セルをクリックしたときに編集モードに入らないようにする
            lvResult.MultiSelect = false;
            lvResult.ColumnCount = 1;
            lvResult.Columns[0].Name = "結果一覧";
            lvResult.Columns[0].MinimumWidth = 2400;
            lvResult.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            lvResult.CellContentDoubleClick += ((Object sender, DataGridViewCellEventArgs e) =>
            {
                try { Process.Start("Explorer", "/select," + lvResult.Rows[e.RowIndex].Cells[0].Value.ToString()); }
                catch { }
            });
            lvResult.CellPainting += lvResult_CellPainting;
            lvResult.ColumnHeaderMouseClick += ((object sender, DataGridViewCellMouseEventArgs e) => SortItem());

            lvResult.AllowUserToAddRows = false;
            lvResult.AllowUserToDeleteRows = false;
            lvResult.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;
            Controls.Add(lvResult);
        }

        private void init_progress()
        {
            progress.Location = new Point(leftControls, 130);
            progress.Size = new Size(150, 20);
            progress.Style = ProgressBarStyle.Continuous;
            Controls.Add(progress);
        }

        private void init_lblFinish()
        {
            lblFinish.Location = new Point(leftControls + 150, 130);
            lblFinish.AutoSize = true;
            lblFinish.Anchor = (AnchorStyles.Top | AnchorStyles.Left);
            lblFinish.Font = new Font(new FontFamily("BIZ UDゴシック"), 9);
            Controls.Add(lblFinish);
        }

        private void load_SaveData()
        {
            SearchCondition sc = SaveData.SearchCondition;
            txtBaseDirectory.Text = sc.BaseDirectory;
            txtKeyword.Text = sc.Keyword;
            txtDepth.Text = sc.Depth.ToString();
            chkFolder.Checked = sc.SearchFolder;
            chkFile.Checked = sc.SearchFile;
            chkSort.Checked = sc.Sort;
            cmbWildcard.SelectedIndex = (int)sc.Wildcard;
        }

        /// <summary>
        /// フォルダ選択ダイアログを表示する。
        /// </summary>
        private void SelectDirectory(object sender, EventArgs e)
        {
            FolderSelectDialog dialog = new FolderSelectDialog();
            if (dialog.ShowDialog() == DialogResult.OK) txtBaseDirectory.Text = dialog.Path;
        }

        /// <summary>
        /// 検索を実行する。
        /// </summary>
        private async void btnSearch_Click(object sender, EventArgs e)
        {
            // 動作中なら停止フラグを立ててリターン
            if (btnSearch.Text == "停止")
            {
                cancel = true;
                return;
            }

            // ガード
            if (!CanParseDepth())
            {
                MessageBox.Show("探索する深さには0以上の整数を入力してください。");
                return;
            }

            InitializeBeforeSearch();
            SearchCondition sc = Save();

            string indexpath = sc.BaseDirectory + @"\findfiles_index.txt";
            StreamReader index = OpenIndexFile(indexpath);
            if (index != null && IsOldIndex(indexpath))
            {
                List<string[]> items = new List<string[]>();
                while (!index.EndOfStream)
                {
                    items.Add(index.ReadLine().Split('|'));
                }
                await Task.Run(() => FindFilesFromIndex(items, sc));
            }
            else
            {
                // インデックスが存在しなかったり作成中で開けなかった場合
                MakeIndex();
                await Task.Run(() => FindFiles(sc.BaseDirectory, sc.Depth, sc));
            }

            if (index != null) index.Close();

            if (chkSort.Checked) SortItem();

            // Finalize
            SetProgress(0);
            lblFinish.Text = "検索完了";
            btnSearch.Text = SwitchString(btnSearch.Text, "検索", "停止");
        }

        private bool IsOldIndex(string indexpath)
        {
            return File.GetLastWriteTime(indexpath).AddMinutes(15) < DateTime.Now;
        }

        private bool CanParseDepth()
        {
            int depth;
            return int.TryParse(txtDepth.Text, out depth);
        }

        private void InitializeBeforeSearch()
        {
            lblFinish.Text = "";
            lvResult.Rows.Clear();
            cancel = false;
            btnSearch.Text = SwitchString(btnSearch.Text, "検索", "停止");
        }

        private SearchCondition Save()
        {
            // 検索条件を保存
            SearchCondition sc = new SearchCondition();
            sc.BaseDirectory = txtBaseDirectory.Text;
            sc.Keyword = txtKeyword.Text;
            sc.Depth = int.Parse(txtDepth.Text);
            sc.SearchFolder = chkFolder.Checked;
            sc.SearchFile = chkFile.Checked;
            sc.Sort = chkSort.Checked;
            sc.Wildcard = (SearchCondition.Wildcards)cmbWildcard.SelectedIndex;
            SaveData.SearchCondition = sc;
            return sc;
        }

        private StreamReader OpenIndexFile(string pathIndexFile)
        {
            try
            {
                FileStream fs = new FileStream(pathIndexFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                return new StreamReader(fs);
            }
            catch
            {
                return null;
            }
        }

        private void MakeIndex()
        {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = System.Reflection.Assembly.GetExecutingAssembly().Location;
            psi.Arguments = "--makeIndex";
            Process.Start(psi);
        }


        /// <summary>
        /// 入力された文字列と異なる方の文字列を返却する。
        /// str1,str2のどちらでもない場合は""を返却する。
        /// </summary>
        /// <param name="exp">入力</param>
        /// <param name="str1">返却したい文字列1</param>
        /// <param name="str2">返却したい文字列2</param>
        /// <returns>str1 => str2, str2 => str1, _ => ""</returns>
        private string SwitchString(string exp, string str1, string str2)
        {
            if (exp == str1) return str2;
            else if (exp == str2) return str1;
            else return "";
        }

        /// <summary>
        /// ファイル検索処理の実装
        /// </summary>
        ///  <param name="searchDirectory">検索したいディレクトリ</param>
        /// <param name="depth">ディレクトリを何段潜って検索するか</param>
        /// <param name="condition">検索条件</param>
        private void FindFiles(string searchDirectory, int depth, SearchCondition condition)
        {
            if (cancel) return;
            if (!condition.SearchFolder && !condition.SearchFile) return;

            IEnumerable<string> subfolders = new List<string>();
            IEnumerable<string> files = new List<string>();


            // アクセス権限のないフォルダだとエラーが発生するのでエラーをもみ消す。
            try
            {
                // サブフォルダを取得
                SetProgress(10);
                subfolders = Directory.EnumerateDirectories(searchDirectory);
                files = Directory.EnumerateFiles(searchDirectory);

                // 条件に合うフォルダを検索結果リストに積む
                SetProgress(33);
                if (condition.SearchFolder)
                {
                    Task.Run(() => AddRows(subfolders.Where(f => IsMatch(Path.GetFileName(f), condition))));
                }
                // 条件に合うファイルを検索結果リストに積む
                SetProgress(66);
                if (condition.SearchFile)
                {
                    Task.Run(() => AddRows(files.Where(f => IsMatch(Path.GetFileName(f), condition))));
                }
            }
            catch { }

            SetProgress(100);

            if (depth > 1)
            {
                Parallel.ForEach(subfolders, nextSearchDirectory => FindFiles(nextSearchDirectory, depth - 1, condition));
            }
        }

        private void FindFilesFromIndex(List<string[]> items, SearchCondition sc)
        {
            AddRows(items
                .Where(item =>
                    IsMatch(Path.GetFileName(item[0]), sc)
                    && int.Parse(item[1]) <= sc.Depth
                    && ((item[2] == "folder" && sc.SearchFolder) || (item[2] == "file" && sc.SearchFile)))
                .Select(item => item[0])
            );
        }

        // なぜかRegExがうまく動かないので作成
        private bool IsMatch(string input, SearchCondition condition)
        {
            string inputLower = input.ToLower();
            string keyword = condition.Keyword.ToLower();

            switch (condition.Wildcard)
            {
                case SearchCondition.Wildcards.Contain:
                    return inputLower.Contains(keyword);

                case SearchCondition.Wildcards.Start:
                    return inputLower.StartsWith(keyword);

                case SearchCondition.Wildcards.End:
                    return inputLower.EndsWith(keyword);

                case SearchCondition.Wildcards.Match:
                    return inputLower == keyword;
            }
            throw new Exception("想定外の分岐:IsMatch");
        }

        /// <summary>
        /// 検索結果をUIに反映する。
        /// </summary>
        /// <param name="items">検索結果のリスト</param>
        private void AddRows(IEnumerable<string> items)
        {
            if (lvResult.InvokeRequired)
            {
                lvResult.Invoke(new Action<IEnumerable<string>>(AddRows), new object[] { items });
            }
            else
            {
                foreach (string item in items)
                {
                    lvResult.Rows.Add(new string[] { item });
                }
            }
        }

        /// <summary>
        /// 進捗状況をプログレスバーに反映する。
        /// </summary>
        /// <param name="value">プログレスバーに設定したい値</param>
        private void SetProgress(int value)
        {
            if (progress.InvokeRequired)
            {
                progress.Invoke(new Action<int>(SetProgress), new object[] { value });
            }
            else
            {
                if (value == 0 || value == 100)
                {
                    progress.Value = value;
                }
                else
                {
                    value = new Random().Next(value - 5, value + 5);
                    value = value < 0 ? 0 : value;
                    value = value > 100 ? 100 : value;
                    progress.Value = value;
                }
            }
        }

        /// <summary>
        /// 検索結果一覧をソートする。
        /// </summary>
        private void SortItem()
        {
            if (lvResult.Rows.Count == 0) return;
            lvResult.Sort(lvResult.Columns[0], System.ComponentModel.ListSortDirection.Ascending);
            lvResult.CurrentCell = lvResult[0, 0];
        }

        /// <summary>
        /// ファルダをドラッグアンドドロップしたときの処理。
        /// </summary>
        private void txtBaseDirectory_DragDrop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

            string[] dragFilePathArr = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            txtBaseDirectory.Text = OriginalPath(dragFilePathArr[0]);
        }

        /// <summary>
        /// ショートカットから参照パスを取得する。
        /// ショートカットの先がショートカットだった場合再帰する。
        /// </summary>
        /// <param name="shortcutPath">ショートカットのパス</param>
        /// <returns>ショートカットがさしている先のパス</returns>
        private string OriginalPath(string shortcutPath)
        {
            if (!shortcutPath.EndsWith(".lnk")) return shortcutPath;

            // WshShellを作成
            Type t = Type.GetTypeFromCLSID(new Guid("72C24DD5-D70A-438B-8A42-98424B88AFB8"));
            dynamic shell = Activator.CreateInstance(t);

            var shortcut = shell.CreateShortcut(shortcutPath);
            string targetPath = shortcut.targetPath;
            // ショートカットの先がショートカットだった場合再帰呼び出しする。
            return targetPath.EndsWith(".lnk") ? OriginalPath(targetPath) : targetPath;
        }

        /// <summary>
        /// テキストボックス内でCtrl+Aしたときの処理。
        /// </summary>
        private void SelectAllText(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.A)
            {
                TextBox txtbox = (TextBox)sender;
                txtbox.SelectAll();
            }
        }

        /// <summary>
        /// 結果一覧に連番を振る処理。ネットで拾ってきた。
        /// </summary>
        private void lvResult_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            //列ヘッダーかどうか調べる
            if (e.ColumnIndex < 0 && e.RowIndex >= 0)
            {
                //セルを描画する
                e.Paint(e.ClipBounds, DataGridViewPaintParts.All);

                //行番号を描画する範囲を決定する
                //e.AdvancedBorderStyleやe.CellStyle.Paddingは無視しています
                Rectangle indexRect = e.CellBounds;
                indexRect.Inflate(-2, -2);
                //行番号を描画する
                TextRenderer.DrawText(e.Graphics,
                    (e.RowIndex + 1).ToString(),
                    e.CellStyle.Font,
                    indexRect,
                    e.CellStyle.ForeColor,
                    TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
                //描画が完了したことを知らせる
                e.Handled = true;
            }
        }
    }

}
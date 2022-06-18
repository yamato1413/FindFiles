using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Configuration;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Runtime.Remoting.Lifetime;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Microsoft.Win32;

public class Program
{
    [STAThread]
    public static void Main()
    {
        Application.Run(new MainForm());
    }
}

class MainForm : Form
{
    private RegistryKey reg = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Yamato\FindFiles");

    [Flags]
    private enum SearchCond
    {
        None = 0,       // 0b_00
        Folders = 1,    // 0b_01
        Files = 2       // 0b_10
    }

    private enum Wildcard
    {
        Contain,
        Match,
        Start,
        End
    }

    private System.Windows.Forms.Label lblBaseDirectory = new System.Windows.Forms.Label();
    private System.Windows.Forms.Label lblPattern = new System.Windows.Forms.Label();
    private System.Windows.Forms.Label lblDepth = new System.Windows.Forms.Label();
    private System.Windows.Forms.Label lblFinish = new System.Windows.Forms.Label();

    private TextBox txtBaseDirectory = new TextBox();
    private TextBox txtPattern = new TextBox();
    private TextBox txtDepth = new TextBox();

    private Button btnDirectory = new Button();
    private Button btnSearch = new Button();

    private CheckBox chkFolder = new CheckBox();
    private CheckBox chkFile = new CheckBox();
    private CheckBox chkSort = new CheckBox();

    private ComboBox cmbCond = new ComboBox();
    private DataGridView lvResult = new DataGridView();
    private ProgressBar progress = new ProgressBar();

    private const int heightControls = 25;
    private const int leftControls = 20;
    private const int topBaseDirectory = 30;
    private const int topSearchCond = 90;
    private const int topResult = 150;

    private bool cancel = false;
    private string baseDir;

    public MainForm()
    {
        initFormWindow();
        initBaseDirectory();
        InitSearchCond();
        InitResult();
    }

    private void initFormWindow()
    {
        WindowPos pos = SaveData.WindowPosition;
        this.StartPosition = FormStartPosition.Manual;
        this.Location = new Pointer(pos.Left, pos.Top);
        this.Size = new Size(pos.Width, pos.Height);
        this.MinimumSize = new Size(400, 250);

        // フォームを閉じるときに位置とサイズを記録する
        this.Closing += ((object sender, System.ComponentModel.CancelEventArgs e) =>
        {
            SaveData.WindowPosition = new WindowPos(this.Top, this.Left, this.Width, this.Height);
        });

        this.Text = "フォルダ・ファイル検索 - ISHII_Tools";
        this.AllowDrop = true;
        this.Activated += ((object sender, EventArgs e) => txtPattern.Focus());
    }

    private void initBaseDirectory()
    {
        lblBaseDirectory.Location = new Point(leftControls, topBaseDirectory - 20);
        lblBaseDirectory.AutoSize = true;
        lblBaseDirectory.Text = "探索開始フォルダ(フォルダをドラッグ&ドロップで指定可)";
        lblBaseDirectory.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
        Controls.Add(lblBaseDirectory);

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
        txtBaseDirectory.Text = reg.GetValue("BaseDirectory", @"C:\").ToString();
        Controls.Add(txtBaseDirectory);

        btnDirectory.Location = new Point(txtBaseDirectory.Right + 10, topBaseDirectory);
        btnDirectory.Size = new Size(30, heightControls);
        btnDirectory.Text = "..";
        btnDirectory.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
        btnDirectory.Click += SelectDirectory;
        btnDirectory.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
        Controls.Add(btnDirectory);
    }

    private void InitSearchCond()
    {
        lblPattern.Location = new Point(leftControls, topSearchCond - 20);
        lblPattern.AutoSize = true;
        lblPattern.Text = "検索条件";
        lblPattern.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
        Controls.Add(lblPattern);

        chkFolder.Location = new Point(lblPattern.Right, topSearchCond - 20);
        chkFolder.AutoSize = true;
        chkFolder.Text = "フォルダ";
        chkFolder.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
        chkFolder.Checked = Convert.ToBoolean(reg.GetValue("ShouldSearchFolder", 1));
        Controls.Add(chkFolder);

        chkFile.Location = new Point(chkFolder.Right, topSearchCond - 20);
        chkFile.AutoSize = true;
        chkFile.Text = "ファイル";
        chkFile.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
        chkFile.Checked = Convert.ToBoolean(reg.GetValue("ShouldSearchFile", 1));
        Controls.Add(chkFile);

        chkSort.Location = new Point(chkFile.Right, topSearchCond - 20);
        chkSort.AutoSize = true;
        chkSort.Text = "検索終了後にソート";
        chkSort.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
        chkSort.Checked = Convert.ToBoolean(reg.GetValue("ShouldSort", 1));
        Controls.Add(chkSort);

        txtPattern.AutoSize = false;
        txtPattern.Location = new Point(leftControls, topSearchCond);
        txtPattern.Size = new Size(Width - 350, heightControls);
        txtPattern.BorderStyle = BorderStyle.FixedSingle;
        txtPattern.Anchor = (AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right);
        txtPattern.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
        txtPattern.Text = reg.GetValue("Pattern", "").ToString();
        txtPattern.KeyDown += SelectAllText;
        Controls.Add(txtPattern);

        cmbCond.Location = new Point(txtPattern.Right + 5, topSearchCond);
        cmbCond.Size = new Size(120, heightControls);
        cmbCond.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
        cmbCond.Items.AddRange(new string[] { "を含む", "と一致する", "から始まる", "で終わる" });
        cmbCond.SelectedIndex = Convert.ToInt32(reg.GetValue("Wildcard", 0));
        cmbCond.IntegralHeight = true;
        cmbCond.FlatStyle = FlatStyle.Flat;
        cmbCond.DropDownStyle = ComboBoxStyle.DropDownList;
        cmbCond.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
        Controls.Add(cmbCond);

        lblDepth.Location = new Point(cmbCond.Right + 5, topSearchCond - 20);
        lblDepth.AutoSize = true;
        lblDepth.Text = "探索する深さ";
        lblDepth.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
        lblDepth.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
        Controls.Add(lblDepth);

        txtDepth.AutoSize = false;
        txtDepth.Location = new Point(cmbCond.Right + 5, topSearchCond);
        txtDepth.Size = new Size(50, heightControls);
        txtDepth.BorderStyle = BorderStyle.FixedSingle;
        txtDepth.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
        txtDepth.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
        txtDepth.Text = reg.GetValue("Depth", 1).ToString();
        txtDepth.KeyDown += SelectAllText;
        txtDepth.TextChanged += ((object sender, EventArgs e) => txtDepth.Text = Strings.StrConv(txtDepth.Text, VbStrConv.Narrow));
        Controls.Add(txtDepth);

        btnSearch.Location = new Point(txtDepth.Right + 10, topSearchCond - 1);
        btnSearch.Size = new Size(100, heightControls);
        btnSearch.Text = "検索";
        btnSearch.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
        btnSearch.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
        btnSearch.Click += btnSearch_Click;
        Controls.Add(btnSearch);
    }

    private void InitResult()
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
            try { Process.Start("Explorer", "/select," + this.baseDir + lvResult.Rows[e.RowIndex].Cells[0].Value.ToString()); }
            catch { }
        });
        lvResult.CellPainting += lvResult_CellPainting;
        lvResult.AllowUserToAddRows = false;
        lvResult.AllowUserToDeleteRows = false;
        lvResult.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;
        Controls.Add(lvResult);

        progress.Location = new Point(leftControls, 130);
        progress.Size = new Size(150, 20);
        progress.Style = ProgressBarStyle.Continuous;
        Controls.Add(progress);

        lblFinish.Location = new Point(leftControls + 150, 130);
        lblFinish.AutoSize = true;
        lblFinish.Text = "";
        lblFinish.Anchor = (AnchorStyles.Top | AnchorStyles.Left);
        lblFinish.Font = new Font(new FontFamily("BIZ UDゴシック"), 9);
        Controls.Add(lblFinish);
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
        if (btnSearch.Text == "停止")
        {
            cancel = true;
            return;
        }

        // Initialize
        lblFinish.Text = "";
        lvResult.Rows.Clear();
        cancel = false;
        btnSearch.Text = SwitchString(btnSearch.Text, "検索", "停止");

        // 検索条件の取り込み
        this.baseDir = txtBaseDirectory.Text == "" ? @"C:\" : txtBaseDirectory.Text;
        string pattern = txtPattern.Text == "" ? "*" : txtPattern.Text;
        int depth = txtDepth.Text == "" ? 1 : int.Parse(txtDepth.Text);

        Wildcard wc = (Wildcard)cmbCond.SelectedIndex;
        if (wc == Wildcard.Contain || wc == Wildcard.End) pattern = "*" + pattern;
        if (wc == Wildcard.Contain || wc == Wildcard.Start) pattern = pattern + "*";

        // 検索条件を保存
        reg.SetValue("BaseDirectory", txtBaseDirectory.Text);
        reg.SetValue("Pattern", txtPattern.Text);
        reg.SetValue("Depth", depth, RegistryValueKind.DWord);
        reg.SetValue("ShouldSearchFolder", Bool2Int(chkFolder.Checked), RegistryValueKind.DWord);
        reg.SetValue("ShouldSearchFile", Bool2Int(chkFile.Checked), RegistryValueKind.DWord);
        reg.SetValue("ShouldSort", Bool2Int(chkSort.Checked), RegistryValueKind.DWord);
        reg.SetValue("Wildcard", cmbCond.SelectedIndex, RegistryValueKind.DWord);

        // 検索処理を実行。
        SearchCond condition = (chkFolder.Checked ? SearchCond.Folders : SearchCond.None) |
                               (chkFile.Checked ? SearchCond.Files : SearchCond.None);
        await Task.Run(() => FindFiles(this.baseDir, this.baseDir, pattern, depth, condition));

        if (chkSort.Checked) SortItem();

        // Finalize
        SetProgress(0);
        lblFinish.Text = "検索完了";
        btnSearch.Text = SwitchString(btnSearch.Text, "検索", "停止");
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
    /// booleanを[1,0]に変換する。
    /// </summary>
    /// <param name="b">真偽値</param>
    /// <returns>true => 1, false => 0</returns>
    private int Bool2Int(bool b)
    {
        return b ? 1 : 0;
    }

    /// <summary>
    /// ファイル検索処理の実装
    /// </summary>
    /// <param name="baseDirectory">おおもとのディレクトリ</param>
    ///  <param name="searchDirectory">検索したいディレクトリ</param>
    /// <param name="pattern">検索したい文字列</param>
    /// <param name="depth">ディレクトリを何段潜って検索するか</param>
    /// <param name="condition">フォルダ・ファイルを検索するかどうかの条件</param>
    private void FindFiles(string baseDirectory, string searchDirectory, string pattern, int depth, SearchCond condition)
    {
        if (depth <= 0 || cancel || condition == SearchCond.None) return;

        List<string> queue = new List<string>();
        List<string> ret = new List<string>();

        // アクセス権限のないフォルダだとエラーが発生するのでエラーをもみ消す。
        try
        {
            // フォルダを探索キューに追加
            SetProgress(10);
            foreach (string folder in Directory.EnumerateDirectories(searchDirectory)) queue.Add(folder);

            // 条件に合うフォルダを検索結果リストに積む
            SetProgress(33);
            if ((SearchCond.Folders & condition) != SearchCond.None)
                foreach (string folder in Directory.EnumerateDirectories(searchDirectory, pattern)) ret.Add(folder.Replace(baseDirectory, ""));

            // 条件に合うファイルを検索結果リストに積む
            SetProgress(66);
            if ((SearchCond.Files & condition) != SearchCond.None)
                foreach (string file in Directory.EnumerateFiles(searchDirectory, pattern)) ret.Add(file.Replace(baseDirectory, ""));
        }
        catch { }

        SetProgress(100);
        AddRows(ret);

        Parallel.ForEach(queue, nextBaseDirectory => FindFiles(baseDirectory, nextBaseDirectory, pattern, depth - 1, condition));
    }

    /// <summary>
    /// 検索結果をUIに反映する。
    /// </summary>
    /// <param name="items">検索結果のリスト</param>
    private void AddRows(List<string> items)
    {
        if (lvResult.InvokeRequired)
            lvResult.Invoke(new Action<List<string>>(AddRows), new object[] { items });
        else
            foreach (string item in items)
                lvResult.Rows.Add(new string[] { item });
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
            // 等価なはずだが、なぜか
            // if (value != 0 && value != 100)
            // だと、意図通りに動かなかった
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

class WindowPos
{
    private int top, left, width, height;
    public WindowPos(int top, int left, int width, int height)
    {
        this.top = top < 0 ? 0 : top;
        this.left = left < 0 ? 0 : left;
        this.width = width;
        this.height = height;
    }

    public int Top { get { return top; } }
    public int Left { get { return left; } }
    public int Width { get { return width; } }
    public int Height { get { return height; } }
}

class SearchCondition
{
    void hoge()
    {
        // 検索条件の取り込み
        this.baseDir = txtBaseDirectory.Text == "" ? @"C:\" : txtBaseDirectory.Text;
        string pattern = txtPattern.Text == "" ? "*" : txtPattern.Text;
        int depth = txtDepth.Text == "" ? 1 : int.Parse(txtDepth.Text);

        Wildcard wc = (Wildcard)cmbCond.SelectedIndex;
        if (wc == Wildcard.Contain || wc == Wildcard.End) pattern = "*" + pattern;
        if (wc == Wildcard.Contain || wc == Wildcard.Start) pattern = pattern + "*";

        // 検索条件を保存
        reg.SetValue("BaseDirectory", txtBaseDirectory.Text);
        reg.SetValue("Pattern", txtPattern.Text);
        reg.SetValue("Depth", depth, RegistryValueKind.DWord);
        reg.SetValue("ShouldSearchFolder", Bool2Int(chkFolder.Checked), RegistryValueKind.DWord);
        reg.SetValue("ShouldSearchFile", Bool2Int(chkFile.Checked), RegistryValueKind.DWord);
        reg.SetValue("ShouldSort", Bool2Int(chkSort.Checked), RegistryValueKind.DWord);
        reg.SetValue("Wildcard", cmbCond.SelectedIndex, RegistryValueKind.DWord);

        // 検索処理を実行。
        SearchCond condition = (chkFolder.Checked ? SearchCond.Folders : SearchCond.None) |
                               (chkFile.Checked ? SearchCond.Files : SearchCond.None);
        await Task.Run(() => FindFiles(this.baseDir, this.baseDir, pattern, depth, condition));
    }

}

class SaveData
{
    private const string DEFAULT_POSITION = "30 30 600 500";
    private RegistryKey key = Registry.CurrentUser.CreateSubKey(@"SOFTWARE\Yamato\FindFiles");
    public static WindowPos WindowPosition
    {
        get
        {
            int[] pos = new SaveData().key.GetValue("Position", DEFAULT_POSITION).ToString().Split(' ').Select(x => int.Parse(x)).ToArray();
            return new WindowPos(pos[0], pos[1], pos[2], pos[3]);
        }
        set
        {
            string pos = String.Join(" ", new int[] { value.Top, value.Left, value.Width, value.Height }.Select(x => x.ToString()));
            new SaveData().key.SetValue("Position", pos);
        }
    }
}

// class FoundItem
// {
//     enum Attributes
//     {
//         File,
//         Folder,
//     }
//     private Attributes _attribute;
//     private string _path;

//     public Attributes Attribute { get { return _attribute; } }
//     public string Path { get { return _path; } }

//     public FoundItem(Attributes Attribute, string Path)
//     {
//         FoundItem fi = new FoundItem();
//         fi._attribute = Attribute;
//         fi._path = Path;
//     }
// }

/// <summary>
/// フォルダ選択ダイアログボックス
/// Qiitaより拾ってきた
/// https://qiita.com/otagaisama-1/items/b0804b9d6d37d82950f7
/// </summary>
public class FolderSelectDialog
{
    public string Path { get; set; }
    public string Title { get; set; }

    public System.Windows.Forms.DialogResult ShowDialog()
    {
        return ShowDialog(IntPtr.Zero);
    }

    public System.Windows.Forms.DialogResult ShowDialog(System.Windows.Forms.IWin32Window owner)
    {
        return ShowDialog(owner.Handle);
    }

    public System.Windows.Forms.DialogResult ShowDialog(IntPtr owner)
    {
        var dlg = new FileOpenDialogInternal() as IFileOpenDialog;
        try
        {
            dlg.SetOptions(FOS.FOS_PICKFOLDERS | FOS.FOS_FORCEFILESYSTEM);

            IShellItem item;
            if (!string.IsNullOrEmpty(this.Path))
            {
                IntPtr idl;
                uint atts = 0;
                if (NativeMethods.SHILCreateFromPath(this.Path, out idl, ref atts) == 0)
                {
                    if (NativeMethods.SHCreateShellItem(IntPtr.Zero, IntPtr.Zero, idl, out item) == 0)
                    {
                        dlg.SetFolder(item);
                    }
                }
            }

            if (!string.IsNullOrEmpty(this.Title))
                dlg.SetTitle(this.Title);

            var hr = dlg.Show(owner);
            if (hr.Equals(NativeMethods.ERROR_CANCELLED))
                return System.Windows.Forms.DialogResult.Cancel;
            if (!hr.Equals(0))
                return System.Windows.Forms.DialogResult.Abort;

            dlg.GetResult(out item);
            string outputPath;
            item.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, out outputPath);
            this.Path = outputPath;

            return System.Windows.Forms.DialogResult.OK;
        }
        finally
        {
            Marshal.FinalReleaseComObject(dlg);
        }
    }

    [ComImport]
    [Guid("DC1C5A9C-E88A-4dde-A5A1-60F82A20AEF7")]
    private class FileOpenDialogInternal
    {
    }

    // not fully defined と記載された宣言は、支障ない範囲で端折ってあります。
    [ComImport]
    [Guid("42f85136-db7e-439c-85f1-e4075d135fc8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IFileOpenDialog
    {
        [PreserveSig]
        UInt32 Show([In] IntPtr hwndParent);
        void SetFileTypes();     // not fully defined
        void SetFileTypeIndex();     // not fully defined
        void GetFileTypeIndex();     // not fully defined
        void Advise(); // not fully defined
        void Unadvise();
        void SetOptions([In] FOS fos);
        void GetOptions(); // not fully defined
        void SetDefaultFolder(); // not fully defined
        void SetFolder(IShellItem psi);
        void GetFolder(); // not fully defined
        void GetCurrentSelection(); // not fully defined
        void SetFileName();  // not fully defined
        void GetFileName();  // not fully defined
        void SetTitle([In, MarshalAs(UnmanagedType.LPWStr)] string pszTitle);
        void SetOkButtonLabel(); // not fully defined
        void SetFileNameLabel(); // not fully defined
        void GetResult(out IShellItem ppsi);
        void AddPlace(); // not fully defined
        void SetDefaultExtension(); // not fully defined
        void Close(); // not fully defined
        void SetClientGuid();  // not fully defined
        void ClearClientData();
        void SetFilter(); // not fully defined
        void GetResults(); // not fully defined
        void GetSelectedItems(); // not fully defined
    }

    [ComImport]
    [Guid("43826D1E-E718-42EE-BC55-A1E261C37BFE")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IShellItem
    {
        void BindToHandler(); // not fully defined
        void GetParent(); // not fully defined
        void GetDisplayName([In] SIGDN sigdnName, [MarshalAs(UnmanagedType.LPWStr)] out string ppszName);
        void GetAttributes();  // not fully defined
        void Compare();  // not fully defined
    }

    private enum SIGDN : uint // not fully defined
    {
        SIGDN_FILESYSPATH = 0x80058000,
    }

    [Flags]
    private enum FOS // not fully defined
    {
        FOS_FORCEFILESYSTEM = 0x40,
        FOS_PICKFOLDERS = 0x20,
    }

    private class NativeMethods
    {
        [DllImport("shell32.dll")]
        public static extern int SHILCreateFromPath([MarshalAs(UnmanagedType.LPWStr)] string pszPath, out IntPtr ppIdl, ref uint rgflnOut);

        [DllImport("shell32.dll")]
        public static extern int SHCreateShellItem(IntPtr pidlParent, IntPtr psfParent, IntPtr pidl, out IShellItem ppsi);

        public const uint ERROR_CANCELLED = 0x800704C7;
    }
}

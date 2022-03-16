using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Linq;
using Microsoft.Win32;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.IO;
using System.Collections.Generic;
using System;
using System.Windows.Forms;
using System.Drawing;

public class Program
{
    [STAThread]
    public static void Main()
    {
        Application.Run(new MainForm());
    }
}

public class MainForm : Form
{
    private const string SAVEDATA = @"SOFTWARE\Yamato\FindFiles";
    private RegistryKey reg = Registry.CurrentUser.CreateSubKey(SAVEDATA);

    private const int BIT_FOLDER = 1;
    private const int BIT_FILE = 2;

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

    private int heightControls = 25;
    private int leftControls = 20;
    private int topBaseDirectory = 30;
    private int topSearchCond = 90;
    private int topResult = 150;

    private bool cancel = false;

    public MainForm()
    {
        // フォームの初期化
        int[] pos = reg.GetValue("Position", "30 30 600 500")
                       .ToString()
                       .Split(' ')
                       .Select(x => int.Parse(x))
                       .ToArray();
        this.StartPosition = FormStartPosition.Manual;
        this.Location = new Point(pos[0], pos[1]);
        this.Size = new Size(pos[2], pos[3]);
        this.MinimumSize = new Size(400, 250);
        this.Text = "フォルダ・ファイル検索 - ISHII_Tools";
        this.AllowDrop = true;
        // フォームを閉じるときに位置とサイズを記録する
        this.Closing += ((object sender, System.ComponentModel.CancelEventArgs e) =>
        {
            pos = new int[] { Left, Top, Width, Height }.Select(x => x < 0 ? 0 : x).ToArray();
            reg.SetValue("Position", string.Join(" ", pos.Select(x => x.ToString())));
        });

        // コントロールの初期化
        InitBaseDirectory();
        InitSearchCond();
        InitResult();
    }

    private void InitBaseDirectory()
    {
        lblBaseDirectory.Location = new Point(leftControls, topBaseDirectory - 20);
        lblBaseDirectory.Size = new Size(370, 18);     // 文字数ベースなのでマジックナンバーでよい
        lblBaseDirectory.Text = "探索開始フォルダ(フォルダをドラッグ&ドロップで指定可)";
        lblBaseDirectory.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
        Controls.Add(lblBaseDirectory);

        txtBaseDirectory.AutoSize = false;
        txtBaseDirectory.Location = new Point(leftControls, topBaseDirectory);
        txtBaseDirectory.Size = new Size(Width - 100, heightControls);
        txtBaseDirectory.BorderStyle = BorderStyle.FixedSingle;
        txtBaseDirectory.AllowDrop = true;
        txtBaseDirectory.DragDrop += txtBaseDirectory_DragDrop;
        txtBaseDirectory.DragEnter += txtBaseDirectory_DragEnter;
        txtBaseDirectory.KeyDown += textBox_KeyDown;
        txtBaseDirectory.Anchor = (AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right);
        txtBaseDirectory.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
        txtBaseDirectory.Text = reg.GetValue("BaseDirectory", @"C:\").ToString();
        Controls.Add(txtBaseDirectory);

        btnDirectory.Location = new Point(txtBaseDirectory.Right + 10, topBaseDirectory);
        btnDirectory.Size = new Size(30, heightControls);
        btnDirectory.Text = "..";
        btnDirectory.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
        btnDirectory.Click += btnDirectory_Click;
        btnDirectory.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
        Controls.Add(btnDirectory);
    }

    private void InitSearchCond()
    {
        lblPattern.Location = new Point(leftControls, topSearchCond - 20);
        lblPattern.Size = new Size(60, 18);
        lblPattern.Text = "検索条件";
        lblPattern.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
        Controls.Add(lblPattern);

        chkFolder.Location = new Point(lblPattern.Right, topSearchCond - 20);
        chkFolder.Size = new Size(80, 18);
        chkFolder.Text = "フォルダ";
        chkFolder.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
        chkFolder.Checked = Convert.ToBoolean(reg.GetValue("ShouldSearchFolder", 1));
        Controls.Add(chkFolder);

        chkFile.Location = new Point(chkFolder.Right, topSearchCond - 20);
        chkFile.Size = new Size(80, 18);
        chkFile.Text = "ファイル";
        chkFile.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
        chkFile.Checked = Convert.ToBoolean(reg.GetValue("ShouldSearchFile", 1));
        Controls.Add(chkFile);

        chkSort.Location = new Point(chkFile.Right, topSearchCond - 20);
        chkSort.Size = new Size(160, 18);
        chkSort.Text = "検索終了後にソート";
        chkSort.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
        chkSort.Checked = Convert.ToBoolean(reg.GetValue("ShouldSort", 0));
        Controls.Add(chkSort);

        txtPattern.AutoSize = false;
        txtPattern.Location = new Point(leftControls, topSearchCond);
        txtPattern.Size = new Size(Width - 350, heightControls);
        txtPattern.BorderStyle = BorderStyle.FixedSingle;
        txtPattern.Anchor = (AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right);
        txtPattern.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
        txtPattern.Text = reg.GetValue("Pattern", "").ToString();
        txtPattern.KeyDown += textBox_KeyDown;
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
        lblDepth.Size = new Size(100, 18);
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
        txtDepth.KeyDown += textBox_KeyDown;
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
        lvResult.CellContentDoubleClick += lvResult_CellContentDoubleClick;
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
        lblFinish.Size = new Size(80, 20);
        lblFinish.Text = "";
        lblFinish.Anchor = (AnchorStyles.Top | AnchorStyles.Left);
        lblFinish.Font = new Font(new FontFamily("BIZ UDゴシック"), 9);
        Controls.Add(lblFinish);
    }

    // イベント処理
    // ディレクトリボタン。フォルダ選択ダイアログを表示する
    private void btnDirectory_Click(object sender, EventArgs e)
    {
        FolderSelectDialog dialog = new FolderSelectDialog();
        if (dialog.ShowDialog() == DialogResult.OK)
            txtBaseDirectory.Text = dialog.Path;
    }

    // 検索ボタン。検索を実行する。
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
        string BaseDirectory = txtBaseDirectory.Text == "" ? @"C:\" : txtBaseDirectory.Text;
        string Pattern = txtPattern.Text == "" ? "*" : txtPattern.Text;
        int Depth = txtDepth.Text == "" ? 0 : int.Parse(txtDepth.Text);

        switch (cmbCond.SelectedIndex)
        {
            // を含む
            case 0:
                Pattern = "*" + Pattern + "*";
                break;
            // から始まる
            case 2:
                Pattern = Pattern + "*";
                break;
            // で終わる
            case 3:
                Pattern = "*" + Pattern;
                break;
            // と一致する
            default:
                break;
        }

        // 検索条件を保存
        reg.SetValue("BaseDirectory", txtBaseDirectory.Text);
        reg.SetValue("Pattern", txtPattern.Text);
        reg.SetValue("Depth", int.Parse(txtDepth.Text), RegistryValueKind.DWord);
        reg.SetValue("ShouldSearchFolder", Bool2Int(chkFolder.Checked), RegistryValueKind.DWord);
        reg.SetValue("ShouldSearchFile", Bool2Int(chkFile.Checked), RegistryValueKind.DWord);
        reg.SetValue("ShouldSort", Bool2Int(chkSort.Checked), RegistryValueKind.DWord);
        reg.SetValue("Wildcard", cmbCond.SelectedIndex, RegistryValueKind.DWord);

        // 検索処理を実行。
        int condition = 0;
        condition += BIT_FOLDER * Bool2Int(chkFolder.Checked);
        condition += BIT_FILE * Bool2Int(chkFile.Checked);
        await Task.Run(() => FindFiles(BaseDirectory, Pattern, Depth, condition));

        if (chkSort.Checked) SortItem();

        // Finarilze
        SetProgress(0);
        lblFinish.Text = "検索完了";
        btnSearch.Text = SwitchString(btnSearch.Text, "検索", "停止");
    }

    private string SwitchString(string exp, string str1, string str2)
    {
        return exp == str1 ? str2 : str1;
    }


    private int Bool2Int(bool b)
    {
        return b ? 1 : 0;
    }


    // ファイル検索処理の実装
    private void FindFiles(string BaseDirectory, string Pattern, int Depth, int condition)
    {
        if (Depth <= 0 || cancel == true || condition == 0) return;

        List<string> Queue = new List<string>();
        List<string> Ret = new List<string>();

        // アクセス権限のないフォルダだとエラーが発生するのでエラーをもみ消す。
        try
        {
            // フォルダを探索キューに追加
            SetProgress(10);
            foreach (string folder in Directory.EnumerateDirectories(BaseDirectory)) Queue.Add(folder);

            const int BIT_FOLDER = 1;
            const int BIT_FILE = 2;
            SetProgress(33);
            if ((BIT_FOLDER & condition) != 0)
                foreach (string folder in Directory.EnumerateDirectories(BaseDirectory, Pattern)) Ret.Add(folder);

            SetProgress(66);
            if ((BIT_FILE & condition) != 0)
                foreach (string file in Directory.EnumerateFiles(BaseDirectory, Pattern)) Ret.Add(file);
        }
        catch { }

        SetProgress(100);
        AddRows(Ret);

        if (Queue.Count < 40)
            foreach (string NextBaseDirectory in Queue) FindFiles(NextBaseDirectory, Pattern, Depth - 1, condition);
        else
            Parallel.ForEach(Queue, NextBaseDirectory => FindFiles(NextBaseDirectory, Pattern, Depth - 1, condition));
    }

    private void AddRows(List<string> items)
    {
        if (lvResult.InvokeRequired)
            lvResult.Invoke(new Action<List<string>>(AddRows), new object[] { items });
        else
            foreach (string item in items)
                lvResult.Rows.Add(new string[] { item });
    }

    private void SetProgress(int value)
    {
        if (progress.InvokeRequired)
            progress.Invoke(new Action<int>(SetProgress), new object[] { value });
        else
            progress.Value = value;
    }

    private void SortItem()
    {
        lvResult.Sort(lvResult.Columns[0], System.ComponentModel.ListSortDirection.Ascending);
        lvResult.CurrentCell = lvResult[0, 0];
    }

    // 検索結果をダブルクリックしたときの処理。アイテムをハイライトした状態でエクスプローラを開く。
    private void lvResult_CellContentDoubleClick(Object sender, DataGridViewCellEventArgs e)
    {
        try { Process.Start("Explorer", "/select," + lvResult.Rows[e.RowIndex].Cells[0].Value.ToString()); }
        catch { }
    }

    // 基準ディレクトリのテキストボックスにファルダをドラッグアンドドロップしたときの処理。
    private void txtBaseDirectory_DragDrop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

        string[] dragFilePathArr = (string[])e.Data.GetData(DataFormats.FileDrop, false);
        txtBaseDirectory.Text = OriginalPath(dragFilePathArr[0]);
    }

    private string OriginalPath(string PathShortcut)
    {
        if (!PathShortcut.EndsWith(".lnk")) return PathShortcut;

        // WshShellを作成
        Type t = Type.GetTypeFromCLSID(new Guid("72C24DD5-D70A-438B-8A42-98424B88AFB8"));
        dynamic shell = Activator.CreateInstance(t);

        var shortcut = shell.CreateShortcut(PathShortcut);
        string TargetPath = shortcut.TargetPath;
        // ショートカットの先がショートカットだった場合再帰呼び出しする。
        return TargetPath.EndsWith(".lnk") ? OriginalPath(TargetPath) : TargetPath;
    }

    // ドラッグアンドドロップで離す前の処理？よくわからん。
    private void txtBaseDirectory_DragEnter(object sender, DragEventArgs e)
    {
        e.Effect = DragDropEffects.All;
    }


    // テキストボックスでCtrl+Aをした時の処理。
    private void textBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Control && e.KeyCode == Keys.A)
        {
            TextBox txt = (TextBox)sender;
            txt.SelectAll();
        }
    }

    // 結果一覧に連番を振る処理。よくわからん。
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

// フォルダ選択ダイアログボックス
// Qiitaより拾ってきた
// https://qiita.com/otagaisama-1/items/b0804b9d6d37d82950f7
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

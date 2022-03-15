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

    public static string OriginalPath(string PathShortcut)
    {
        if (!PathShortcut.EndsWith(".lnk")) return PathShortcut;

        // WshShellを作成
        Type t = Type.GetTypeFromCLSID(new Guid("72C24DD5-D70A-438B-8A42-98424B88AFB8"));
        dynamic shell = Activator.CreateInstance(t);

        //WshShortcutを作成
        var shortcut = shell.CreateShortcut(PathShortcut);

        // ショートカットがさしているパスを取得
        string TargetPath = shortcut.TargetPath;

        // ショートカットの先がショートカットだった場合再帰呼び出しする。
        return TargetPath.EndsWith(".lnk") ? OriginalPath(TargetPath) : TargetPath;
    }
}

public class MainForm : Form
{
    private const string SAVEDATA = @"SOFTWARE\Yamato\FindFiles";

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

    private int ControlsHeight = 25;
    private int ControlsLeft = 20;
    private int BaseDirectoryTop = 30;
    private int SearchCondTop = 90;
    private int ResultTop = 150;

    private bool Cancel = false;

    // コンストラクタ
    public MainForm()
    {
        this.Text = "フォルダ・ファイル検索 - ISHII_Tools";
        this.Size = new Size(600, 510);
        this.AllowDrop = true;

        InitBaseDirectory();
        InitSearchCond();
        InitResult();

        // 前回の検索条件を読み込む
        RegistryKey reg = Registry.CurrentUser.OpenSubKey(SAVEDATA);
        if (reg != null)
        {
            this.txtBaseDirectory.Text = reg.GetValue("BaseDirectory").ToString();
            this.txtPattern.Text = reg.GetValue("Pattern").ToString();
            this.txtDepth.Text = reg.GetValue("Depth").ToString();
        }
    }

    // コントロールの初期化処理
    private void InitBaseDirectory()
    {
        this.lblBaseDirectory.Location = new Point(ControlsLeft, BaseDirectoryTop - 20);
        this.lblBaseDirectory.Size = new Size(500, 18);
        this.lblBaseDirectory.Text = "探索開始フォルダ(フォルダをドラッグ&ドロップで指定可)";
        this.lblBaseDirectory.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
        this.Controls.Add(lblBaseDirectory);

        this.txtBaseDirectory.AutoSize = false;
        this.txtBaseDirectory.Location = new Point(ControlsLeft, BaseDirectoryTop);
        this.txtBaseDirectory.Size = new Size(510, ControlsHeight);
        this.txtBaseDirectory.BorderStyle = BorderStyle.FixedSingle;
        this.txtBaseDirectory.AllowDrop = true;
        this.txtBaseDirectory.DragDrop += txtBaseDirectory_DragDrop;
        this.txtBaseDirectory.DragEnter += txtBaseDirectory_DragEnter;
        this.txtBaseDirectory.KeyDown += textBox_KeyDown;
        this.txtBaseDirectory.Anchor = (AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right);
        this.txtBaseDirectory.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
        this.Controls.Add(txtBaseDirectory);

        this.btnDirectory.Location = new Point(530, BaseDirectoryTop);
        this.btnDirectory.Size = new Size(30, ControlsHeight);
        this.btnDirectory.Text = "..";
        this.btnDirectory.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
        this.btnDirectory.Click += btnDirectory_Click;
        this.btnDirectory.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
        this.Controls.Add(btnDirectory);
    }

    private void InitSearchCond()
    {
        this.lblPattern.Location = new Point(ControlsLeft, SearchCondTop - 20);
        this.lblPattern.Size = new Size(60, 18);
        this.lblPattern.Text = "検索条件";
        this.lblPattern.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
        this.Controls.Add(lblPattern);

        this.chkFolder.Location = new Point(ControlsLeft + 60, SearchCondTop - 20);
        this.chkFolder.Size = new Size(80, 18);
        this.chkFolder.Text = "フォルダ";
        this.chkFolder.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
        this.chkFolder.Checked = true;
        this.Controls.Add(chkFolder);

        this.chkFile.Location = new Point(ControlsLeft + 140, SearchCondTop - 20);
        this.chkFile.Size = new Size(80, 18);
        this.chkFile.Text = "ファイル";
        this.chkFile.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
        this.chkFile.Checked = true;
        this.Controls.Add(chkFile);

        this.chkSort.Location = new Point(ControlsLeft + 220, SearchCondTop - 20);
        this.chkSort.Size = new Size(160, 18);
        this.chkSort.Text = "検索終了後にソート";
        this.chkSort.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
        this.chkSort.Checked = false;
        this.Controls.Add(chkSort);

        this.txtPattern.AutoSize = false;
        this.txtPattern.Location = new Point(ControlsLeft, SearchCondTop);
        this.txtPattern.Size = new Size(250, ControlsHeight);
        this.txtPattern.BorderStyle = BorderStyle.FixedSingle;
        this.txtPattern.Anchor = (AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right);
        this.txtPattern.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
        this.txtPattern.KeyDown += textBox_KeyDown;
        this.Controls.Add(txtPattern);

        this.cmbCond.Location = new Point(275, SearchCondTop);
        this.cmbCond.Size = new Size(120, ControlsHeight);
        this.cmbCond.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
        this.cmbCond.Items.AddRange(new string[] { "を含む", "と一致する", "から始まる", "で終わる" });
        this.cmbCond.SelectedIndex = 0;
        this.cmbCond.IntegralHeight = true;
        this.cmbCond.FlatStyle = FlatStyle.Flat;
        this.cmbCond.DropDownStyle = ComboBoxStyle.DropDownList;
        this.cmbCond.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
        this.Controls.Add(cmbCond);

        this.lblDepth.Location = new Point(400, SearchCondTop - 20);
        this.lblDepth.Size = new Size(100, 18);
        this.lblDepth.Text = "探索する深さ";
        this.lblDepth.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
        this.lblDepth.Font = new Font(new FontFamily("BIZ UDゴシック"), 10);
        this.Controls.Add(lblDepth);

        this.txtDepth.AutoSize = false;
        this.txtDepth.Location = new Point(400, SearchCondTop);
        this.txtDepth.Size = new Size(50, ControlsHeight);
        this.txtDepth.BorderStyle = BorderStyle.FixedSingle;
        this.txtDepth.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
        this.txtDepth.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
        this.txtDepth.KeyDown += textBox_KeyDown;
        this.Controls.Add(txtDepth);

        this.btnSearch.Location = new Point(460, SearchCondTop - 1);
        this.btnSearch.Size = new Size(100, ControlsHeight);
        this.btnSearch.Text = "検索";
        this.btnSearch.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
        this.btnSearch.Font = new Font(new FontFamily("BIZ UDゴシック"), 12);
        this.btnSearch.Click += btnSearch_Click;
        this.Controls.Add(btnSearch);
    }

    private void InitResult()
    {
        this.lvResult.Location = new Point(ControlsLeft, ResultTop);
        this.lvResult.Size = new Size(540, 310);
        this.lvResult.BorderStyle = BorderStyle.FixedSingle;
        this.lvResult.Anchor = (AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right);
        this.lvResult.ReadOnly = true;   // セルをクリックしたときに編集モードに入らないようにする
        this.lvResult.MultiSelect = false;
        this.lvResult.ColumnCount = 1;
        this.lvResult.Columns[0].Name = "結果一覧";
        this.lvResult.Columns[0].MinimumWidth = 1350;
        this.lvResult.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        this.lvResult.CellContentDoubleClick += lvResult_CellContentDoubleClick;
        this.lvResult.CellPainting += lvResult_CellPainting;
        this.lvResult.AllowUserToAddRows = false;
        this.lvResult.AllowUserToDeleteRows = false;
        this.lvResult.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;
        this.Controls.Add(lvResult);

        this.progress.Location = new Point(ControlsLeft, 130);
        this.progress.Size = new Size(150, 20);
        this.progress.Style = ProgressBarStyle.Continuous;
        this.Controls.Add(progress);

        this.lblFinish.Location = new Point(ControlsLeft + 150, 130);
        this.lblFinish.Size = new Size(80, 20);
        this.lblFinish.Text = "";
        this.lblFinish.Anchor = (AnchorStyles.Top | AnchorStyles.Left);
        this.lblFinish.Font = new Font(new FontFamily("BIZ UDゴシック"), 9);
        this.Controls.Add(lblFinish);
    }

    // イベント処理
    // ディレクトリボタン。フォルダ選択ダイアログを表示する
    private void btnDirectory_Click(object sender, EventArgs e)
    {
        FolderSelectDialog dialog = new FolderSelectDialog();
        if (dialog.ShowDialog() == DialogResult.OK)
            this.txtBaseDirectory.Text = dialog.Path;
    }

    // 検索ボタン。検索を実行する。
    private void btnSearch_Click(object sender, EventArgs e)
    {
        if (this.btnSearch.Text == "停止")
        {
            this.Cancel = true;
            return;
        }

        this.btnSearch.Text = this.btnSearch.Text == "検索" ? "停止" : "検索";

        this.lblFinish.Text = "";

        this.lvResult.Rows.Clear();
        Application.DoEvents();

        this.Cancel = false;

        string BaseDirectory = this.txtBaseDirectory.Text;
        string Pattern = this.txtPattern.Text;
        int Depth = this.txtDepth.Text == "" ? 0 : int.Parse(this.txtDepth.Text);

        if (BaseDirectory == "") BaseDirectory = @"C:\";
        if (Pattern == "") Pattern = "*";
        if (Depth <= 0) Depth = 1;

        if (this.cmbCond.Text == "を含む") Pattern = "*" + Pattern + "*";
        if (this.cmbCond.Text == "から始まる") Pattern = Pattern + "*";
        if (this.cmbCond.Text == "で終わる") Pattern = "*" + Pattern;

        // 検索条件を保存
        RegistryKey reg = Registry.CurrentUser.CreateSubKey(SAVEDATA);
        reg.SetValue("BaseDirectory", this.txtBaseDirectory.Text);
        reg.SetValue("Pattern", this.txtPattern.Text);
        reg.SetValue("Depth", txtDepth.Text);

        // 検索処理を実行。
        const int BIT_FOLDER = 1;
        const int BIT_FILE = 2;
        int condition = 0;
        if (this.chkFolder.Checked) condition += BIT_FOLDER;
        if (this.chkFile.Checked) condition += BIT_FILE;

        this.FindFiles(BaseDirectory, Pattern, Depth, condition);

        if (this.chkSort.Checked) this.SortItem();

        // Finarilze
        this.progress.Value = 0;
        this.lblFinish.Text = "検索完了";
        this.btnSearch.Text = this.btnSearch.Text == "検索" ? "停止" : "検索";
    }

    // 検索結果をダブルクリックしたときの処理。アイテムをハイライトした状態でエクスプローラを開く。
    private void lvResult_CellContentDoubleClick(Object sender, DataGridViewCellEventArgs e)
    {
        try
        {
            Process.Start("Explorer", "/select," + this.lvResult.Rows[e.RowIndex].Cells[0].Value.ToString());
        }
        catch { }
    }

    // 基準ディレクトリのテキストボックスにファルダをドラッグアンドドロップしたときの処理。
    private void txtBaseDirectory_DragDrop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

        string[] dragFilePathArr = (string[])e.Data.GetData(DataFormats.FileDrop, false);
        txtBaseDirectory.Text = Program.OriginalPath(dragFilePathArr[0]);
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

    // ファイル検索処理の実装
    private void FindFiles(string BaseDirectory, string Pattern, int Depth, int condition)
    {
        if (Depth <= 0 || this.Cancel == true || condition == 0) return;

        List<string> Queue = new List<string>();

        lock (this.progress) this.progress.Value = 33;

        // アクセス権限のないフォルダだとエラーが発生するのでエラーをもみ消す。
        try
        {
            // フォルダを探索キューに追加
            foreach (string folder in Directory.EnumerateDirectories(BaseDirectory)) Queue.Add(folder);

            int i = 0;
            Action<string> AddRow = (item =>
            {
                this.lvResult.Rows.Add(new string[] { item });

                if (i++ % 100 == 0)
                {
                    this.progress.Value++;
                    Application.DoEvents();
                }
            });

            const int BIT_FOLDER = 1;
            const int BIT_FILE = 2;
            if (Convert.ToBoolean(BIT_FOLDER & condition))
            {
                foreach (string folder in Directory.EnumerateDirectories(BaseDirectory, Pattern))
                {
                    lock (this.lvResult)
                    {
                        AddRow(folder);
                    }
                }
            }

            i = 0;
            if (Convert.ToBoolean(BIT_FILE & condition))
            {
                foreach (string file in Directory.EnumerateFiles(BaseDirectory, Pattern))
                {
                    lock (this.lvResult)
                    {
                        AddRow(file);
                    }
                }
            }
        }
        catch { }

        Application.DoEvents();

        if (Queue.Count < 40)
            foreach (string NextBaseDirectory in Queue) FindFiles(NextBaseDirectory, Pattern, Depth - 1, condition);
        else
            Parallel.ForEach(Queue, NextBaseDirectory => FindFiles(NextBaseDirectory, Pattern, Depth - 1, condition));
    }

    private void SortItem()
    {
        this.lvResult.Sort(this.lvResult.Columns[0], System.ComponentModel.ListSortDirection.Ascending);
        this.lvResult.CurrentCell = this.lvResult[0, 0];
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

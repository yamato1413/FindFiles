using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Configuration;
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

        // ??????????????????????????????????????????????????????????????????
        this.Closing += ((object sender, System.ComponentModel.CancelEventArgs e) =>
        {
            SaveData.WindowPosition = new WindowPos(this.Top, this.Left, this.Width, this.Height);
        });

        this.Text = "????????????????????????????????? - ISHII_Tools";
        this.AllowDrop = true;
        this.Shown += ((object sender, EventArgs e) => txtKeyword.Focus());
    }

    private void init_lblBaseDirectory()
    {
        lblBaseDirectory.Location = new Point(leftControls, topBaseDirectory - 20);
        lblBaseDirectory.AutoSize = true;
        lblBaseDirectory.Text = "????????????????????????(???????????????????????????&????????????????????????)";
        lblBaseDirectory.Font = new Font(new FontFamily("BIZ UD????????????"), 10);
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
        txtBaseDirectory.Font = new Font(new FontFamily("BIZ UD????????????"), 12);
        Controls.Add(txtBaseDirectory);
    }

    private void init_btnDirectory()
    {
        btnDirectory.Location = new Point(txtBaseDirectory.Right + 10, topBaseDirectory);
        btnDirectory.Size = new Size(30, heightControls);
        btnDirectory.Text = "..";
        btnDirectory.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
        btnDirectory.Click += SelectDirectory;
        btnDirectory.Font = new Font(new FontFamily("BIZ UD????????????"), 12);
        Controls.Add(btnDirectory);
    }

    private void init_lblKeyword()
    {
        lblKeyword.Location = new Point(leftControls, topSearchCond - 20);
        lblKeyword.AutoSize = true;
        lblKeyword.Text = "????????????";
        lblKeyword.Font = new Font(new FontFamily("BIZ UD????????????"), 10);
        Controls.Add(lblKeyword);
    }

    private void init_chkFolder()
    {
        chkFolder.Location = new Point(lblKeyword.Right, topSearchCond - 20);
        chkFolder.AutoSize = true;
        chkFolder.Text = "????????????";
        chkFolder.Font = new Font(new FontFamily("BIZ UD????????????"), 10);
        Controls.Add(chkFolder);
    }

    private void init_chkFile()
    {
        chkFile.Location = new Point(chkFolder.Right, topSearchCond - 20);
        chkFile.AutoSize = true;
        chkFile.Text = "????????????";
        chkFile.Font = new Font(new FontFamily("BIZ UD????????????"), 10);
        Controls.Add(chkFile);
    }

    private void init_chkSort()
    {
        chkSort.Location = new Point(chkFile.Right, topSearchCond - 20);
        chkSort.AutoSize = true;
        chkSort.Text = "???????????????????????????";
        chkSort.Font = new Font(new FontFamily("BIZ UD????????????"), 10);
        Controls.Add(chkSort);
    }

    private void init_txtKeyword()
    {
        txtKeyword.AutoSize = false;
        txtKeyword.Location = new Point(leftControls, topSearchCond);
        txtKeyword.Size = new Size(Width - 350, heightControls);
        txtKeyword.BorderStyle = BorderStyle.FixedSingle;
        txtKeyword.Anchor = (AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right);
        txtKeyword.Font = new Font(new FontFamily("BIZ UD????????????"), 12);
        txtKeyword.KeyDown += SelectAllText;
        Controls.Add(txtKeyword);
    }

    private void init_cmbWildcard()
    {
        cmbWildcard.Location = new Point(txtKeyword.Right + 5, topSearchCond);
        cmbWildcard.Size = new Size(120, heightControls);
        cmbWildcard.Font = new Font(new FontFamily("BIZ UD????????????"), 12);
        cmbWildcard.Items.AddRange(new string[] { "?????????", "???????????????", "???????????????", "????????????" });
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
        lblDepth.Text = "??????????????????";
        lblDepth.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
        lblDepth.Font = new Font(new FontFamily("BIZ UD????????????"), 10);
        Controls.Add(lblDepth);
    }

    private void init_txtDepth()
    {
        txtDepth.AutoSize = false;
        txtDepth.Location = new Point(cmbWildcard.Right + 5, topSearchCond);
        txtDepth.Size = new Size(50, heightControls);
        txtDepth.BorderStyle = BorderStyle.FixedSingle;
        txtDepth.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
        txtDepth.Font = new Font(new FontFamily("BIZ UD????????????"), 12);
        txtDepth.KeyDown += SelectAllText;
        txtDepth.TextChanged += ((object sender, EventArgs e) => txtDepth.Text = Strings.StrConv(txtDepth.Text, VbStrConv.Narrow));
        Controls.Add(txtDepth);
    }

    private void init_btnSearch()
    {
        btnSearch.Location = new Point(txtDepth.Right + 10, topSearchCond - 1);
        btnSearch.Size = new Size(100, heightControls);
        btnSearch.Text = "??????";
        btnSearch.Anchor = (AnchorStyles.Top | AnchorStyles.Right);
        btnSearch.Font = new Font(new FontFamily("BIZ UD????????????"), 12);
        btnSearch.Click += btnSearch_Click;
        Controls.Add(btnSearch);
    }

    private void init_lvResult()
    {
        lvResult.Location = new Point(leftControls, topResult);
        lvResult.Size = new Size(Width - 60, Height - 200);
        lvResult.BorderStyle = BorderStyle.FixedSingle;
        lvResult.Anchor = (AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right);
        lvResult.ReadOnly = true;   // ?????????????????????????????????????????????????????????????????????????????????
        lvResult.MultiSelect = false;
        lvResult.ColumnCount = 1;
        lvResult.Columns[0].Name = "????????????";
        lvResult.Columns[0].MinimumWidth = 2400;
        lvResult.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        lvResult.CellContentDoubleClick += ((Object sender, DataGridViewCellEventArgs e) =>
        {
            try { Process.Start("Explorer", "/select," + SaveData.SearchCondition.BaseDirectory + lvResult.Rows[e.RowIndex].Cells[0].Value.ToString()); }
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
        lblFinish.Font = new Font(new FontFamily("BIZ UD????????????"), 9);
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
    /// ???????????????????????????????????????????????????
    /// </summary>
    private void SelectDirectory(object sender, EventArgs e)
    {
        FolderSelectDialog dialog = new FolderSelectDialog();
        if (dialog.ShowDialog() == DialogResult.OK) txtBaseDirectory.Text = dialog.Path;
    }

    /// <summary>
    /// ????????????????????????
    /// </summary>
    private async void btnSearch_Click(object sender, EventArgs e)
    {
        // ??????????????????????????????????????????????????????
        if (btnSearch.Text == "??????")
        {
            cancel = true;
            return;
        }

        // ?????????
        int depth;
        if (!int.TryParse(txtDepth.Text, out depth))
        {
            MessageBox.Show("????????????????????????0?????????????????????????????????????????????");
            return;
        }

        // Initialize
        lblFinish.Text = "";
        lvResult.Rows.Clear();
        cancel = false;
        btnSearch.Text = SwitchString(btnSearch.Text, "??????", "??????");

        // ?????????????????????
        SearchCondition sc = new SearchCondition();
        sc.BaseDirectory = txtBaseDirectory.Text;
        sc.Keyword = txtKeyword.Text;
        sc.Depth = depth;
        sc.SearchFolder = chkFolder.Checked;
        sc.SearchFile = chkFile.Checked;
        sc.Sort = chkSort.Checked;
        sc.Wildcard = (SearchCondition.Wildcards)cmbWildcard.SelectedIndex; ;
        SaveData.SearchCondition = sc;

        // ????????????????????????
        await Task.Run(() => FindFiles(sc.BaseDirectory, sc.Depth, sc));

        if (chkSort.Checked) SortItem();

        // Finalize
        SetProgress(0);
        lblFinish.Text = "????????????";
        btnSearch.Text = SwitchString(btnSearch.Text, "??????", "??????");
    }

    /// <summary>
    /// ?????????????????????????????????????????????????????????????????????
    /// str1,str2?????????????????????????????????""??????????????????
    /// </summary>
    /// <param name="exp">??????</param>
    /// <param name="str1">????????????????????????1</param>
    /// <param name="str2">????????????????????????2</param>
    /// <returns>str1 => str2, str2 => str1, _ => ""</returns>
    private string SwitchString(string exp, string str1, string str2)
    {
        if (exp == str1) return str2;
        else if (exp == str2) return str1;
        else return "";
    }

    /// <summary>
    /// ?????????????????????????????????
    /// </summary>
    ///  <param name="searchDirectory">?????????????????????????????????</param>
    /// <param name="depth">???????????????????????????????????????????????????</param>
    /// <param name="condition">????????????</param>
    private void FindFiles(string searchDirectory, int depth, SearchCondition condition)
    {
        if (cancel) return;
        if (!condition.SearchFolder && !condition.SearchFile) return;

        List<string> queue = new List<string>();
        List<string> ret = new List<string>();

        // ??????????????????????????????????????????????????????????????????????????????????????????????????????
        try
        {
            // ???????????????????????????????????????
            SetProgress(10);
            foreach (string folder in Directory.EnumerateDirectories(searchDirectory))
            {
                queue.Add(folder);
            }

            // ????????????????????????????????????????????????????????????
            SetProgress(33);
            if (condition.SearchFolder)
            {
                foreach (string folder in Directory.EnumerateDirectories(searchDirectory, condition.Keyword_AppliedWildcard))
                {
                    ret.Add(folder);
                }
            }
            // ????????????????????????????????????????????????????????????
            SetProgress(66);
            if (condition.SearchFile)
            {
                foreach (string file in Directory.EnumerateFiles(searchDirectory, condition.Keyword_AppliedWildcard))
                {
                    ret.Add(file);
                }
            }
        }
        catch { }

        SetProgress(100);
        AddRows(ret);

        if (depth > 1)
        {
            Parallel.ForEach(queue, nextSearchDirectory => FindFiles(nextSearchDirectory, depth - 1, condition));
        }
    }

    /// <summary>
    /// ???????????????UI??????????????????
    /// </summary>
    /// <param name="items">????????????????????????</param>
    private void AddRows(List<string> items)
    {
        if (lvResult.InvokeRequired)
        {
            lvResult.Invoke(new Action<List<string>>(AddRows), new object[] { items });
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
    /// ??????????????????????????????????????????????????????
    /// </summary>
    /// <param name="value">??????????????????????????????????????????</param>
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
    /// ???????????????????????????????????????
    /// </summary>
    private void SortItem()
    {
        if (lvResult.Rows.Count == 0) return;
        lvResult.Sort(lvResult.Columns[0], System.ComponentModel.ListSortDirection.Ascending);
        lvResult.CurrentCell = lvResult[0, 0];
    }

    /// <summary>
    /// ????????????????????????????????????????????????????????????????????????
    /// </summary>
    private void txtBaseDirectory_DragDrop(object sender, DragEventArgs e)
    {
        if (!e.Data.GetDataPresent(DataFormats.FileDrop)) return;

        string[] dragFilePathArr = (string[])e.Data.GetData(DataFormats.FileDrop, false);
        txtBaseDirectory.Text = OriginalPath(dragFilePathArr[0]);
    }

    /// <summary>
    /// ?????????????????????????????????????????????????????????
    /// ?????????????????????????????????????????????????????????????????????????????????
    /// </summary>
    /// <param name="shortcutPath">??????????????????????????????</param>
    /// <returns>???????????????????????????????????????????????????</returns>
    private string OriginalPath(string shortcutPath)
    {
        if (!shortcutPath.EndsWith(".lnk")) return shortcutPath;

        // WshShell?????????
        Type t = Type.GetTypeFromCLSID(new Guid("72C24DD5-D70A-438B-8A42-98424B88AFB8"));
        dynamic shell = Activator.CreateInstance(t);

        var shortcut = shell.CreateShortcut(shortcutPath);
        string targetPath = shortcut.targetPath;
        // ?????????????????????????????????????????????????????????????????????????????????????????????
        return targetPath.EndsWith(".lnk") ? OriginalPath(targetPath) : targetPath;
    }

    /// <summary>
    /// ??????????????????????????????Ctrl+A????????????????????????
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
    /// ?????????????????????????????????????????????????????????????????????
    /// </summary>
    private void lvResult_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
    {
        //????????????????????????????????????
        if (e.ColumnIndex < 0 && e.RowIndex >= 0)
        {
            //?????????????????????
            e.Paint(e.ClipBounds, DataGridViewPaintParts.All);

            //?????????????????????????????????????????????
            //e.AdvancedBorderStyle???e.CellStyle.Padding????????????????????????
            Rectangle indexRect = e.CellBounds;
            indexRect.Inflate(-2, -2);
            //????????????????????????
            TextRenderer.DrawText(e.Graphics,
                (e.RowIndex + 1).ToString(),
                e.CellStyle.Font,
                indexRect,
                e.CellStyle.ForeColor,
                TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
            //??????????????????????????????????????????
            e.Handled = true;
        }
    }
}

class WindowPos
{
    private int top, left, width, height;

    public const string DEFAULT_POSITION = "30 30 600 500";

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
        get
        {
            return _BaseDirectory == "" ? DEFAULT_BASEDIRECTORY : _BaseDirectory;
        }
        set { _BaseDirectory = value; }
    }

    private string _Keyword_AppliedWildcard = DEFAULT_KEYWORD;
    public string Keyword_AppliedWildcard
    {
        get { return _Keyword_AppliedWildcard; }
    }

    private string _Keyword = DEFAULT_KEYWORD;
    public string Keyword
    {
        get { return _Keyword; }
        set
        {
            _Keyword = value;
            applyWildcard();
        }
    }

    private int _Depth = DEFAULT_DEPTH;
    public int Depth
    {
        get { return _Depth; }
        set
        {
            if (value < 1) throw new Exception("1????????????????????????????????????");
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
        set
        {
            _Wildcard = value;
            applyWildcard();
        }
    }

    private void applyWildcard()
    {
        _Keyword_AppliedWildcard = _Keyword;
        if (Wildcard == Wildcards.Contain || Wildcard == Wildcards.End) _Keyword_AppliedWildcard = "*" + _Keyword_AppliedWildcard;
        if (Wildcard == Wildcards.Contain || Wildcard == Wildcards.Start) _Keyword_AppliedWildcard = _Keyword_AppliedWildcard + "*";
    }

    public SearchCondition Clone()
    {
        SearchCondition newsc = new SearchCondition();
        newsc._BaseDirectory = this._BaseDirectory;
        newsc._Depth = this._Depth;
        newsc._Keyword = this._Keyword;
        newsc._SearchFile = this._SearchFile;
        newsc._SearchFolder = this._SearchFolder;
        newsc._Sort = this._Sort;
        newsc._Wildcard = this._Wildcard;
        return newsc;
    }
}

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

/// <summary>
/// ?????????????????????????????????????????????
/// Qiita?????????????????????
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

    // not fully defined ??????????????????????????????????????????????????????????????????????????????
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

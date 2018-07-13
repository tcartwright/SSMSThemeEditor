using SSMSThemeEditor.Properties;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

namespace SSMSThemeEditor
{
    public partial class frmMain : Form
    {
        private string _settingsInitialDirectory;
        private string _saveFile;
        private bool _isDirty = false;
        private LimitedSizeStack<ChangeItem> _undoStack = new LimitedSizeStack<ChangeItem>(50);
        private LimitedSizeStack<ChangeItem> _redoStack = new LimitedSizeStack<ChangeItem>(50);
        private List<string> _changes = new List<string>();
        private readonly AutoCompleteStringCollection _fonts = new AutoCompleteStringCollection();
        private readonly AutoCompleteStringCollection _fontSizes = new AutoCompleteStringCollection { "8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24", "26", "28", "36", "72" };
        private string _defaultFontName = string.Empty;
        private string _defaultFontSize = string.Empty;
        private string _fontName = string.Empty;
        private string _fontSize = string.Empty;
        private readonly StringComparer _comparer = StringComparer.InvariantCultureIgnoreCase;

        #region events
        public frmMain()
        {
            InitializeComponent();
        }
        private void frmMain_Load(object sender, EventArgs e)
        {
            foreach (FontFamily font in System.Drawing.FontFamily.Families)
            {
                _fonts.Add(font.Name);
            }
            cboFont.DataSource = _fonts;
            cboFont.AutoCompleteCustomSource = _fonts;
            cboFontSize.DataSource = _fontSizes;
            cboFontSize.AutoCompleteCustomSource = _fontSizes;
            _defaultFontName = _fontName = ConfigurationManager.AppSettings["DefaultFontName"];
            _defaultFontSize = _fontSize = ConfigurationManager.AppSettings["DefaultFontSize"];
            SetFont(_fontName, _fontSize);

            ModifyButtonEvents(tlpGeneral);
            ModifyButtonEvents(tlpSQL);
            ModifyButtonEvents(tlpXML);
            webBrowser1.DocumentText = Resources.SQLSample;
            SetupColorLabels(tlpGeneral);
            SetupColorLabels(tlpSQL);
            SetupColorLabels(tlpXML);
        }
        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (SaveChanges())
            {
                e.Cancel = true;
                return;
            }

            ModifyButtonEvents(tlpGeneral, false);
            ModifyButtonEvents(tlpSQL, false);
            ModifyButtonEvents(tlpXML, false);
        }
        private void ColorButton_Click(object sender, EventArgs e)
        {
            var btn = sender as Button;
            var tblPanel = btn.Parent as TableLayoutPanel;
            var col = tblPanel.GetColumn(btn);
            var row = tblPanel.GetRow(btn);

            if (tblPanel.GetControlFromPosition(col - 1, row) is Label lbl)
            {
                colorDialog1.Color = lbl.BackColor;
                if (colorDialog1.ShowDialog() == DialogResult.OK && lbl.BackColor != colorDialog1.Color)
                {
                    var ci = GetChangeItem(tblPanel, lbl, colorDialog1.Color, col, row);
                    _undoStack.Push(ci);
                    SetLabelColor(lbl, colorDialog1.Color);
                }
            }
        }
        private void btnLoad_Click(object sender, EventArgs e)
        {
            if (SaveChanges()) { return; }

            if (string.IsNullOrWhiteSpace(_settingsInitialDirectory))
            {
                _settingsInitialDirectory = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Sample Themes");
            }

            var openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = _settingsInitialDirectory,
                Filter = "Visual Studio Settings (*.vssettings)|*.vssettings",
                RestoreDirectory = true
            };

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var doc = new XmlDocument();
                    doc.Load(openFileDialog1.FileName);
                    GetFontInfo(doc);

                    BlankLabels();
                    LoadColors(doc, tlpGeneral);
                    LoadColors(doc, tlpSQL);
                    LoadColors(doc, tlpXML);

                    _saveFile = openFileDialog1.FileName;
                    _settingsInitialDirectory = Path.GetDirectoryName(openFileDialog1.FileName);
                    FormHasChanged(false);
                    ClearUndo();
                    this.Text = $"{Resources.AppTitle} - {Path.GetFileName(openFileDialog1.FileName)}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void btnNew_Click(object sender, EventArgs e)
        {
            if (SaveChanges()) { return; }

            this.Text = Resources.AppTitle;
            BlankLabels();
            ClearUndo();
            SetFont(_defaultFontName, _defaultFontSize);
            btnSave.Enabled = _isDirty = false;
            _saveFile = string.Empty;
            webBrowser1.DocumentText = Resources.SQLSample;
        }
        private void btnSave_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Visual Studio Settings (*.vssettings)|*.vssettings";
            saveFileDialog1.Title = "Save an SSMS Settings File";
            saveFileDialog1.InitialDirectory = _settingsInitialDirectory;
            if (!string.IsNullOrWhiteSpace(_saveFile))
            {
                saveFileDialog1.FileName = Path.GetFileName(_saveFile);
            }
            else
            {
                saveFileDialog1.FileName = string.Empty;
            }
            saveFileDialog1.OverwritePrompt = true;
            saveFileDialog1.CheckPathExists = true;

            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var items = BuildItems();

                    var settings = Resources.VSSettingsTemplate
                        .Replace("<!--<<ITEMS>>-->", items)
                        .Replace("{FONTNAME}", _fontName)
                        .Replace("{FONTSIZE}", _fontSize);

                    File.WriteAllText(saveFileDialog1.FileName, settings);

                    ClearUndo();

                    btnSave.Enabled = _isDirty = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void tsmClearColor_Click(object sender, EventArgs e)
        {
            var tsm = ((ToolStripMenuItem)sender);
            if ((tsm.GetCurrentParent() as ContextMenuStrip).SourceControl is Label lbl && lbl.BackColor != Color.Transparent)
            {
                var tlp = (TableLayoutPanel)lbl.Parent;
                var row = tlp.GetRow(lbl);
                var col = tlp.GetColumn(lbl);

                var ci = GetChangeItem(tlp, lbl, Color.Transparent, col, row);
                _undoStack.Push(ci);
                SetLabelColor(lbl, Color.Transparent);
            }
        }
        private void tsmClearColors_Click(object sender, EventArgs e)
        {
            var tsm = ((ToolStripMenuItem)sender);
            if ((tsm.GetCurrentParent() as ContextMenuStrip).SourceControl is TabPage tp)
            {
                if (MessageBox.Show("Are you sure you want to clear the colors for this tab?", "Clear Colors", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    var hasChanged = false;
                    if (tp == tabPageGeneral)
                    {
                        hasChanged = BlankLabels(tlpGeneral, true);
                    }
                    else if (tp == tabPageSQL)
                    {
                        hasChanged = BlankLabels(tlpSQL, true);
                    }
                    else if (tp == tabPageXML)
                    {
                        hasChanged = BlankLabels(tlpXML, true);
                    }
                    if (hasChanged)
                    {
                        FormHasChanged();
                    }
                }
            }
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.Z))
            {
                if (_undoStack.Count > 0)
                {
                    var ci = _undoStack.Pop();
                    if (Control.FromHandle(ci.LabelPtr) is Label lbl)
                    {
                        UpdateChanges(null);
                        SetLabelColor(lbl, ci.OldLabelColor, _undoStack.Count > 0);
                    }
                    _redoStack.Push(ci);
                }

                if (_undoStack.Count == 0)
                {
                    txtChanges.Text = string.Empty;
                }
                return true;
            }
            else if (keyData == (Keys.Control | Keys.Y))
            {
                if (_redoStack.Count > 0)
                {
                    var ci = _redoStack.Pop();
                    if (Control.FromHandle(ci.LabelPtr) is Label lbl)
                    {
                        UpdateChanges($"{ci.Style} -> {ColorToRGBString(ci.OldLabelColor)} TO {ColorToRGBString(ci.NewLabelColor)}");
                        SetLabelColor(lbl, ci.NewLabelColor);
                    }
                    _undoStack.Push(ci);
                }
                return true;
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }
        private void cboFont_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (!_comparer.Equals(cboFont.Text, _fontName))
            {
                _fontName = cboFont.Text;
                FormHasChanged();
            }
        }
        private void cboFontSize_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (!_comparer.Equals(cboFontSize.Text, _fontSize))
            {
                _fontSize = cboFontSize.Text;
                FormHasChanged();
            }
        }
        #endregion events

        #region private methods
        private void UpdateChanges(string msg)
        {
            if (msg == null)
            {
                _changes.Remove(_changes.Last());
            }
            else
            {
                _changes.Add(msg);
            }
            txtChanges.Text = string.Join("\r\n", _changes.Select((d, i) => $"[{i + 1}] {d}"));
            txtChanges.SelectionStart = txtChanges.Text.Length;
            txtChanges.ScrollToCaret();
        }
        private void ClearUndo()
        {
            _undoStack.Clear();
            _redoStack.Clear();
            _changes.Clear();
            txtChanges.Text = string.Empty;
        }
        private string ColorToRGBString(Color color)
        {
            if (color != Color.Transparent)
            {
                return $"RGB:{color.R},{color.G},{color.B}";
            }
            else
            {
                return "Transparent";
            }
        }
        private ChangeItem GetChangeItem(TableLayoutPanel tblPanel, Label lbl, Color newColor, int col, int row)
        {
            var styleLabel = tblPanel.GetControlFromPosition(0, row) as Label;
            var styleType = col == 1 ? "Foreground" : "Background";
            var ci = new ChangeItem()
            {
                Style = $"{styleLabel.Text} ({styleType})",
                LabelPtr = lbl.Handle,
                OldLabelColor = lbl.BackColor,
                NewLabelColor = newColor
            };
            UpdateChanges($"{ci.Style} -> {ColorToRGBString(ci.OldLabelColor)} TO {ColorToRGBString(ci.NewLabelColor)}");
            return ci;
        }
        private void SetLabelColor(Label lbl, Color color, bool saveEnabled = true)
        {
            lbl.BackColor = color;
            SetLabelImage(lbl);
            SetColorLabelToolTip(lbl);
            FormHasChanged(saveEnabled);
        }
        private void ModifyButtonEvents(TableLayoutPanel tlp, bool addEvents = true)
        {
            foreach (var btn in tlp.Controls)
            {
                if (btn is Button colorButton && colorButton.Name.ToLower().StartsWith("button"))
                {
                    if (addEvents)
                    {
                        colorButton.Click += ColorButton_Click;
                    }
                    else
                    {
                        colorButton.Click -= ColorButton_Click;
                    }
                }
            }
        }
        private bool SaveChanges()
        {
            return _isDirty && MessageBox.Show("There are unsaved changes, would you like to save them first?", "Save Changes", MessageBoxButtons.YesNo) == DialogResult.Yes;
        }
        private void RefreshSample()
        {
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                var css = BuildCss();

                webBrowser1.DocumentText = Resources.SQLSample.Replace("/*<<CSS>>*/", css);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }
        private string BuildCss()
        {
            var sb = new StringBuilder();

            sb.Append("body {");
            sb.Append($"font-family:\"{_fontName}\";font-size:{_fontSize}pt;");
            if (lblPlainTextFG.BackColor != Color.Transparent) { sb.Append($"color:{ToCssHex(lblPlainTextFG.BackColor)};"); }
            if (lblPlainTextBG.BackColor != Color.Transparent) { sb.Append($"background-color:{ToCssHex(lblPlainTextBG.BackColor)};"); }
            sb.AppendLine("}");

            if (lblSelectedTextBG.BackColor != Color.Transparent) { sb.AppendLine($"::selection {{background-color:{ToCssHex(lblSelectedTextBG.BackColor)};}}"); }

            AddCssStyles(sb, tlpGeneral);
            AddCssStyles(sb, tlpSQL);
            AddCssStyles(sb, tlpXML);

            return sb.ToString();
        }
        private string BuildItems()
        {
            var sb = new StringBuilder();

            AddItemStyles(sb, tlpGeneral);
            AddItemStyles(sb, tlpSQL);
            AddItemStyles(sb, tlpXML);

            return sb.ToString();
        }
        private void AddItemStyles(StringBuilder sb, TableLayoutPanel tlp)
        {
            var spacer = new string('\t', 7);
            for (int i = 0; i <= tlp.RowCount; i++)
            {
                if (!(tlp.GetControlFromPosition(0, i) is Label lbl)) { continue; }

                var fgColor = tlp.GetControlFromPosition(1, i) as Label;
                var bgColor = tlp.GetControlFromPosition(3, i) as Label;

                if ((fgColor != null && fgColor.BackColor != Color.Transparent)
                    || (bgColor != null && bgColor.BackColor != Color.Transparent))
                {
                    sb.Append($"{spacer}<Item Name=\"{lbl.Text}\"");

                    if (fgColor != null) { sb.Append($" Foreground=\"{ToVssSettingsHex(fgColor.BackColor)}\""); }
                    if (bgColor != null) { sb.Append($" Background=\"{ToVssSettingsHex(bgColor.BackColor)}\""); }

                    sb.AppendLine(" BoldFont=\"No\"/>");

                    if (lbl.Text == "Plain Text")
                    {
                        sb.AppendLine($"{spacer}<Item Name=\"Refactoring Current Field\" Foreground=\"{ToVssSettingsHex(fgColor.BackColor)}\" Background=\"{ToVssSettingsHex(bgColor.BackColor)}\" BoldFont=\"No\"/>");
                    }
                }
            }
        }
        private void AddCssStyles(StringBuilder sb, TableLayoutPanel tlp)
        {
            for (int i = 0; i <= tlp.RowCount; i++)
            {
                if (!(tlp.GetControlFromPosition(0, i) is Label lbl)
                    || lbl.Text == "Plain Text"
                    || lbl.Text == "Selected Text") { continue; }

                var fgColor = tlp.GetControlFromPosition(1, i) as Label;
                var bgColor = tlp.GetControlFromPosition(3, i) as Label;

                if ((fgColor != null && fgColor.BackColor != Color.Transparent)
                    || (bgColor != null && bgColor.BackColor != Color.Transparent))
                {
                    var className = Regex.Replace(lbl.Text, @"[ \(\)/]", "");
                    sb.Append($".{className} {{");

                    if (fgColor != null) { sb.Append($" color:{ToCssHex(fgColor.BackColor)} !important;"); }
                    if (bgColor != null) { sb.Append($" background-color:{ToCssHex(bgColor.BackColor)} !important;"); }

                    sb.AppendLine("}");
                }
            }
        }
        private void LoadColors(XmlDocument doc, TableLayoutPanel tlp)
        {
            for (int i = 0; i <= tlp.RowCount; i++)
            {
                if (!(tlp.GetControlFromPosition(0, i) is Label lbl)) { continue; }

                var node = doc.SelectSingleNode($"//FontsAndColors//Categories/Category/Items/Item[@Name = '{lbl.Text}']");
                if (node == null) { continue; }

                var fgAttribute = node.Attributes["Foreground"];
                if (tlp.GetControlFromPosition(1, i) is Label fgColor && fgAttribute != null)
                {
                    fgColor.BackColor = GetColor(fgAttribute.Value);
                    SetColorLabelToolTip(fgColor);
                    SetLabelImage(fgColor);
                }

                var bgAttribute = node.Attributes["Background"];
                if (tlp.GetControlFromPosition(3, i) is Label bgColor && bgAttribute != null)
                {
                    bgColor.BackColor = GetColor(bgAttribute.Value);
                    SetColorLabelToolTip(bgColor);
                    SetLabelImage(bgColor);
                }
            }

        }
        private void SetupColorLabels(TableLayoutPanel tlp)
        {
            foreach (var ctl in tlp.Controls)
            {
                if (ctl is Label lbl && string.IsNullOrWhiteSpace(lbl.Text))
                {
                    lbl.Image = Resources.BlankImage;
                    lbl.ContextMenuStrip = cmsClearColor;
                }
            }
        }
        private void BlankLabels()
        {
            BlankLabels(tlpGeneral);
            BlankLabels(tlpSQL);
            BlankLabels(tlpXML);
        }
        private bool BlankLabels(TableLayoutPanel tlp, bool captureUndo = false)
        {
            var hasChanged = false;
            for (int i = 0; i <= tlp.RowCount; i++)
            {
                if (tlp.GetControlFromPosition(1, i) is Label fgColor)
                {
                    hasChanged = SetLabelBlank(tlp, fgColor, captureUndo, hasChanged);
                }
                if (tlp.GetControlFromPosition(3, i) is Label bgColor)
                {
                    hasChanged = SetLabelBlank(tlp, bgColor, captureUndo, hasChanged);
                }
            }
            return hasChanged;
        }
        private bool SetLabelBlank(TableLayoutPanel tlp, Label lbl, bool captureUndo, bool hasChanged)
        {
            if (lbl.BackColor == Color.Transparent) { return hasChanged; }
            if (!hasChanged && lbl.BackColor != Color.Transparent) { hasChanged = true; }
            if (captureUndo && hasChanged)
            {
                var row = tlp.GetRow(lbl);
                var col = tlp.GetColumn(lbl);

                var ci = GetChangeItem(tlp, lbl, Color.Transparent, col, row);
                _undoStack.Push(ci);
            }
            lbl.BackColor = Color.Transparent;
            lbl.Image = Resources.BlankImage;
            return hasChanged;
        }
        private void SetColorLabelToolTip(Label lbl)
        {
            string colorName = lbl.BackColor.IsKnownColor ? lbl.BackColor.ToKnownColor().ToString() : lbl.BackColor.Name;
            toolTip1.SetToolTip(lbl, $"Color: {colorName}, RGB: {lbl.BackColor.R}, {lbl.BackColor.G}, {lbl.BackColor.B}");
        }
        private Color GetColor(string value)
        {
            //for w/e reason, visual studio stores the rgb reversed as bgr. so reverse the values as they are extracted
            var rgb = Color.FromArgb(
                            Convert.ToInt32(value.Substring(8, 2), 16),
                            Convert.ToInt32(value.Substring(6, 2), 16),
                            Convert.ToInt32(value.Substring(4, 2), 16)
                        );
            return rgb;
        }
        private string ToVssSettingsHex(Color color)
        {
            //return the string in reverse order as VS expects it
            return $"0x00{color.B.ToString("X2")}{color.G.ToString("X2")}{color.R.ToString("X2")}";
        }
        private string ToCssHex(Color color)
        {
            return $"#{color.R.ToString("X2")}{color.G.ToString("X2")}{color.B.ToString("X2")}";
        }
        private void SetLabelImage(Label lbl)
        {
            if (lbl.BackColor == Color.Transparent)
            {
                lbl.Image = Resources.BlankImage;
            }
            else
            {
                lbl.Image = null;
            }
        }
        private void FormHasChanged(bool saveEnabled = true)
        {
            btnSave.Enabled = _isDirty = saveEnabled;
            RefreshSample();
        }
        private void SetFont(string fontName, string fontSize)
        {
            var fontIndex = cboFont.FindStringExact(fontName);
            if (fontIndex >= 0)
            {
                cboFont.SelectedIndex = fontIndex;
            }
            else
            {
                cboFont.Text = _defaultFontName;
            }

            var fontSizeIndex = cboFontSize.FindStringExact(fontSize);
            if (fontSizeIndex >= 0)
            {
                cboFontSize.SelectedIndex = fontSizeIndex;
            }
            else
            {
                cboFontSize.Text = _defaultFontSize;
            }
        }
        private void GetFontInfo(XmlDocument doc)
        {
            //scan the doc for a default font
            var node = doc.SelectSingleNode("/UserSettings/Category/Category/FontsAndColors/Categories/Category[@FontIsDefault='Yes']");

            //default not found, look for the plain text item, and get its category
            if (node == null)
            {
                node = doc.SelectSingleNode("/UserSettings/Category/Category/FontsAndColors/Categories/Category/Items/Item[@Name='Plain Text']/ancestor::Category[1]");
                //still null, just grab the first one
                if (node == null)
                {
                    node = doc.SelectSingleNode("/UserSettings/Category/Category/FontsAndColors/Categories/Category[1]");
                    if (node == null)
                    {
                        SetFont(_defaultFontName, _defaultFontSize);
                        return;
                    }
                }
            }
            _fontName = node.Attributes["FontName"].Value;
            _fontSize = node.Attributes["FontSize"].Value;
            SetFont(_fontName, _fontSize);
        }
        #endregion private methods
    }
}

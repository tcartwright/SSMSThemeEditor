using SSMSThemeEditor.Properties;
using System;
using System.Collections.Generic;
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

        #region events
        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
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
                    lbl.BackColor = colorDialog1.Color;
                    SetLabelImage(lbl);
                    SetColorLabelToolTip(lbl);
                    ColorsHaveChanged();
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

                    BlankColors();
                    LoadColors(doc, tlpGeneral);
                    LoadColors(doc, tlpSQL);
                    LoadColors(doc, tlpXML);

                    _saveFile = openFileDialog1.FileName;
                    _settingsInitialDirectory = Path.GetDirectoryName(openFileDialog1.FileName);
                    ColorsHaveChanged(false);

                    this.Text = $"{Resources.AppTitle} - {Path.GetFileName(openFileDialog1.FileName)}";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }
        private void btnNew_Click(object sender, EventArgs e)
        {
            if (SaveChanges()) { return; }

            this.Text = Resources.AppTitle;
            BlankColors();
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

                    var settings = Resources.VSSettingsTemplate.Replace("<!--<<ITEMS>>-->", items);

                    File.WriteAllText(saveFileDialog1.FileName, settings);

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
                lbl.BackColor = Color.Transparent;
                lbl.Image = Resources.BlankImage;
                ColorsHaveChanged();
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
                        hasChanged = BlankColors(tlpGeneral);
                    }
                    else if (tp == tabPageSQL)
                    {
                        hasChanged = BlankColors(tlpSQL);
                    }
                    else if (tp == tabPageXML)
                    {
                        hasChanged = BlankColors(tlpXML);
                    }
                    if (hasChanged)
                    {
                        ColorsHaveChanged();
                    }
                }
            }
        }

        #endregion events

        #region private methods
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

            sb.AppendLine($"body {{ color:{ToCssHex(lblPlainTextFG.BackColor)}; background-color:{ToCssHex(lblPlainTextBG.BackColor)}; }}");
            sb.AppendLine($"::selection {{ background-color:{ToCssHex(lblSelectedTextBG.BackColor)}; }}");

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

                var className = Regex.Replace(lbl.Text, @"[ \(\)/]", "");

                var fgColor = tlp.GetControlFromPosition(1, i) as Label;
                var bgColor = tlp.GetControlFromPosition(3, i) as Label;

                if ((fgColor != null && fgColor.BackColor != Color.Transparent)
                    || (bgColor != null && bgColor.BackColor != Color.Transparent))
                {
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
        private void BlankColors()
        {
            BlankColors(tlpGeneral);
            BlankColors(tlpSQL);
            BlankColors(tlpXML);
        }
        private bool BlankColors(TableLayoutPanel tlp)
        {
            var hasChanged = false;
            for (int i = 0; i <= tlp.RowCount; i++)
            {
                if (tlp.GetControlFromPosition(1, i) is Label fgColor)
                {
                    if (!hasChanged && fgColor.BackColor != Color.Transparent) { hasChanged = true; }
                    fgColor.BackColor = Color.Transparent;
                    fgColor.Image = Resources.BlankImage;
                }
                if (tlp.GetControlFromPosition(3, i) is Label bgColor)
                {
                    if (!hasChanged && bgColor.BackColor != Color.Transparent) { hasChanged = true; }
                    bgColor.BackColor = Color.Transparent;
                    bgColor.Image = Resources.BlankImage;
                }
            }
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
        private void ColorsHaveChanged(bool saveEnabled = true)
        {
            btnSave.Enabled = _isDirty = saveEnabled;
            RefreshSample();
        }
        #endregion private methods
    }
}

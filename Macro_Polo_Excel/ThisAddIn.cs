using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Drawing;
using Microsoft.Office.Core;
using System.Collections.Generic;

namespace Macro_Polo_Excel
{
    public partial class ThisAddIn
    {
        private UserControl UserControl1;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        private Label warningLabel;
        private int taskPaneHeight;
        private Ribbon1 ribbon;
        private bool isTaskPaneOpen = false;

        private Dictionary<Excel.Workbook, Microsoft.Office.Tools.CustomTaskPane> documentTaskPanes =
            new Dictionary<Excel.Workbook, Microsoft.Office.Tools.CustomTaskPane>();

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            ribbon = new Ribbon1();
            return ribbon;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            foreach (var taskPane in documentTaskPanes.Values)
            {
                taskPane.Dispose();
            }
            documentTaskPanes.Clear();
        }

        private void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            isTaskPaneOpen = myCustomTaskPane.Visible;
        }

        private float GetDpiScale()
        {
            using (Graphics graphics = Graphics.FromHwnd(IntPtr.Zero))
            {
                return graphics.DpiX / 96f; // 96 DPI is the default/standard DPI
            }
        }

        private int CalculateTaskPaneHeight()
        {
            float dpiScale = GetDpiScale();
            int baseHeight = 65; // Base height at 100% scaling
            return (int)(baseHeight * dpiScale);
        }

        private float CalculateFontSize()
        {
            float dpiScale = GetDpiScale();
            float baseSize = 12f; // Base font size at 100% scaling
            return baseSize * dpiScale;
        }

        private int AreMacrosEnabled()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Excel\Security"))
            {
                object value = key?.GetValue("VBAWarnings");

                if (value == null)
                {
                    using (RegistryKey key2 = Registry.CurrentUser.OpenSubKey(@"Software\Policies\Microsoft\Office\16.0\Excel\Security"))
                    {
                        object value2 = key2?.GetValue("VBAWarnings");
                        return value2 == null ? 0 : (int)value2;
                    }
                }
                return (int)value;
            }
        }

        public void CheckMacroStatus()
        {
            try
            {
                Excel.Workbook Wb = this.Application.ActiveWorkbook;

                if (documentTaskPanes.TryGetValue(Wb, out var existingTaskPane))
                {
                    if (existingTaskPane.Visible)
                    {
                        return;
                    }
                    existingTaskPane.Visible = true;
                    return;
                }

                string text;
                Color forecolor;
                Color backcolor;
                taskPaneHeight = CalculateTaskPaneHeight();
                float fontSize = CalculateFontSize();

                if (Wb.HasVBProject)
                {
                    if (AreMacrosEnabled() == 4)
                    {
                        if (!Wb.VBASigned)
                        {
                            text = "The macro in this file is not signed. You also do not have permission to run macros.";
                            forecolor = ColorTranslator.FromHtml("#FFFFFF");
                            backcolor = ColorTranslator.FromHtml("#205493");
                        }
                        else
                        {
                            text = "This file has a signed macro, but you do not have permission to run them.";
                            forecolor = ColorTranslator.FromHtml("#FFFFFF");
                            backcolor = ColorTranslator.FromHtml("#981b1e");
                        }
                    }
                    else
                    {
                        if (!Wb.VBASigned)
                        {
                            text = "This file contains a macro, which has not been digitally signed.";
                            forecolor = ColorTranslator.FromHtml("#212121");
                            backcolor = ColorTranslator.FromHtml("#F9C642");
                        }
                        else
                        {
                            text = "The macro in this document is digitally signed.";
                            forecolor = ColorTranslator.FromHtml("#FFFFFF");
                            backcolor = ColorTranslator.FromHtml("#225D2E");
                        }
                    }
                }
                else
                {
                    text = "There is no macro in this document.";
                    forecolor = ColorTranslator.FromHtml("#FFFFFF");
                    backcolor = ColorTranslator.FromHtml("#323A45");
                }

                warningLabel = new Label
                {
                    Text = text,
                    Font = new Font("Segoe UI", fontSize, FontStyle.Bold),
                    ForeColor = forecolor,
                    Location = new Point(5, 2),
                    AutoSize = false,
                    TextAlign = ContentAlignment.TopLeft,
                    Padding = new Padding(3),
                    Dock = DockStyle.Fill
                };

                UserControl1 = new UserControl();
                UserControl1.BackColor = backcolor;
                UserControl1.Controls.Add(warningLabel);
                UserControl1.Resize += UserControl1_Resize;

                myCustomTaskPane = this.CustomTaskPanes.Add(UserControl1, "Macro Status");
                myCustomTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionTop;
                myCustomTaskPane.Height = taskPaneHeight;
                myCustomTaskPane.Visible = true;
                isTaskPaneOpen = true;

                myCustomTaskPane.VisibleChanged += (sender, e) =>
                {
                    if (!myCustomTaskPane.Visible)
                    {
                        documentTaskPanes.Remove(Wb);
                    }
                };

                documentTaskPanes[Wb] = myCustomTaskPane;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error: {ex.Message}",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        private void UserControl1_Resize(object sender, EventArgs e)
        {
            if (warningLabel != null)
            {
                warningLabel.MaximumSize = new Size(UserControl1.Width - 20, 0);
                warningLabel.MinimumSize = new Size(0, taskPaneHeight - 10);
            }
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
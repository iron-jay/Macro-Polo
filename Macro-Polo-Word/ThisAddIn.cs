using System;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Drawing;
using Microsoft.Office.Core;
using System.Collections.Generic;

namespace Macro_Polo_Word
{
    public partial class ThisAddIn
    {
        private UserControl UserControl1;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        private Label warningLabel;
        private int taskPaneHeight;
        private Ribbon1 ribbon;
        private bool isTaskPaneOpen = false;
        
        private Dictionary<Word.Document, Microsoft.Office.Tools.CustomTaskPane> documentTaskPanes =
            new Dictionary<Word.Document, Microsoft.Office.Tools.CustomTaskPane>();

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
        private int CalculateHeight()
        {
            int screenheight = Screen.PrimaryScreen.Bounds.Height;

            if (screenheight >= 2000)
            {
                return 75;
            }
            else if (screenheight >= 1080)
            {
                return 65;
            }
            else
            {
                float scaleFactor = screenheight / 1080f;
                int scaledHeight = (int)(60 * scaleFactor);
                return Math.Max(scaledHeight, 40);
            }
        }

        private float CalculateFont()
        {
            if (taskPaneHeight >= 75)
            {
                return 16f;
            }
            else if (taskPaneHeight >= 60)
            {
                return 12f;
            }
            else
            {
                return 12f;
            }
        }

        private int AreMacrosEnabled()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Word\Security"))
            {
                object value = key.GetValue("VBAWarnings");
                // 1 = Enable all macros (not recommended)
                // 2 = Disable all with notification
                // 3 = Disable all except digitally signed macros
                // 4 = Disable all without notification

                if (value == null)
                {
                    using (RegistryKey key2 = Registry.CurrentUser.OpenSubKey(@"Software\Policies\Microsoft\Office\16.0\Word\Security"))
                    {
                        object value2 = key2.GetValue("VBAWarnings");
                        // 1 = Enable all macros (not recommended)
                        // 2 = Disable all with notification
                        // 3 = Disable all except digitally signed macros
                        // 4 = Disable all without notification
                        if (value2 == null)
                        {
                            return 0;
                        }
                        else
                        {
                            return (int)value2;
                        }
                    }
                }
                else
                {
                    return (int)value;
                }
            }
        }

        public void CheckMacroStatus()
        {
            try
            {
                Word.Document Doc = this.Application.ActiveDocument;

                // Check if task pane already exists for this document
                if (documentTaskPanes.TryGetValue(Doc, out var existingTaskPane))
                {
                    // If it exists and is visible, do nothing
                    if (existingTaskPane.Visible)
                    {
                        return;
                    }

                    // If it exists but is not visible, make it visible
                    existingTaskPane.Visible = true;
                    return;
                }
                string text;
                Color forecolor;
                Color backcolor;
                taskPaneHeight = CalculateHeight();
                float fontSzie = CalculateFont();

                if (Doc.HasVBProject)
                {
                    if (AreMacrosEnabled() == 4)
                    {
                        if (!Doc.VBASigned)
                        {
                            text = "The macro in this file is not signed. You also do not have permission to run macros.";
                            forecolor = (Color)((new ColorConverter()).ConvertFromString("#FFFFFF"));
                            backcolor = (Color)((new ColorConverter()).ConvertFromString("#205493"));
                        }
                        else
                        {
                            text = "This file has a signed macro, but you do not have permission to run them.";
                            forecolor = (Color)((new ColorConverter()).ConvertFromString("#FFFFFF"));
                            backcolor = (Color)((new ColorConverter()).ConvertFromString("#981b1e"));
                        }

                    }
                    else
                    {
                        if (!Doc.VBASigned)
                        {
                            text = "This file contains a macro, which has not been digitally signed.";
                            forecolor = (Color)((new ColorConverter()).ConvertFromString("#212121"));
                            backcolor = (Color)((new ColorConverter()).ConvertFromString("#F9C642"));
                        }

                        else
                        {

                            text = "The macro in this document is digitally signed.";
                            forecolor = (Color)((new ColorConverter()).ConvertFromString("#FFFFFF"));
                            backcolor = (Color)((new ColorConverter()).ConvertFromString("#225D2E"));
                        }
                    }
                }
                else
                {
                    text = "There is no macro in this document.";
                    forecolor = (Color)((new ColorConverter()).ConvertFromString("#FFFFFF"));
                    backcolor = (Color)((new ColorConverter()).ConvertFromString("#323A45"));
                }

                warningLabel = new Label
                {
                    Text = text,
                    Font = new System.Drawing.Font("Arial", fontSzie, System.Drawing.FontStyle.Bold),
                    ForeColor = forecolor,
                    Location = new System.Drawing.Point(5, 2),
                    AutoSize = true,
                    TextAlign = System.Drawing.ContentAlignment.MiddleCenter,

                };

                UserControl1 = new UserControl();
                UserControl1.BackColor = backcolor;
                UserControl1.Controls.Add(warningLabel);

                myCustomTaskPane = this.CustomTaskPanes.Add(UserControl1, "Macro Status");
                myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop;
                myCustomTaskPane.Height = taskPaneHeight;
                myCustomTaskPane.Visible = true;
                isTaskPaneOpen = true;

                // Add a custom event handler to remove the task pane when the document is closed
                myCustomTaskPane.VisibleChanged += (sender, e) =>
                {
                    if (!myCustomTaskPane.Visible)
                    {
                        documentTaskPanes.Remove(Doc);
                    }
                };

                // Store the task pane for this document
                documentTaskPanes[Doc] = myCustomTaskPane;

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

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

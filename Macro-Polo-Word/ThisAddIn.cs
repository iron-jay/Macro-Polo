﻿using System;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using Microsoft.Win32;

namespace Macro_Polo_Word
{
    public partial class ThisAddIn
    {
        private UserControl UserControl1;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        private Label warningLabel;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Subscribe to the WorkbookOpen event
            this.Application.DocumentOpen += Application_DocumentOpen;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Clean up event subscription
            this.Application.DocumentOpen -= Application_DocumentOpen;
        }

        private int AreMacrosEnabled()
        {

            if (DoesRegistryKeyExist(@"Software\Microsoft\Office\16.0\Word\Security"))
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Word\Security"))
                {
                    object value = key.GetValue("VBAWarnings");
                    // 1 = Enable all macros (not recommended)
                    // 2 = Disable all with notification
                    // 3 = Disable all except digitally signed macros
                    // 4 = Disable all without notification
                    return (int)value;
                }
            }

            else if (DoesRegistryKeyExist(@"Software\Policies\Microsoft\Office\16.0\Word\Security"))
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Policies\Microsoft\Office\16.0\Word\Security"))
                {
                    object value = key.GetValue("VBAWarnings");
                    // 1 = Enable all macros (not recommended)
                    // 2 = Disable all with notification
                    // 3 = Disable all except digitally signed macros
                    // 4 = Disable all without notification
                    return (int)value;
                }
            }
            else
            {
                return 0;
            }
        }

        private bool DoesRegistryKeyExist(string path)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(path))
            {
                return key != null;
            }
        }


        private void Application_DocumentOpen(Word.Document Doc)
        {
            try
            {
                if (Doc.HasVBProject)
                {

                    if (AreMacrosEnabled() == 4)
                    {
                        warningLabel = new Label
                        {
                            Text = "This file has a macro, but you do not have permission to run them.",
                            Font = new System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Bold),
                            ForeColor = System.Drawing.Color.Black,
                            Location = new System.Drawing.Point(5, 2),
                            AutoSize = true,
                            TextAlign = System.Drawing.ContentAlignment.MiddleCenter,

                        };

                        UserControl1 = new UserControl();
                        UserControl1.BackColor = System.Drawing.Color.Gray;
                        UserControl1.Controls.Add(warningLabel);
                
                        myCustomTaskPane = this.CustomTaskPanes.Add(UserControl1, "Macro Status");

                        myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop;
                        myCustomTaskPane.Height = 75;
                        myCustomTaskPane.Visible = true;
                    }
                    else
                    {
                        if (!Doc.VBASigned)
                        {
                            warningLabel = new Label
                            {
                                Text = "This file contains a macro, which has not been digitally signed.",
                                Font = new System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Bold),
                                ForeColor = System.Drawing.Color.Red,
                                Location = new System.Drawing.Point(5, 2),
                                AutoSize = true,
                                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,

                            };

                            UserControl1 = new UserControl();
                            UserControl1.BackColor = System.Drawing.Color.Orange;
                            UserControl1.Controls.Add(warningLabel);

                            myCustomTaskPane = this.CustomTaskPanes.Add(UserControl1, "Macro Status");

                            myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop;
                            myCustomTaskPane.Height = 75;
                            myCustomTaskPane.Visible = true;

                        }

                        else
                        {
                            warningLabel = new Label
                            {
                                Text = "The macro in this document is digitally signed.",
                                Font = new System.Drawing.Font("Arial", 12, System.Drawing.FontStyle.Bold),
                                ForeColor = System.Drawing.Color.LightBlue,
                                Location = new System.Drawing.Point(5, 2),
                                AutoSize = true,
                                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,

                            };

                            UserControl1 = new UserControl();
                            UserControl1.BackColor = System.Drawing.Color.Green;
                            UserControl1.Controls.Add(warningLabel);

                            myCustomTaskPane = this.CustomTaskPanes.Add(UserControl1, "Macro Status");

                            myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionTop;
                            myCustomTaskPane.Height = 75;
                            myCustomTaskPane.Visible = true;

                        }
                    }
                }
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

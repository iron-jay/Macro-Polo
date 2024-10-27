﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;

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



        private void Application_DocumentOpen(Word.Document Doc)
        {
            try
            {
                if (Doc.HasVBProject)
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
                        myCustomTaskPane.Height = 65;
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
                        myCustomTaskPane.Height = 65;
                        myCustomTaskPane.Visible = true;

                    }
                }

                else
                {
                    MessageBox.Show(
                        $"The workbook '{Doc.Name}' has no macro.",
                        "No Macro Found",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                        );
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

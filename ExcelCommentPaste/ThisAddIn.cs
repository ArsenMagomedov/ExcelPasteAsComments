using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace ExcelCommentPaste
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Office.CommandBar cellbar = this.Application.CommandBars["Cell"];
            Office.CommandBarButton button = (Office.CommandBarButton)cellbar.FindControl(Office.MsoControlType.msoControlButton, 0, "MYRIGHTCLICKMENU", Missing.Value, Missing.Value);
            if (button == null)
            {
                // add the button
                button = (Office.CommandBarButton)cellbar.Controls.Add(Office.MsoControlType.msoControlButton, Missing.Value, Missing.Value, cellbar.Controls.Count, true);
                button.Caption = "Refresh";
                button.BeginGroup = true;
                button.Tag = "MYRIGHTCLICKMENU";
                button.Click += new Office._CommandBarButtonEvents_ClickEventHandler(MyButton_Click);
            }
        }

        private void MyButton_Click(Office.CommandBarButton cmdBarbutton, ref bool cancel)
        {
            System.Windows.Forms.MessageBox.Show("MyButton was Clicked", "MyCOMAddin");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace CommentsPaste
{
    public partial class ThisAddIn
    {

        private Office.CommandBarButton _button;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var cellbar = Application.CommandBars["Cell"];
            _button = (Office.CommandBarButton)cellbar.FindControl(Office.MsoControlType.msoControlButton, 0, "MYRIGHTCLICKMENU", Missing.Value, Missing.Value);

            if (_button != null) return;

            // add the button
            _button = (Office.CommandBarButton)cellbar.Controls.Add(Office.MsoControlType.msoControlButton, Missing.Value, Missing.Value, cellbar.Controls.Count, true);
            _button.Caption = "Paste as comments";
            _button.BeginGroup = true;
            _button.Tag = "Paste as comments";
            _button.Click += MyButton_Click; ;
        }

        private void MyButton_Click(Office.CommandBarButton cmdBarbutton, ref bool cancel)
        {
            try
            {

                var app = Application;

                var clipBoardText = System.Windows.Forms.Clipboard.GetText().Replace("\r\n", "\n");
                var clipBoard = clipBoardText.Split('\n').Select(txt => txt.Split('\t'));

                var selected = Application.Selection as IEnumerable;

                if (Application.Selection is Excel.Range && selected != null)
                {
                    dynamic cell = First(selected);

                    var rowNum = cell.Row;
                    var colNum = cell.Column;

                    var firstCol = colNum;

                    foreach (var row in clipBoard)
                    {
                        foreach (var val in row)
                        {
                            if (String.IsNullOrEmpty(val))
                                continue;

                            cell = app.Cells[rowNum, colNum];

                            if (cell.Comment != null)
                            {
                                cell.Comment.Delete();
                            }

                            cell.AddComment(val);

                            colNum++;
                        }

                        colNum = firstCol;
                        rowNum++;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, @"Paste As Comment");
                throw;
            }
        }

        private dynamic First(IEnumerable collection)
        {
            var enumer = collection.GetEnumerator();
            return enumer.MoveNext() ? enumer.Current : null;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            Application.CommandBars["Cell"].Reset();
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

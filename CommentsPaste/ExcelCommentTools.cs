using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace CommentsPaste
{
	public partial class ExcelCommentTools
	{

		private Office.CommandBarButton _pasteAsCommentButton;

		private Office.CommandBarButton _copyCommentsButton;

		private void ExcelCommentTools_Startup(object sender, EventArgs e)
		{
			AddButtonToCellContextMenu("Paste as comments", PasteAsCommentButton_Click);

			AddButtonToCellContextMenu("Copy comments", CopyCommentsButton_Click);
		}

		private void AddButtonToCellContextMenu(string caption, Office._CommandBarButtonEvents_ClickEventHandler handler)
		{
			var cellbar = Application.CommandBars["Cell"];

			_copyCommentsButton = (Office.CommandBarButton)cellbar.Controls.Add(Office.MsoControlType.msoControlButton, Missing.Value, Missing.Value, cellbar.Controls.Count, true);
			_copyCommentsButton.Caption = caption;
			_copyCommentsButton.BeginGroup = true;
			_copyCommentsButton.Tag = caption;
			_copyCommentsButton.Click += handler;
		}

		/// <summary>
		/// Paste a text from cliboard to cells comments, starting with the first selected.
		/// Text should be \t and \n separated (for cols and rows accordingly)
		/// </summary>
		private void PasteAsCommentButton_Click(Office.CommandBarButton cmdBarbutton, ref bool cancel)
		{
			try
			{
				var app = Application;

				var clipBoardText = System.Windows.Forms.Clipboard.GetText().Replace("\r\n", "\n");
				var clipBoard = clipBoardText.Split('\n').Select(txt => txt.Split('\t'));

				var selected = Application.Selection as IEnumerable;

				if (Application.Selection is Excel.Range && selected != null)
				{
					dynamic cell = selected.Cast<dynamic>().First();

					var rowNum = cell.Row;
					var colNum = cell.Column;

					var firstCol = colNum;

					foreach (var row in clipBoard)
					{
						foreach (var val in row)
						{
							cell = app.Cells[rowNum, colNum];

							if (cell.Comment != null)
							{
								cell.Comment.Delete();
							}

							if (String.IsNullOrEmpty(val))
							{
								colNum++;

								continue;
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

		/// <summary>
		/// Copies comments from selected area to clipboard.
		/// Comments will be \t and \n separated (for cols and rows accordingly)
		/// </summary>
		private void CopyCommentsButton_Click(Office.CommandBarButton cmdBarbutton, ref bool cancel)
		{
			try
			{
				dynamic selected = Application.Selection as IEnumerable;

				if (Application.Selection is Excel.Range && selected != null)
				{
					var copiedString = "";

					var lastCellsRow = -1;

					foreach (dynamic cell in (IEnumerable) selected.Cells)
					{
						//New Row
						if (lastCellsRow != cell.Row)
						{
							copiedString += "\n";
							lastCellsRow = cell.Row;
						}
						else //New Column
						{
							copiedString += "\t";
						}

						var comment = cell.Comment;

						if (comment != null)
						{
							copiedString += comment.Text;
						}
					}

					copiedString = copiedString.Trim('\t', '\n');

					System.Windows.Forms.Clipboard.SetText(copiedString);
				}
			}
			catch (Exception ex)
			{
				System.Windows.Forms.MessageBox.Show(ex.Message, @"Paste As Comment");
				throw;
			}
		}

		private void ExcelCommentTools_Shutdown(object sender, EventArgs e)
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
			this.Startup += new System.EventHandler(ExcelCommentTools_Startup);
			this.Shutdown += new System.EventHandler(ExcelCommentTools_Shutdown);
		}

		#endregion
	}
}

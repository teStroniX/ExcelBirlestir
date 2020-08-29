using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelBirlestir
{
	public partial class Form1 : Form
	{
		bool go = true;
		string name = "";
		string[] files;
		public Form1()
		{
			InitializeComponent();
		}

		public void Output()
		{
			int x = 1, y = Convert.ToInt32(numericUpDown2.Value) - 1;
			int[] toplamlar = new int[Convert.ToInt32(numericUpDown1.Value)];

			for (int i = 0; i < numericUpDown1.Value - 1; i++)
			{
				toplamlar[i] = 0;
			}

			object Missing = System.Reflection.Missing.Value;
			Excel.Application ExcelUygulama = new Excel.Application();
			Excel.Workbook ExcelProje = ExcelUygulama.Workbooks.Add(files[0]);
			Excel.Worksheet ExcelSayfa = (Excel.Worksheet)ExcelProje.Worksheets.get_Item(1);
			Excel.Range ExcelRange = ExcelSayfa.UsedRange;
			ExcelSayfa = (Excel.Worksheet)ExcelUygulama.ActiveSheet;
			ExcelUygulama.Visible = false;

			for (int i = 0; i < files.Length; i++)
			{
				Excel.Application ExcelUygulama2 = new Excel.Application();
				Excel.Workbook ExcelProje2 = ExcelUygulama.Workbooks.Add(files[i]);
				Excel.Worksheet ExcelSayfa2 = (Excel.Worksheet)ExcelProje.Worksheets.get_Item(1);
				Excel.Range ExcelRange2 = ExcelSayfa.UsedRange;
				ExcelSayfa2 = (Excel.Worksheet)ExcelUygulama.Sheets[1];

				for (int j = 3; go; j++)
				{
					ExcelRange2 = (Excel.Range)ExcelSayfa2.Cells[j, 1];
					if (ExcelRange2.Value2 == "TOPLAM")
						break;
					y++;
					for (int k = 1; k <= numericUpDown1.Value; k++)
					{
						ExcelRange2 = (Excel.Range)ExcelSayfa2.Cells[j, k];
						ExcelRange = (Excel.Range)ExcelSayfa.Cells[y, k];
						ExcelRange.ClearFormats();
						if (k == 1)
						{
							ExcelRange.Interior.Color = Color.FromArgb(123, 123, 123);
							ExcelRange.RowHeight = 40;
							ExcelRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
							ExcelRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
						}
						else
						{
							try { toplamlar[k - 2] += ExcelRange2.Value2 == null ? 0 : ExcelRange2.Value2; } catch { }
						}
						ExcelRange.Borders.Color = Color.Black;
						ExcelRange.Value = ExcelRange2.Value2;
					}
				}

				ExcelRange2 = null;
				ExcelSayfa2 = null;
				ExcelProje2.Close(false);
				ExcelUygulama2.Quit();
				ExcelProje2 = null;
				ExcelUygulama2 = null;
			}
			y++;
			for (int i = 1; i <= numericUpDown1.Value; i++)
			{
				ExcelRange = (Excel.Range)ExcelSayfa.Cells[y, i];
				if (i == 1)
				{
					ExcelRange.Interior.Color = Color.FromArgb(123, 123, 123);
					ExcelRange.RowHeight = 40;
					ExcelRange.Value2 = "TOPLAM";
					ExcelRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
					ExcelRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
				}
				else
				{
					ExcelRange.Value2 = toplamlar[i - 2];
				}
				ExcelRange.Borders.Color = Color.Black;
			}

			//ExcelSayfa.Columns.AutoFit();
			//ExcelSayfa.Rows.AutoFit();

			ExcelProje.SaveAs(name, Excel.XlFileFormat.xlWorkbookDefault, Missing, Missing, false, Missing, Excel.XlSaveAsAccessMode.xlNoChange);
			ExcelProje.Close(true, Missing, Missing);
			ExcelUygulama.Quit();
			MessageBox.Show("Tamamlandı!");
			progressBar1.Invoke(new MethodInvoker(() =>
			{
				progressBar1.Style = ProgressBarStyle.Blocks;
			}));

			foreach (var process in Process.GetProcessesByName("EXCEL"))
			{
				process.Kill();
			}
		}

		private void button1_Click(object sender, EventArgs e)
		{
			OpenFileDialog dialog = new OpenFileDialog();
			dialog.Multiselect = true;
			var result = dialog.ShowDialog();

			if(result == DialogResult.OK)
			{
				for(int i=0; i< dialog.FileNames.Length; i++)
					richTextBox1.AppendText(dialog.FileNames[i] + (i==dialog.FileNames.Length - 1 ? "" : "\n"));
			}
		}

		private void button2_Click(object sender, EventArgs e)
		{
			SaveFileDialog dialog = new SaveFileDialog();
			dialog.Filter = "Excel Dosyası | .xlsx";

			var result = dialog.ShowDialog();

			if (result == DialogResult.OK)
			{
				name = dialog.FileName;
				files = richTextBox1.Lines;
				progressBar1.Style = ProgressBarStyle.Marquee;
				BackgroundWorker worker = new BackgroundWorker();
				worker.DoWork += DoWork;
				worker.RunWorkerAsync();
			}
		}
		public void DoWork(object obj, EventArgs e)
		{
			Output();
		}
	}
}

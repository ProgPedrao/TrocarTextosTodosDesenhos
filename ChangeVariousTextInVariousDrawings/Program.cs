using System;
using System.IO;
using System.Windows.Forms;
using Autodesk.AutoCAD.Interop;
using System.Runtime.InteropServices;
using Autodesk.AutoCAD.Interop.Common;

namespace ChangeVariousTextInVariousDrawings
{
	class Program
	{
		private static string[] arquivos, numeros;

		[STAThread]
		public static void Main(string[] args)
		{
			SelectFiles();
			SelectFilesTXT();
			
			if (arquivos.Length == 0)
			{
				Environment.Exit(0);
			}

            AcadApplication acApp = null;

            try
            {
                acApp = Marshal.GetActiveObject("AutoCAD.Application") as AcadApplication;
                acApp.Visible = true;
            }
            catch
            {
                MessageBox.Show("Não foi possível abrir o AutoCAD");
            }
			
        	foreach (var desenho in arquivos)
        	{
                AcadDocument doc;

                try
                {
                    doc = acApp.Documents.Open(desenho, false, string.Empty);
                    AcadModelSpace modelSpace = doc.ModelSpace;
                }
                catch (Exception)
                {
                    MessageBox.Show("Não foi possível abrir o desenho: ", desenho);
                    break;
                }

                AcadSelectionSet selset = null;
                selset = doc.SelectionSets.Add("txt");
                short[] ftype = new short[1];
                ftype[0] = 0;
                object[] fdata = new object[1];
                fdata[0] = "TEXT";
                selset.Select(AcSelect.acSelectionSetAll, null, null, ftype, fdata);

                foreach (IAcadText txt in selset)
                {
                    foreach (var item in numeros)
                    {
                        if (txt.TextString == item.Split(';')[0])
                        {
                            txt.TextString = item.Split(';')[1];
                        }
                    }
                }
                
                selset = doc.SelectionSets.Add("Mtxt");
                ftype[0] = 0;
                fdata[0] = "MTEXT";
                selset.Select(AcSelect.acSelectionSetAll, null, null, ftype, fdata);

                foreach (IAcadMText mtxt in selset)
                {
                    foreach (var item in numeros)
                    {
                        if (mtxt.TextString == item.Split(';')[0])
                        {
                            mtxt.TextString = item.Split(';')[1];
                        }
                    }
                }

                selset.Delete();
				acApp.ZoomExtents();
				doc.Save();
				doc.Close();
                acApp.Visible = true;
        	}
		}
		
        static void SelectFiles()
        {
            // Displays an OpenFileDialog so the user can select a Cursor.
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "Drawing Files|*.dwg";
            openFileDialog2.Title = "Selecione os arquivos DWG";
            openFileDialog2.RestoreDirectory = true;
            openFileDialog2.Multiselect = true;

            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                // Guarda a lista dos caminhos completos dos arquivos selecionados
                arquivos = openFileDialog2.FileNames;
            }
            else 
            {
				Console.WriteLine("Arquivos não selecionados!");
				Console.ReadKey();
            }
        }

        static void SelectFilesTXT()
        {
            // Displays an OpenFileDialog so the user can select a Cursor.
            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "TXT Files|*.txt";
            openFileDialog2.Title = "Selecione o arquivo TXT";
            openFileDialog2.RestoreDirectory = true;
            openFileDialog2.Multiselect = false;

            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                // Guarda a lista dos caminhos completos dos arquivos selecionados
                numeros = File.ReadAllLines(openFileDialog2.FileName);
            }
            else 
            {
				Console.WriteLine("Arquivo não selecionado!");
				Console.ReadKey();
            }
        }
	}
}
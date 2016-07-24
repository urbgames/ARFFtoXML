using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.IO;
using System.Text.RegularExpressions;

namespace LeituraExcel
{
    class Program
    {

        Application oXML;
        _Workbook oWB;
        _Worksheet oSheet;
        Range oRng;

        public void loadConfig()
        {

            oXML = new Application();

            oWB = (_Workbook)(oXML.Workbooks.Add(Missing.Value));
            oSheet = (_Worksheet)oWB.ActiveSheet;

        }

        static void Main(string[] args)
        {
            Program CreateExcelFile = new Program();
            CreateExcelFile.loadConfig();
            CreateExcelFile.lerArquivo();
            CreateExcelFile.showSheet();
        }

        private void showSheet()
        {
            oXML.Visible = true;
        }

        private void carregarPlanilhaTeste()
        {

            oSheet.Cells[1, 1] = "First Name";
            oSheet.Cells[1, 2] = "Last Name";
            oSheet.Cells[1, 3] = "Full Name";
            oSheet.Cells[1, 4] = "Document";
            oSheet.Cells[1, 5] = "Phone";

            oSheet.get_Range("A1", "E1").Font.Bold = true;
            oSheet.get_Range("A2", "E6").HorizontalAlignment = XlVAlign.xlVAlignCenter;

            string[,] saNames = new string[5, 2];

            saNames[0, 0] = "John";
            saNames[0, 1] = "Smith";
            saNames[1, 0] = "Tom";
            saNames[1, 1] = "Brown";
            saNames[2, 0] = "Sue";
            saNames[2, 1] = "Thomas";
            saNames[3, 0] = "Jane";
            saNames[3, 1] = "Jones";
            saNames[4, 0] = "Adam";
            saNames[4, 1] = "Johnson";

            oSheet.get_Range("A2", "B6").Value2 = saNames;

            oRng = oSheet.get_Range("C2", "C6");
            oRng.Formula = "=A2 & \" \" & B2";

            oRng = oSheet.get_Range("D2", "D6");
            oRng.Formula = "=RAND()*100000";
            oRng.NumberFormat = "$0.00";

            oRng = oSheet.get_Range("A1", "D1");
            oRng.EntireColumn.AutoFit();

        }

        private void addCell(int[] position, string text)
        {
            oSheet.Cells[position[0], position[1]] = text;
        }

        public void lerArquivo()
        {
            string[] FileLines = System.IO.File.ReadAllLines(@"C:\Users\Urbgames\Documents\_KEYSTROKE\_BASE 01\Resultado.arff");
            int count = 1, lineCount = 1, collumnCount = 1;
            foreach (string line in FileLines)
            {
                if (line.StartsWith("@attribute"))
                {
                    string[] attribute = line.Split(' ');
                    int[] cell = { lineCount, collumnCount };
                    addCell(cell, attribute[1]);
                    collumnCount++;
                }
                if (line.StartsWith("@data"))
                    lineCount++;
                if (lineCount > 1 && !line.Equals(""))
                {
                    string lineTemp = line;

                    string dataRegex = "";
                    if (line.StartsWith("'"))
                    {

                        Regex rgx = new Regex("\'([^\']*)\'");
                        Match match = rgx.Match(line);
                        dataRegex = match.Value;
                        lineTemp = rgx.Replace(line, "|*|");

                    }

                    string[] dataLine = lineTemp.Split(',');
                    collumnCount = 1;
                    foreach (string data in dataLine)
                    {
                        string dataTemp = data;
                        if (dataTemp.Equals("|*|"))
                            dataTemp = dataRegex;

                        int[] cell = { lineCount, collumnCount };
                        addCell(cell, dataTemp);
                        collumnCount++;
                    }
                    lineCount++;

                }

                Console.WriteLine(count + "/" + FileLines.Length);
                count++;
            }


        }


    }
}

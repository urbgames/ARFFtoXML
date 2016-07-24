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
            CreateExcelFile.loadFiles();
            CreateExcelFile.showSheet();
        }

        private void showSheet()
        {
            oXML.Visible = true;
        }

        private void addCell(int[] position, string text)
        {
            oSheet.Cells[position[0], position[1]] = text;
        }

        public void loadFiles()
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

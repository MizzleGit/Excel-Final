using OfficeOpenXml;
using System;
using System.Diagnostics;

namespace ExcelONE
{
    public partial class ExcelONEFinal : Form
    {

        // Variables
        // String arrays
        string[]? filePaths = null;
        //

        // Strings
        string folderPath = "";
        string fileName = "";
        //

        // Booleans
        bool filesOpened = false;
        //
        //

        public ExcelONEFinal()
        {
            InitializeComponent();
        }

        //
        //
        //
        //
        //

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            openExcel = new OpenFileDialog();
            openExcel.Filter = "Fichiers Excel (*.xlsx) | *xlsx";
            openExcel.Multiselect = true;

            if (openExcel.ShowDialog() == DialogResult.OK)
            {
                filePaths = openExcel.FileNames;
                filesOpened = true;
                Array.Sort(filePaths);
                try { folderPath = Path.GetDirectoryName(filePaths[0]); }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur! Selectionner les fichiers!\n\n" + ex, "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    filesOpened = false;
                }
            }
        } // private void btnBrowse_Click

        //
        //
        //

        private void btnModify_Click(object sender, EventArgs e)
        {
            if (filesOpened)
            {
                lblDebug.Text = "";
                // Local variables
                ExcelPackage[] mainPkgs = new ExcelPackage[filePaths.Count()];
                ExcelWorksheet[] mainWss = new ExcelWorksheet[filePaths.Count()];
                string fileName = "";
                string fileYear = "";
                string fileNature = "";
                string fileNatureABBR = "";
                string fileSheet = "";
                List<int> rowsToDelete = new List<int>();
                int lastRow;
                // Indexes
                int[] emptyCellIndex = new int[2] { -1, -1 };
                int[] drIndex = new int[2];
                int[] anneeIndex = new int[2];
                int[] natureIndex = new int[2];
                int[] concatIndex = new int[2];
                int[] ccIndex = new int[2];
                int[] tcIndex = new int[2];
                int[] gpeIndex = new int[2];
                int[] montantEchuIndex = new int[2];
                int[] montantRegleIndex = new int[2];

                int iterations = 0;
                //


                // Progress bar code

                pbarMain.Value = 0;
                int pbarInt = 100 / filePaths.Count();
                pbarInt++;

                // Looping through each file

                foreach (string path in filePaths)
                {
                    fileName = Path.GetFileNameWithoutExtension(path);
                    fileYear = getYearAndNature(fileName)[0];
                    fileNatureABBR = getYearAndNature(fileName)[1];
                    if (fileNatureABBR == "EN")
                    {
                        fileNature = "Energie";
                    }
                    if (fileNatureABBR == "TR")
                    {
                        fileNature = "Travaux";
                    }

                    // Package
                    mainPkgs[iterations] = new ExcelPackage(path);
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    // Worksheet
                    mainWss[iterations] = mainPkgs[iterations].Workbook.Worksheets[0];
                    fileSheet = mainWss[iterations].Name;

                    // Indexes
                    tcIndex = findCell(mainPkgs[iterations], fileSheet, "Type client");
                    ccIndex = findCell(mainPkgs[iterations], fileSheet, "Classe de compte");
                    gpeIndex = findCell(mainPkgs[iterations], fileSheet, "GpeStrReg");
                    montantEchuIndex = findCell(mainPkgs[iterations], fileSheet, "Montant échu");
                    montantRegleIndex = findCell(mainPkgs[iterations], fileSheet, "Montant réglé");


                    // Deleting empty cc rows (rows that represent totals etc...)
                    lastRow = mainWss[iterations].Dimension.End.Row;
                    for (int row = ccIndex[0] + 1; row <= lastRow; row++)
                    {
                        if (mainWss[iterations].Cells[row, ccIndex[1]].Value == null || string.IsNullOrEmpty(mainWss[iterations].Cells[row, ccIndex[1]].Value.ToString()))
                        {
                            rowsToDelete.Add(row);
                        }
                    }
                    foreach (int row in rowsToDelete.OrderByDescending(r => r))
                    {
                        mainWss[iterations].DeleteRow(row);
                    }
                    lastRow = mainWss[iterations].Dimension.End.Row;

                    // Adding the 3 cells Annee Nature Concat
                    if (findCell(mainPkgs[iterations], fileSheet, "Année").SequenceEqual(emptyCellIndex))
                    {
                        addColumnToTheRight(mainPkgs[iterations], fileSheet, "Année");
                    }
                    if (findCell(mainPkgs[iterations], fileSheet, "Nature").SequenceEqual(emptyCellIndex))
                    {
                        addColumnToTheRight(mainPkgs[iterations], fileSheet, "Nature");
                    }
                    if (findCell(mainPkgs[iterations], fileSheet, "Concat").SequenceEqual(emptyCellIndex))
                    {
                        addColumnToTheRight(mainPkgs[iterations], fileSheet, "Concat");
                    }
                    anneeIndex = findCell(mainPkgs[iterations], fileSheet, "Année");
                    natureIndex = findCell(mainPkgs[iterations], fileSheet, "Nature");
                    concatIndex = findCell(mainPkgs[iterations], fileSheet, "Concat");

                    // Filling Annee and Nature
                    mainWss[iterations].Cells[anneeIndex[0] + 1, anneeIndex[1], lastRow, anneeIndex[1]].Value = fileYear;
                    mainWss[iterations].Cells[natureIndex[0] + 1, natureIndex[1], lastRow, natureIndex[1]].Value = fileNature;

                    rowsToDelete.Clear();
                    mainPkgs[iterations].Save();
                    iterations++;

                    if (pbarMain.Value + pbarInt > 100) { pbarMain.Value = 100;}
                    else { pbarMain.Value += pbarInt; }
                }
            }
        } //private void btnModify_Click(object sender, EventArgs e)



        // Global methods


        // getYearAndNature
        public string[] getYearAndNature(string fnFileName)
        {
            string fnYear;
            string fnNature;
            string[] fnFileSplit = fnFileName.Split(" ");
            if (fnFileSplit[0].All(char.IsDigit))
            {
                fnYear = fnFileSplit[0];
                fnNature = fnFileSplit[1].ToUpper();
            }
            else
            {
                fnYear = fnFileSplit[1];
                fnNature = fnFileSplit[0].ToUpper();
            }
            return new string[] { fnYear, fnNature };
        }
        // getYearAndNature

        // FindCell
        public int[] findCell(ExcelPackage fnExcelPackage, string fnExcelWorksheet, string fnValueNeeded)
        {
            int columnIndex = -1;
            int rowIndex = -1;
            bool valueFound = false;
            int numberOfRows;
            int numberOfColumns;
            ExcelWorksheet fnWorksheet = fnExcelPackage.Workbook.Worksheets[fnExcelWorksheet];
            numberOfColumns = fnWorksheet.Dimension.End.Column;
            numberOfRows = fnWorksheet.Dimension.End.Row;
            for (int row = 1; row <= numberOfRows; row++)
            {
                for (int column = 1; column <= numberOfColumns; column++)
                {
                    var cellValue = fnWorksheet.Cells[row, column].Value?.ToString();
                    if (cellValue == fnValueNeeded)
                    {
                        rowIndex = row;
                        columnIndex = column;
                        valueFound = true;
                        break;
                    }
                }
                if (valueFound)
                {
                    break;
                }
            }
            return new int[2] { rowIndex, columnIndex };
        }
        // FindCell

        // addColumnToTheRight
        public void addColumnToTheRight(ExcelPackage fnExcelPackage, string fnExcelWorksheet, string fnColumnName)
        {
            ExcelWorksheet fnWorksheet = fnExcelPackage.Workbook.Worksheets[fnExcelWorksheet];
            int fnFirstRow = fnWorksheet.Dimension.Start.Row;
            int fnFirstCol = fnWorksheet.Dimension.Start.Column;
            int fnLastRow = fnWorksheet.Dimension.End.Row;
            int fnLastCol = fnWorksheet.Dimension.End.Column;
            ExcelRange fnTable = fnWorksheet.Cells[fnFirstRow, fnFirstCol, fnLastRow, fnLastCol];
            ExcelRange fnNewCol = fnWorksheet.Cells[fnFirstRow, fnLastCol + 1, fnLastRow, fnLastCol + 1];
            fnTable.CopyStyles(fnNewCol);
            int fnNewColFirstRow = fnNewCol.Start.Row;
            int fnNewColFirstColumn = fnNewCol.Start.Column;
            fnWorksheet.Cells[fnNewColFirstRow, fnNewColFirstColumn].Value = fnColumnName;
            fnNewCol.AutoFitColumns();
            fnExcelPackage.Save();
        }
        // addColumnToTheRight



    } // public partial class ExcelONEFinal : Form
} // namespace ExcelONE
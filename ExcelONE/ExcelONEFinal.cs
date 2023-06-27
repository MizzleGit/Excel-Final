using OfficeOpenXml;
using OfficeOpenXml.Style;
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

        // Excel variables
        List<ExcelPackage> mainPkgs = new List<ExcelPackage>();
        List<ExcelWorksheet> mainWss = new List<ExcelWorksheet>();
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
            mainPkgs.Clear();
            mainWss.Clear();
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
                // Local variables
                string fileName = "";
                string fileYear = "";
                string fileNature = "";
                string fileNatureABBR = "";
                string fileSheet = "";
                List<int> rowsToDelete = new List<int>();
                int lastRow;
                int iterations = 0;
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
                //

                // Setting Packages and Worksheets using lists for dynamic size management
                try
                {
                    foreach (string path in filePaths)
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        mainPkgs.Add(new ExcelPackage(path));
                    }
                    for (int i = 0; i < mainPkgs.Count(); i++)
                    {
                        mainWss.Add(mainPkgs[i].Workbook.Worksheets[0]);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception rencontrée!\n\n" + ex, "Erreur!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


                // Progress bar code

                pbarMain.Value = 0;
                int pbarInt = 100 / filePaths.Count();
                pbarInt++;

                // Looping through each file
                try
                {
                    foreach (string path in filePaths)
                    {
                        // Retrieving basic file information
                        fileName = Path.GetFileNameWithoutExtension(path);
                        fileYear = getYearAndNature(fileName)[0];
                        fileNatureABBR = getYearAndNature(fileName)[1];

                        // Finding file nature (Energie/Travaux)
                        if (fileNatureABBR == "EN")
                        {
                            fileNature = "Energie";
                        }
                        if (fileNatureABBR == "TR")
                        {
                            fileNature = "Travaux";
                        }

                        // Worksheet
                        fileSheet = mainWss[iterations].Name;

                        // Deleting empty cc rows (rows that represent totals etc...)
                        ccIndex = findCell(mainPkgs[iterations], fileSheet, "Classe de compte"); // Finding ccIndex
                        lastRow = mainWss[iterations].Dimension.End.Row; // Finding Last Row
                        for (int row = ccIndex[0] + 1; row <= lastRow; row++)
                        {
                            if (mainWss[iterations].Cells[row, ccIndex[1]].Value == null || string.IsNullOrEmpty(mainWss[iterations].Cells[row, ccIndex[1]].Value.ToString()))
                            {
                                rowsToDelete.Add(row); // Adding to a list so it can be removed later
                            }
                        }
                        foreach (int row in rowsToDelete.OrderByDescending(r => r))
                        {
                            mainWss[iterations].DeleteRow(row); // Deleting items in the list
                        }
                        lastRow = mainWss[iterations].Dimension.End.Row; // Resetting Last Row (because there are less rows now)

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

                        // Setting indexes
                        anneeIndex = findCell(mainPkgs[iterations], fileSheet, "Année");
                        natureIndex = findCell(mainPkgs[iterations], fileSheet, "Nature");
                        concatIndex = findCell(mainPkgs[iterations], fileSheet, "Concat");
                        tcIndex = findCell(mainPkgs[iterations], fileSheet, "Type client");
                        gpeIndex = findCell(mainPkgs[iterations], fileSheet, "GpeStrReg");
                        montantEchuIndex = findCell(mainPkgs[iterations], fileSheet, "Montant échu");
                        montantRegleIndex = findCell(mainPkgs[iterations], fileSheet, "Montant réglé");

                        // Filling Annee and Nature
                        mainWss[iterations].Cells[anneeIndex[0] + 1, anneeIndex[1], lastRow, anneeIndex[1]].Value = fileYear;
                        mainWss[iterations].Cells[natureIndex[0] + 1, natureIndex[1], lastRow, natureIndex[1]].Value = fileNature;

                        // Changing Type Client accordingly
                        for (int row = tcIndex[0] + 1; row <= lastRow; row++)
                        {
                            var tcValue = mainWss[iterations].Cells[row, tcIndex[1]].Value;
                            switch (tcValue)
                            {
                                case "BT":
                                case "CB":
                                case "CX":
                                case "EB":
                                case "EC":
                                case "EP":
                                case "NA":
                                case "PP":
                                    mainWss[iterations].Cells[row, tcIndex[1]].Value = "BT";
                                    break;
                                case "MT":
                                case "CM":
                                case "EM":
                                case "GC":
                                case "HT":
                                    mainWss[iterations].Cells[row, tcIndex[1]].Value = "MT";
                                    break;
                            }
                        }

                        // Changing Classe de Compte accordingly
                        if (fileNatureABBR == "EN") // For Energie
                        {
                            for (int row = ccIndex[0] + 1; row <= lastRow; row++)
                            {
                                var ccValue = mainWss[iterations].Cells[row, ccIndex[1]].Value;
                                switch (ccValue)
                                {
                                    case "PALAIS ROYAL":
                                    case "Administrations":
                                        mainWss[iterations].Cells[row, ccIndex[1]].Value = "Administrations";
                                        break;
                                    case "Autres Etablissements  Publics":
                                    case "Stés nationales":
                                        mainWss[iterations].Cells[row, ccIndex[1]].Value = "Stés nationales";
                                        break;
                                    case "Clients occasionnels":
                                    case "Particuliers":
                                        mainWss[iterations].Cells[row, ccIndex[1]].Value = "Particuliers";
                                        break;
                                    case "Multi-Contrats (Régl Reg) Autres":
                                        mainWss[iterations].Cells[row, ccIndex[1]].Value = "Multi-Contrats (Régl Regional)";
                                        break;
                                    case "Multi-Contrats(Régl Centr)Administration":
                                        mainWss[iterations].Cells[row, ccIndex[1]].Value = "Multi-Contrats(Régl Centr)Administration";
                                        break;
                                }
                            }
                        }


                        if (fileNatureABBR == "TR") // For Travaux
                        {
                            for (int row = ccIndex[0] + 1; row <= lastRow; row++)
                            {
                                var ccValue = mainWss[iterations].Cells[row, ccIndex[1]].Value;
                                switch (ccValue)
                                {
                                    case "Administrations":
                                        mainWss[iterations].Cells[row, ccIndex[1]].Value = "Administrations";
                                        break;
                                    case "Autres Etablissements  Publics":
                                        mainWss[iterations].Cells[row, ccIndex[1]].Value = "Stés nationales";
                                        break;
                                    case "Les agents ONE":
                                        mainWss[iterations].Cells[row, ccIndex[1]].Value = "Particuliers";
                                        break;
                                    case "Multi-Contrats (Régl Reg) Autres":
                                        mainWss[iterations].Cells[row, ccIndex[1]].Value = "Multi-Contrats (Régl Regional)";
                                        break;
                                }
                            }
                        }



                        // Dividing montants by 1000
                        for (int row = montantEchuIndex[0] + 1; row <= lastRow; row++)
                        {
                            var montantToDecimal = Convert.ToDecimal(mainWss[iterations].Cells[row, montantEchuIndex[1]].Value);
                            mainWss[iterations].Cells[row, montantEchuIndex[1]].Value = montantToDecimal / 1000;
                        }

                        for (int row = montantRegleIndex[0] + 1; row <= lastRow; row++)
                        {
                            var montantToDecimal = Convert.ToDecimal(mainWss[iterations].Cells[row, montantRegleIndex[1]].Value);
                            mainWss[iterations].Cells[row, montantRegleIndex[1]].Value = montantToDecimal / 1000;
                        }

                        // Changing GpeStrReg Accordingly
                        for (int row = gpeIndex[0] + 1; row <= lastRow; row++)
                        {
                            var grpValue = mainWss[iterations].Cells[row, gpeIndex[1]].Value;
                            if (grpValue != null)
                            {
                                if (grpValue.ToString() == "AGENCE DE SERVICES PROVINCIALE LAAYOUNE")
                                {
                                    mainWss[iterations].Cells[row, gpeIndex[1]].Value = "Agence de Services Provinciale Laâyoune";
                                }
                                if (grpValue.ToString() == "AGENCE DE SERVICES LAKHSSASS")
                                {
                                    mainWss[iterations].Cells[row, gpeIndex[1]].Value = "AGENCE DE SERVICES T. LAKHSSASS";
                                }
                                if (grpValue.ToString() == "SUCCURSALE BIR GANDOUZ")
                                {
                                    mainWss[iterations].Cells[row, gpeIndex[1]].Value = "Succursale Bir Gandouz";
                                }
                            }
                        }

                        // Filling Concat Column
                        for (int row = concatIndex[0] + 1; row <= lastRow; row++)
                        {
                            var agenceValue = mainWss[iterations].Cells[row, gpeIndex[1]].Value;
                            var ccValue = mainWss[iterations].Cells[row, ccIndex[1]].Value;
                            var tcValue = mainWss[iterations].Cells[row, tcIndex[1]].Value;
                            if (agenceValue != null && ccValue != null && tcValue != null)
                            {
                                agenceValue = agenceValue.ToString();
                                ccValue = ccValue.ToString();
                                tcValue = tcValue.ToString();
                                mainWss[iterations].Cells[row, concatIndex[1]].Value = agenceValue + fileNature + ccValue + tcValue;
                            }
                        }
                        mainWss[iterations].Column(concatIndex[1]).AutoFit();
                        mainWss[iterations].Column(concatIndex[1]).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;




                        rowsToDelete.Clear(); // Resetting this list so it can be used by next package
                        mainPkgs[iterations].Save(); // Saving modifications
                        iterations++; // Iterating through the files

                        if (pbarMain.Value + pbarInt > 100) { pbarMain.Value = 100; }
                        else { pbarMain.Value += pbarInt; }
                    }               // foreach

                    MessageBox.Show("Les fichiers selectionnées sont modifiées!", "Modification de fichiers", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }

                catch (Exception ex)
                {
                    MessageBox.Show("Exception rencontrée!\n\n" + ex, "Erreur!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        } //private void btnModify_Click(object sender, EventArgs e)



        // // // Global methods


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
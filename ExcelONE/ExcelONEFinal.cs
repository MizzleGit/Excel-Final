using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ExcelONE
{
    public partial class ExcelONEFinal : Form
    {
        // Variables
        // // String arrays and lists
        string[]? filePaths = null;
        List<string> fileYears = new List<string>();
        // //

        // // Integers
        int numYears;
        // //

        // // Strings
        string folderPath = "";
        string fileName = "";
        string masterPathGlobal = "";
        // //

        // // Excel variables
        List<ExcelPackage> mainPkgs = new List<ExcelPackage>();
        List<ExcelWorksheet> mainWss = new List<ExcelWorksheet>();
        // //

        // // Booleans
        bool filesOpened = false;
        // //
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
                bool tryPassed = false;
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
                    mainPkgs.Clear();
                    mainWss.Clear();
                    foreach (string path in filePaths)
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        mainPkgs.Add(new ExcelPackage(path));
                    }
                    for (int i = 0; i < mainPkgs.Count(); i++)
                    {
                        mainWss.Add(mainPkgs[i].Workbook.Worksheets[0]);
                    }
                    tryPassed = true;

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
                    // Clearing the years list
                    fileYears.Clear();
                    foreach (string path in filePaths)
                    {
                        // Retrieving basic file information
                        fileName = Path.GetFileNameWithoutExtension(path);
                        fileYear = getYearAndNature(fileName)[0];
                        fileNatureABBR = getYearAndNature(fileName)[1];

                        // Adding year to years list for global file button
                        fileYears.Add(fileYear);

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

                        // Skipping the file if it already has Annee in it | | | Annee was picked randomly, this is to ensure files don't get modified twice which can give wrong resutls
                        if (!findCell(mainPkgs[iterations], fileSheet, "Année").SequenceEqual(emptyCellIndex))
                        {
                            continue;
                        }

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

                        //
                        //
                        //

                        rowsToDelete.Clear(); // Resetting this list so it can be used by next package
                        mainPkgs[iterations].Save(); // Saving modifications
                        iterations++; // Iterating through the files

                        if (pbarMain.Value + pbarInt > 100) { pbarMain.Value = 100; }
                        else { pbarMain.Value += pbarInt; }
                    }               // foreach
                    if (tryPassed) { MessageBox.Show("Les fichiers selectionnées sont modifiées!", "Modification de fichiers", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                }

                catch (Exception ex)
                {
                    MessageBox.Show("Exception rencontrée!\n\n" + ex, "Erreur!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else { MessageBox.Show("Ouvrez d'abord un ou plusieurs fichiers!", "Erreur", MessageBoxButtons.OK, MessageBoxIcon.Error); }
        } //private void btnModify_Click(object sender, EventArgs e)

        //
        //
        //

        private void btnGlobal_Click(object sender, EventArgs e)
        {
            if (!OnlyFourValues(fileYears) || mainPkgs.Count != 8) // Ensuring exactly 8 files with exactly 4 different years are put into global file
            {
                MessageBox.Show("Assurez-vous qu'il n'y ait que 8 fichiers sélectionnés et que seules 4 années soient utilisées pour chaque nature!", "Erreur!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                pbarMain.Value = 0;
                // Variables
                ExcelPackage masterPkg = new ExcelPackage();
                ExcelWorksheet masterWs = masterPkg.Workbook.Worksheets.Add("Master");
                ExcelRange header = mainWss[0].Cells[mainWss[0].Dimension.Start.Row, mainWss[0].Dimension.Start.Column, mainWss[0].Dimension.Start.Row, mainWss[0].Dimension.End.Column];
                ExcelRange[] masterData = new ExcelRange[8];
                string[] fourFileYears = GetOnlyFourValues(fileYears);
                string masterPath = "";
                pbarMain.Value = 10;

                // Adding all data into an ExcelRange Array
                for (int i = 0; i < 8; i++)
                {
                    pbarMain.Value++;
                    masterData[i] = mainWss[i].Cells[mainWss[i].Dimension.Start.Row + 1, mainWss[i].Dimension.Start.Column, mainWss[i].Dimension.End.Row, mainWss[i].Dimension.End.Column];
                }
                pbarMain.Value = 20;
                // Creating the master file
                try
                {
                    header.Copy(masterWs.Cells[1, 1]);
                    for (int i = 0; i < 8; i++)
                    {
                        masterData[i].Copy(masterWs.Cells[masterWs.Dimension.End.Row + 1, masterWs.Dimension.Start.Column]);
                    }
                    for (int i = 1; i <= masterWs.Dimension.End.Column; i++)
                    {
                        masterWs.Column(i).AutoFit();
                    }
                    pbarMain.Value = 70;
                    masterWs.Column(masterWs.Dimension.End.Column).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                    ExcelWorksheet masterPivotWs = masterPkg.Workbook.Worksheets.Add("Pivot");
                    ExcelRange masterUsedRange = masterWs.Cells[masterWs.Dimension.Start.Row, masterWs.Dimension.Start.Column, masterWs.Dimension.End.Row, masterWs.Dimension.End.Column];
                    ExcelPivotTable masterPivot = masterPivotWs.PivotTables.Add(masterPivotWs.Cells[1, 1], masterUsedRange, "Pivot");
                    masterPivot.DataOnRows = false;
                    masterPivot.ColumnGrandTotals = false;
                    masterPivot.RowGrandTotals = false;
                    masterPivot.RowFields.Add(masterPivot.Fields["Concat"]);
                    masterPivot.ColumnFields.Add(masterPivot.Fields["Année"]);
                    ExcelPivotTableDataField masterMontantE = masterPivot.DataFields.Add(masterPivot.Fields["Montant échu"]);
                    ExcelPivotTableDataField masterMontantR = masterPivot.DataFields.Add(masterPivot.Fields["Montant réglé"]);
                    masterMontantE.Function = DataFieldFunctions.Sum;
                    masterMontantR.Function = DataFieldFunctions.Sum;
                    pbarMain.Value = 90;
                    foreach (ExcelPackage pkg in mainPkgs)
                    {
                        pkg.Dispose();
                    }
                    try
                    {
                        masterPath = folderPath + "/global.xlsx";
                        masterPkg.SaveAs(folderPath + "/global.xlsx");
                        pbarMain.Value = 100;
                        masterPathGlobal = masterPkg.File.FullName;
                        MessageBox.Show("Le fichier global a été crée! Merci de vérifier le dossier original.", "Fichier global", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch
                    {
                        folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                        masterPath = Path.Combine(folderPath, "global.xlsx");
                        masterPkg.SaveAs(masterPath);
                        pbarMain.Value = 100;
                        masterPathGlobal = masterPkg.File.FullName;
                        MessageBox.Show("Le fichier global a été crée! Merci de vérifier votre bureau.", "Fichier global", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Exception rencontrée!\n\n" + ex, "Erreur!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        } // private void btnGlobal_Click(object sender, EventArgs e)

        //
        //
        //

        private void btnDestination_Click(object sender, EventArgs e)
        {
            // Variables
            string destinationPath = "";
            string destinationName = "";
            bool noConcat = false;
            bool donneeFound = false;

            Thread.Sleep(500);
            openDestination = new OpenFileDialog();
            openDestination.Filter = "Excel Files (*.xlsx) | *.xlsx";
            if (openDestination.ShowDialog() == DialogResult.OK)
            {

                // Getting general file data
                destinationPath = openDestination.FileName;
                destinationName = Path.GetFileNameWithoutExtension(destinationPath);

                // Opening Pivot
                Excel.Application pivotApp = new Excel.Application();
                Excel.Workbook pivotWorkbook = pivotApp.Workbooks.Open(masterPathGlobal, ReadOnly: false);
                Excel.Worksheet pivotWs = (Excel.Worksheet)pivotWorkbook.Sheets[2];
                pivotWs.Activate();
                Excel.Range pivotRange = pivotWs.UsedRange;
                Excel.Application destinationApp = new Excel.Application();
                Excel.Workbook destinationWorkbook = destinationApp.Workbooks.Open(destinationPath, ReadOnly: false);
                foreach (Excel.Worksheet worksheet in destinationWorkbook.Worksheets)
                {
                    if (worksheet.Name == "données TR")
                    {
                        donneeFound = true;
                        // Finding worksheet and selecting used range
                        Excel.Worksheet destinationSheet = (Excel.Worksheet)destinationWorkbook.Sheets["données TR"];
                        Excel.Range destinationRange = destinationSheet.UsedRange;

                        // Resetting progress bar
                        pbarMain.Value = 0;

                        // Finding Concat in destination folder
                        Excel.Range concatRange = destinationRange.Find("Concat");
                        if (concatRange == null) { noConcat = true; break; }

                        // Setting indexes for concat range
                        int crFirstRow = concatRange.Row + 1;
                        int crCol = concatRange.Column;

                        // Finding last used row
                        Excel.Range concatUsed = concatRange.End[Excel.XlDirection.xlDown];
                        int concatLastRow = concatUsed.Row + concatUsed.Rows.Count - 1;

                        // Indexing concat and year columns
                        pbarMain.Value++;
                        Excel.Range colToFind = (Excel.Range)destinationSheet.Range[destinationSheet.Cells[crFirstRow, crCol], destinationSheet.Cells[concatLastRow, crCol]];
                        pbarMain.Value++;
                        Excel.Range colToChange1 = (Excel.Range)destinationSheet.Range[destinationSheet.Cells[crFirstRow, crCol + 1], destinationSheet.Cells[concatLastRow, crCol + 1]];
                        pbarMain.Value++;
                        Excel.Range colToChange2 = (Excel.Range)destinationSheet.Range[destinationSheet.Cells[crFirstRow, crCol + 2], destinationSheet.Cells[concatLastRow, crCol + 2]];
                        pbarMain.Value++;
                        Excel.Range colToChange3 = (Excel.Range)destinationSheet.Range[destinationSheet.Cells[crFirstRow, crCol + 3], destinationSheet.Cells[concatLastRow, crCol + 3]];
                        pbarMain.Value++;
                        Excel.Range colToChange4 = (Excel.Range)destinationSheet.Range[destinationSheet.Cells[crFirstRow, crCol + 4], destinationSheet.Cells[concatLastRow, crCol + 4]];
                        pbarMain.Value++;
                        Excel.Range colToChange5 = (Excel.Range)destinationSheet.Range[destinationSheet.Cells[crFirstRow, crCol + 5], destinationSheet.Cells[concatLastRow, crCol + 5]];
                        pbarMain.Value++;
                        Excel.Range colToChange6 = (Excel.Range)destinationSheet.Range[destinationSheet.Cells[crFirstRow, crCol + 6], destinationSheet.Cells[concatLastRow, crCol + 6]];
                        pbarMain.Value++;
                        Excel.Range colToChange7 = (Excel.Range)destinationSheet.Range[destinationSheet.Cells[crFirstRow, crCol + 7], destinationSheet.Cells[concatLastRow, crCol + 7]];
                        pbarMain.Value = 9;
                        Excel.Range colToChange8 = (Excel.Range)destinationSheet.Range[destinationSheet.Cells[crFirstRow, crCol + 8], destinationSheet.Cells[concatLastRow, crCol + 8]];
                        pbarMain.Value = 10;

                        lblWait.Visible = true;

                        colToChange1.Value = pivotApp.WorksheetFunction.VLookup(colToFind, pivotRange, 2, 0);
                        pbarMain.Value = 15;
                        colToChange2.Value = pivotApp.WorksheetFunction.VLookup(colToFind, pivotRange, 3, 0);
                        pbarMain.Value = 20;
                        colToChange3.Value = pivotApp.WorksheetFunction.VLookup(colToFind, pivotRange, 4, 0);
                        pbarMain.Value = 30;
                        colToChange4.Value = pivotApp.WorksheetFunction.VLookup(colToFind, pivotRange, 5, 0);
                        pbarMain.Value = 40;
                        colToChange5.Value = pivotApp.WorksheetFunction.VLookup(colToFind, pivotRange, 6, 0);
                        pbarMain.Value = 50;
                        colToChange6.Value = pivotApp.WorksheetFunction.VLookup(colToFind, pivotRange, 7, 0);
                        pbarMain.Value = 60;
                        colToChange7.Value = pivotApp.WorksheetFunction.VLookup(colToFind, pivotRange, 8, 0);
                        pbarMain.Value = 70;
                        colToChange8.Value = pivotApp.WorksheetFunction.VLookup(colToFind, pivotRange, 9, 0);
                        pbarMain.Value = 80;

                        // Fixing null errors for empty cells
                        Excel.Range[] ranges = new Excel.Range[] { colToChange1, colToChange2, colToChange3, colToChange4, colToChange5, colToChange6, colToChange7, colToChange8 };
                        foreach (Excel.Range range in ranges)
                        {
                            foreach (Excel.Range cell in range)
                            {
                                if (cell.Value == -2146826246.00)
                                {
                                    cell.Value = 0;
                                }
                            }
                        }
                        pbarMain.Value = 90;

                        // Saving changes and closing
                        pivotWorkbook.Close(SaveChanges: true);
                        destinationWorkbook.Close(SaveChanges: true);
                        pivotApp.Quit();
                        destinationApp.Quit();
                        pbarMain.Value = 99;

                        // Releasing COM Objects
                        Marshal.ReleaseComObject(colToChange1);
                        Marshal.ReleaseComObject(colToChange2);
                        Marshal.ReleaseComObject(colToChange3);
                        Marshal.ReleaseComObject(colToChange4);
                        Marshal.ReleaseComObject(colToChange5);
                        Marshal.ReleaseComObject(colToChange6);
                        Marshal.ReleaseComObject(colToChange7);
                        Marshal.ReleaseComObject(colToChange8);
                        Marshal.ReleaseComObject(colToFind);
                        Marshal.ReleaseComObject(concatRange);
                        Marshal.ReleaseComObject(pivotWs);
                        Marshal.ReleaseComObject(destinationSheet);
                        Marshal.ReleaseComObject(pivotWs);
                        Marshal.ReleaseComObject(destinationRange);
                        Marshal.ReleaseComObject(pivotWorkbook);
                        Marshal.ReleaseComObject(destinationWorkbook);
                        Marshal.ReleaseComObject(pivotApp);
                        Marshal.ReleaseComObject(destinationApp);
                        pbarMain.Value = 100;
                        MessageBox.Show("La feuil a été modifié\n\nMerci de fermer le programme avant de l'utiliser une autre fois", destinationName, MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // Breaking after finding données TR
                        break;
                    }
                } // foreach
                if (!donneeFound)
                {
                    MessageBox.Show("Erreur, données TR n'existe pas dans ce fichier", "Erreur!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                if (noConcat)
                {
                    MessageBox.Show("'Concat' n'a pas été trouvé. Assurez-vous de nommer la cellule vide 'Concat' avec un C majuscule.", "Erreur!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                lblWait.Visible = false;
            } // if Dialog
        } // private void btnDestination_Click(object sender, EventArgs e)






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

        // OnlyFourValues
        bool OnlyFourValues(List<string> fnEnteredYears)
        {
            List<string> fnYears = new List<string>();
            foreach (string fnYear in fnEnteredYears)
            {
                if (!fnYears.Contains(fnYear))
                {
                    fnYears.Add(fnYear);
                }
            }
            if (fnYears.Count == 4)
            {
                return true;
            }
            return false;
        }
        // OnlyFourValues

        // GetOnlyFourValues
        string[] GetOnlyFourValues(List<string> fnEnteredYears)
        {
            List<string> fnYears = new List<string>();
            foreach (string fnYear in fnEnteredYears)
            {
                if (!fnYears.Contains(fnYear))
                {
                    fnYears.Add(fnYear);
                }
            }
            string[] fnYearsArray = fnYears.ToArray();
            Array.Sort(fnYearsArray);
            return fnYearsArray;
        }
        // GetOnlyFourValues


    } // public partial class ExcelONEFinal : Form
} // namespace ExcelONE
using Manatee.Trello;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata.Ecma335;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Trello.Models;

namespace Trello
{
    public class Excel
    {
        public static void WriteToExcel() 
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo(@"data.xlsx");
            // SHEETEK LÉTREHOZÁSA/NULLÁZÁSA
            InitializeSheets(file);
            // SHEETEK KITÖLTÉSE
            FillSheets(file);
            var settings = Settings.LoadSettings();
            

        }
        static void InitializeSheets(FileInfo file)
        {
            var sheetSettings = Settings.GetSheets();
            // MEGADOTT SHEET NEVEK KIGYŰJTÉSE
            List<string> sheetNames = new List<string> 
                { sheetSettings.Summary ?? throw new Exception("Summary sheet name is missing"), sheetSettings.Cards ?? throw new Exception("Cards sheet name is missing") };

            // MEGADOTT SHEETEK LÉTREHOZÁSA/NULLÁZÁSA
            foreach (string name in sheetNames)
                InitializeSheets(file, name);

            // CSAK MEGADOTT NEVŰ SHEETEK LÉTEZZENEK
            DeleteUndefinedSheets(file, sheetNames);
        }
        static void DeleteUndefinedSheets(FileInfo file, List<string> definedNames) 
        {
            using (var package = new ExcelPackage(file))
            {
                // Az összes munkalap nevének lekérdezése
                var sheetNames = package.Workbook.Worksheets.Select(ws => ws.Name).ToList();

                // Azoknak a munkalapoknak a neve, amelyek nem szerepelnek a definedNames listában
                var sheetsToDelete = sheetNames.Except(definedNames).ToList();

                // Munkalapok törlése
                foreach (var sheetName in sheetsToDelete)
                {
                    var worksheet = package.Workbook.Worksheets[sheetName];
                    worksheet.Cells.Clear(); // Munkalap tartalmának törlése
                    package.Workbook.Worksheets.Delete(worksheet); // Munkalap törlése
                }

                // Menti a fájlt
                package.Save();
            }
        }
        static void InitializeSheets(FileInfo file, string name)
        {
            using (var package = new ExcelPackage(file))
            {
                // Ellenőrizzük, hogy létezik-e már az adott nevű munkalap
                if (package.Workbook.Worksheets.Any(ws => ws.Name == name))
                {
                    // Ha már létezik, akkor az adott munkalapot válasszuk ki
                    var worksheet = package.Workbook.Worksheets[name];
                    // Töröljük a munkalap tartalmát
                    worksheet.Cells.Clear();
                }
                else
                {
                    // Ha nem létezik, akkor létrehozzuk az új munkalapot
                    var worksheet = package.Workbook.Worksheets.Add(name);
                }

                // Menti a fájlt
                package.Save();
            }
        }
        static void FillSheets(FileInfo file) 
        {
            // ÖSSZESÍTŐ SHEET KITÖLTÉSE
            FillSummarySheet(file);
            // KÁRTYÁK SHEET KITÖLTÉSE
            FillCardsSheet(file);
        }
        static void FillSummarySheet(FileInfo file) 
        {
            using (ExcelPackage package = new ExcelPackage(file))
            {
                try
                {
                    // SUMMARY SHEET VÁLTOZÓBA GYŰJTÉSE - HA NINCS -> ERROR
                    string summarySheetName = Settings.GetSheets().Summary ?? throw new Exception("Summary sheet name is required");
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[summarySheetName];
                    if (worksheet == null)
                        throw new Exception("Something went wrong: summary sheet doesn't exist");

                    // ADATBÁZISBÓL SUMMARY ADATOK KISZEDÉSE
                    var dbOptions = new DbContextOptionsBuilder<ApplicationDbContext>()
                            .UseSqlServer(Settings.LoadSettings().ConnectionString ?? throw new Exception("Connection string is required"))
                            .Options;
                    List<Completed> data;
                    using (var dbContext = new ApplicationDbContext(dbOptions))
                    {
                        try
                        {
                            data = dbContext.Completed?.ToList() ?? throw new Exception("Something went wrong: summary in DB is null");
                        }
                        catch { throw new Exception("Database error while filling summary sheet"); }
                    }
                    // SUMMARY SHEET KERETÉNEK LÉTREHOZÁSA
                    CreateBaseSummarySheet(worksheet);

                    int excelRow = worksheet.Dimension?.End.Row + 1 ?? 3;
                    foreach (var dbRow in data)
                    {
                        worksheet.Cells[excelRow, 1].Value = dbRow.Date.ToString("yyyy.MMMM");
                        worksheet.Cells[excelRow, 2].Value = dbRow.AllCompleted;
                        worksheet.Cells[excelRow, 3].Value = dbRow.ShoperiaAllCompleted;
                        worksheet.Cells[excelRow, 4].Value = dbRow.Shoperia_W1;
                        worksheet.Cells[excelRow, 5].Value = dbRow.Shoperia_W2;
                        worksheet.Cells[excelRow, 6].Value = dbRow.Shoperia_W3;
                        worksheet.Cells[excelRow, 7].Value = dbRow.Shoperia_UnWeighted;
                        worksheet.Cells[excelRow, 8].Value = dbRow.Home12AllCompleted;
                        worksheet.Cells[excelRow, 9].Value = dbRow.Home12_W1;
                        worksheet.Cells[excelRow, 10].Value = dbRow.Home12_W2;
                        worksheet.Cells[excelRow, 11].Value = dbRow.Home12_W3;
                        worksheet.Cells[excelRow, 12].Value = dbRow.Home12_UnWeighted;
                        worksheet.Cells[excelRow, 13].Value = dbRow.XpressAllCompleted;
                        worksheet.Cells[excelRow, 14].Value = dbRow.Xpress_W1;
                        worksheet.Cells[excelRow, 15].Value = dbRow.Xpress_W2;
                        worksheet.Cells[excelRow, 16].Value = dbRow.Xpress_W3;
                        worksheet.Cells[excelRow, 17].Value = dbRow.Xpress_UnWeighted;
                        worksheet.Cells[excelRow, 18].Value = dbRow.MatebikeAllCompleted;
                        worksheet.Cells[excelRow, 19].Value = dbRow.Matebike_W1;
                        worksheet.Cells[excelRow, 20].Value = dbRow.Matebike_W2;
                        worksheet.Cells[excelRow, 21].Value = dbRow.Matebike_W3;
                        worksheet.Cells[excelRow, 22].Value = dbRow.Matebike_UnWeighted;
                        excelRow++;
                    }
                    SetSummarySheetFormats(worksheet);
                    
                    CreateInterpretationGuide(worksheet);

                    worksheet.Cells[worksheet.Dimension?.Address].AutoFitColumns();
                    package.Save();
                }
                catch { throw new Exception("Excel error while filling summary sheet"); }
            }
        }
        static void CreateBaseSummarySheet(ExcelWorksheet sheet) 
        {
            // JELENLEGI EXCEL SOR
            int excelRow = 1;
            
            // DÁTUM OSZLOP CÍMÉNEK (ELSŐ KETTŐ CELLA AZ "A" OSZLOPBAN) BEÁLLÍTÁSA
            var dateCell = sheet.Cells["A1:A2"];
            dateCell.Merge = true;

            // ÖSSZESÍTETT OSZLOP CÍMÉNEK (ELSŐ KETTŐ CELLA A "B" OSZLOPBAN) BEÁLLÍTÁSA
            var summaryCell = sheet.Cells["B1:B2"];
            summaryCell.Merge = true;

            // CÉGEK NEVEIVEL FELTÖLTÉS
            sheet.Cells[excelRow, 1].Value = "Dátum";
            sheet.Cells[excelRow, 2].Value = "Összes";
            sheet.Cells[excelRow, 3].Value = "SHOPERIA";
            sheet.Cells[excelRow, 8].Value = "HOM12";
            sheet.Cells[excelRow, 13].Value = "XPRESS";
            sheet.Cells[excelRow, 18].Value = "MATEBIKE";

            // ÚJ SOR
            excelRow++;

            for (int i = 0; i < 4; i++)
            {
                var startColumn = 3 + (i * 5); // C oszloptól (3. oszlop) kezdve, minden 5-dik kijelölése
                var endColumn = startColumn + 4; // start oszlop + 4 oszlop (egy cég 4 oszlopot foglal le)
                sheet.Cells[1, startColumn, 1, endColumn].Merge = true; // cég címek celláinak egyesítése

                // CÉGENKÉNT ALL W1 W2 W3 UW OSZLOPCÍMEK BEÁLLÍTÁSA
                sheet.Cells[excelRow, 3 + i * 5].Value = "All";
                sheet.Cells[excelRow, 4 + i * 5].Value = "W1";
                sheet.Cells[excelRow, 5 + i * 5].Value = "W2";
                sheet.Cells[excelRow, 6 + i * 5].Value = "W3";
                sheet.Cells[excelRow, 7 + i * 5].Value = "UW";
            }
        }
        static void SetSummarySheetFormats(ExcelWorksheet sheet)
        {
            //PALETTA
            ExcelColorList color = new ExcelColorList();

            int lastRow = sheet.Dimension?.End.Row ?? throw new Exception("dimension is null");
            int lastColumn = sheet.Dimension.End.Column;

            // Cellák kijelölése a teljes tartományban
            var range = sheet.Cells[1, 1, lastRow, lastColumn];
            range.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
            var borderStyle = ExcelBorderStyle.Thin;

            // ÖSSZES CELLÁRA PATTERN STYLE, IGAZÍTÁS, FÉLKÖVÉRSÉG ÉS SZEGÉLY BEÁLLÍTÁSA
            for (int row = 1; row <= lastRow; row++)
                for (int col = 1; col <= lastColumn; col++)
                {
                    sheet.Cells[row, col].Style.Border.Left.Style = borderStyle;
                    sheet.Cells[row, col].Style.Border.Right.Style = borderStyle;
                    sheet.Cells[row, col].Style.Border.Top.Style = borderStyle;
                    sheet.Cells[row, col].Style.Border.Bottom.Style = borderStyle;
                    sheet.Cells[row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells[row, col].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[row, col].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    sheet.Cells[row, col].Style.Font.Bold = true;
                }

            // CÉG CÍMEK FORMÁZÁSA
            for (int i = 0; i < 4; i++)
            {
                var companyCell = sheet.Cells[1, 3 + i * 5];
                companyCell.Style.Fill.BackgroundColor.SetColor(color.ShopColors[i].Title);
                companyCell.Style.Font.Color.SetColor(System.Drawing.Color.White);
            }


            int excelRow = 2;

            // "All", "W1", "W2", "W3", "UW" cellák formázása
            for (int i = 0; i < 4; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    var valueCell = sheet.Cells[excelRow, 3 + i * 5 + j, lastRow, 3 + i * 5 + j];
                    valueCell.Style.Fill.BackgroundColor.SetColor(color.DateColor);
                }
                
                //ALL + UW
                sheet.Cells[excelRow, 3 + i * 5 + 0, lastRow, 3 + i * 5 + 0].Style.Fill.BackgroundColor.SetColor(color.ShopColors[i].Default);
                sheet.Cells[excelRow, 3 + i * 5 + 4, lastRow, 3 + i * 5 + 4].Style.Fill.BackgroundColor.SetColor(color.ShopColors[i].Default);
                //W1
                sheet.Cells[excelRow, 3 + i * 5 + 1, lastRow, 3 + i * 5 + 1].Style.Fill.BackgroundColor.SetColor(color.ShopColors[i].W1);
                //W2
                sheet.Cells[excelRow, 3 + i * 5 + 2, lastRow, 3 + i * 5 + 2].Style.Fill.BackgroundColor.SetColor(color.ShopColors[i].W2);
                //W3
                sheet.Cells[excelRow, 3 + i * 5 + 3, lastRow, 3 + i * 5 + 3].Style.Fill.BackgroundColor.SetColor(color.ShopColors[i].W3);

            }
            //DATE
            sheet.Cells[1, 1, lastRow, 1].Style.Fill.BackgroundColor.SetColor(color.DateColor);
            //SUMMARY
            sheet.Cells[1, 2, lastRow, 2].Style.Fill.BackgroundColor.SetColor(color.SummaryColor);


        }
        static void CreateInterpretationGuide(ExcelWorksheet sheet) 
        {
            sheet.Cells[1, 25, 1,29].Merge = true;
            sheet.Cells[1, 25].Value = "Értelmezési segédlet";
            // KULCSOK
            sheet.Cells[2, 25, 3, 25].Merge = true;
            sheet.Cells[2, 25].Value = "All";
            sheet.Cells[2, 26, 3, 26].Merge = true;
            sheet.Cells[2, 26].Value = "W1";
            sheet.Cells[2, 27, 3, 27].Merge = true;
            sheet.Cells[2, 27].Value = "W2";
            sheet.Cells[2, 28, 3, 28].Merge = true;
            sheet.Cells[2, 28].Value = "W3";
            sheet.Cells[2, 29, 3, 29].Merge = true;
            sheet.Cells[2, 29].Value = "UW";
            // MAGYARÁZAT
            sheet.Cells[4, 25, 5, 25].Merge = true;
            sheet.Cells[4, 25].Value = "Összes projekt";
            sheet.Cells[4, 26, 5, 26].Merge = true;
            sheet.Cells[4, 26].Value = "Kis projekt";
            sheet.Cells[4, 27, 5, 27].Merge = true;
            sheet.Cells[4, 27].Value = "Közepes projekt";
            sheet.Cells[4, 28, 5, 28].Merge = true;
            sheet.Cells[4, 28].Value = "Nagy projekt";
            sheet.Cells[4, 29, 5, 29].Merge = true;
            sheet.Cells[4, 29].Value = "Besorolatlan projekt";

            // Értelmezési segédlet formázása
            var ertelmezesiSegedletCella = sheet.Cells[1, 25];
            ertelmezesiSegedletCella.Style.Font.Color.SetColor(System.Drawing.Color.White);
            ertelmezesiSegedletCella.Style.Fill.PatternType = ExcelFillStyle.Solid;
            ertelmezesiSegedletCella.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Black);
            ertelmezesiSegedletCella.Style.Font.Bold = true;
            

            // Oszlopok formázása
            for (int i = 25; i <= 29; i++)
            {
                var oszlop = sheet.Cells[2, i, 5, i];
                oszlop.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                oszlop.Style.Fill.BackgroundColor.SetColor(
                    System.Drawing.Color.FromArgb(
                        System.Drawing.Color.FromArgb(211, 211, 211).R - (i - 25) * 10,
                        System.Drawing.Color.FromArgb(211, 211, 211).G - (i - 25) * 10,
                        System.Drawing.Color.FromArgb(211, 211, 211).B - (i - 25) * 10
                    )
                );
            }

            // Szegély hozzáadása
            var range = sheet.Cells[1, 25, 5, 29];
            var border = range.Style.Border;
            border.Bottom.Style = border.Top.Style = border.Left.Style = border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
            border.Bottom.Color.SetColor(System.Drawing.Color.Black);
            border.Top.Color.SetColor(System.Drawing.Color.Black);
            border.Left.Color.SetColor(System.Drawing.Color.Black);
            border.Right.Color.SetColor(System.Drawing.Color.Black);

            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;


        }
        static void FillCardsSheet(FileInfo file) 
        {
            using (ExcelPackage package = new ExcelPackage(file))
            {
                try
                {
                    // CARDS SHEET VÁLTOZÓBA GYŰJTÉSE - HA NINCS -> ERROR
                    string cardsSheetName = Settings.GetSheets().Cards ?? throw new Exception("Summary sheet name is required");
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[cardsSheetName];
                    if (worksheet == null)
                        throw new Exception("Something went wrong: summary sheet doesn't exist");

                    // ADATBÁZISBÓL SUMMARY ADATOK KISZEDÉSE
                    var dbOptions = new DbContextOptionsBuilder<ApplicationDbContext>()
                            .UseSqlServer(Settings.LoadSettings().ConnectionString ?? throw new Exception("Connection string is required"))
                            .Options;
                    List<CardModel> data;
                    using (var dbContext = new ApplicationDbContext(dbOptions))
                    {
                        try
                        {
                            data = dbContext.Cards?.ToList() ?? throw new Exception("Something went wrong: summary in DB is null");
                        }
                        catch { throw new Exception("Database error while filling summary sheet"); }
                    }
                    // SUMMARY SHEET KERETÉNEK LÉTREHOZÁSA
                    CreateBaseCardsSheet(worksheet);

                    // ADATOK SORONKÉNTI KITÖLTÉSE
                    int excelRow = worksheet.Dimension?.End.Row + 1 ?? 2;

                    // TRELLO KÁRTYÁK LEKÉRÉSE
                    string boardId = Settings.LoadSettings().BoardId ?? throw new Exception("Board ID can't be null");
                    TrelloFactory factory = new TrelloFactory();
                    var board = factory.Board(boardId);
                    board.Lists.Refresh();
                    var cards = board.Lists.SelectMany(l => l.Cards).ToList();

                    foreach (var dbRow in data)
                    {
                        var trelloCard = cards.Find(c => c.Id.Trim() == dbRow.Id?.Trim());
                        if (trelloCard == null)
                            Console.WriteLine($"NEM TALÁLT TASK: Név: \t {dbRow.Name}");
                        else
                        {
                            worksheet.Cells[excelRow, 1].Value = dbRow.Id;
                            worksheet.Cells[excelRow, 2].Value = dbRow.Name;
                            worksheet.Cells[excelRow, 3].Value = trelloCard.CreationDate.ToString("yyyy-MM-dd");
                            worksheet.Cells[excelRow, 4].Value = dbRow.IsComplete == true ? dbRow.Date.ToString("yyyy-MM-dd") : "";
                            string weight = dbRow.Weight == 1 ? "W1"
                                            : dbRow.Weight == 2 ? "W2"
                                            : dbRow.Weight == 3 ? "W3"
                                            : "UW";
                            worksheet.Cells[excelRow, 5].Value = weight;
                            worksheet.Cells[excelRow, 6].Value = dbRow.IsComplete.ToString();
                            string? shop = Utilities.GetShopByListId(dbRow.ListId ?? throw new Exception("ListID can't be null"));
                            worksheet.Cells[excelRow, 7].Value = shop ?? "EGYÉB";

                            worksheet.Cells[excelRow, 8].Formula = $"=HYPERLINK(\"{trelloCard.Url}\", \"Feladat megnyitása..\")";
                            worksheet.Cells[excelRow, 8].Style.Font.Color.SetColor(System.Drawing.Color.Blue);

                            excelRow++;
                        }
                    }

                    //AUTOMATIKUS CELLAMÉRETEZÉS, SZEGÉLY BEÁLLÍTÁSA
                    int lastRow = worksheet.Dimension?.End.Row ?? throw new Exception("dimension is null");
                    int lastColumn = worksheet.Dimension.End.Column;

                    // Cellák kijelölése a teljes tartományban
                    var range = worksheet.Cells[1, 1, lastRow, lastColumn];
                    range.Style.Border.BorderAround(ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                    var borderStyle = ExcelBorderStyle.Thin;

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    // ÖSSZES CELLÁRA PATTERN STYLE, IGAZÍTÁS, FÉLKÖVÉRSÉG ÉS SZEGÉLY BEÁLLÍTÁSA
                    for (int row = 1; row <= lastRow; row++)
                        for (int col = 1; col <= lastColumn; col++)
                        {
                            worksheet.Cells[row, col].Style.Border.Left.Style = borderStyle;
                            worksheet.Cells[row, col].Style.Border.Right.Style = borderStyle;
                            worksheet.Cells[row, col].Style.Border.Top.Style = borderStyle;
                            worksheet.Cells[row, col].Style.Border.Bottom.Style = borderStyle;
                        }

                    package.Save();
                }
                catch (Exception ex) { throw new Exception("Excel error while filling summary sheet: ", ex); }
            }
        }
        static void CreateBaseCardsSheet(ExcelWorksheet sheet)
        {
            // FEJLÉCEK KITÖLTÉSE
            sheet.Cells[1, 1].Value = "ID";
            sheet.Cells[1, 2].Value = "Név";
            sheet.Cells[1, 3].Value = "Létrehozva";
            sheet.Cells[1, 4].Value = "Elfogadva";
            sheet.Cells[1, 5].Value = "Súlyozás";
            sheet.Cells[1, 6].Value = "Befejezett?";
            sheet.Cells[1, 7].Value = "Cég";
            sheet.Cells[1, 8].Value = "Link";
        }
    
    }
}

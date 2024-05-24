using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using Serilog.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Trello.Models;

namespace Trello
{
    public class Utilities
    {
        /// <summary>
        /// Frissíti a teljesített feladatok számát a megadott kártya alapján.
        /// </summary>
        /// <param name="dbContext">Az adatbázis kapcsolat változója.</param>
        /// <param name="card">A kártya modell, amely alapján a frissítés történik.</param>
        /// <param name="operation">Az operáció típusa: "ADD" a hozzáadáshoz, bármilyen más érték a kivonáshoz.</param>
        /// <exception cref="Exception">Ha a kártya neve, azonosítója vagy lista azonosítója null, kivételt dob.</exception>
        public static void UpdateTes(ApplicationDbContext dbContext, CardModel card, ILogger logger, string operation = "ADD")
        {
            // Ellenőrzi, hogy a kártya szükséges mezői nem null értékűek
            if (card.Name == null || card.Id == null || card.ListId == null)
                throw new Exception("A kártya egyik kötelező mezője null értékű.");

            // Kivonja az év és hónap értékét a kártya dátumából
            int year = card.Date.Year;
            int month = card.Date.Month;

            // Megkeresi a megfelelő dátumot (sort) a 'Completed' táblában, ha nem találja, létrehozza
            var row = dbContext.Completed?.FirstOrDefault(e => e.Date == new DateTime(year, month, 1));
            if (row == null)
            {
                row = dbContext.AddDateToDB(year, month);
                logger.LogInformation("{year} - {month} dátum hozzáadása a Completed táblához", year, month);
            }

            // Lekéri a shop nevét a lista azonosító alapján
            string? shop = GetShopByListId(card.ListId);

            // Ha az adott lista azonosító az "ORDERS"-hez tartozik, akkor a kártya neve alapján lekéri a shop nevét
            if (shop == "ORDERS")
                shop = getShopByCardName(card.Name);

            // Meghatározza a számítás segítőjét; alapértelmezettként hozzáadást végez (érték: 1), kivonás esetén (nem "ADD") az érték: -1
            int calcHelper = 1;
            if (operation != "ADD")
                calcHelper = -1;

            // Csak akkor frissíti az adatokat, ha a shop neve nem null
            if (shop != null)
            {
                // Összesített számláló növelése vagy csökkentése
                row.AllCompleted += calcHelper;

                // A megfelelő shop és súlyozás szerinti számláló frissítése
                switch (shop)
                {
                    case "SHOPERIA":
                        row.ShoperiaAllCompleted += calcHelper;
                        switch (card.Weight)
                        {
                            case 0:
                                row.Shoperia_UnWeighted += calcHelper;
                                break;
                            case 1:
                                row.Shoperia_W1 += calcHelper;
                                break;
                            case 2:
                                row.Shoperia_W2 += calcHelper;
                                break;
                            case 3:
                                row.Shoperia_W3 += calcHelper;
                                break;
                        }
                        break;
                    case "HOME12":
                        row.Home12AllCompleted += calcHelper;
                        switch (card.Weight)
                        {
                            case 0:
                                row.Home12_UnWeighted += calcHelper;
                                break;
                            case 1:
                                row.Home12_W1 += calcHelper;
                                break;
                            case 2:
                                row.Home12_W2 += calcHelper;
                                break;
                            case 3:
                                row.Home12_W3 += calcHelper;
                                break;
                        }
                        break;
                    case "MATEBIKE":
                        row.MatebikeAllCompleted += calcHelper;
                        switch (card.Weight)
                        {
                            case 0:
                                row.Matebike_UnWeighted += calcHelper;
                                break;
                            case 1:
                                row.Matebike_W1 += calcHelper;
                                break;
                            case 2:
                                row.Matebike_W2 += calcHelper;
                                break;
                            case 3:
                                row.Matebike_W3 += calcHelper;
                                break;
                        }
                        break;
                    case "XPRESS":
                        row.XpressAllCompleted += calcHelper;
                        switch (card.Weight)
                        {
                            case 0:
                                row.Xpress_UnWeighted += calcHelper;
                                break;
                            case 1:
                                row.Xpress_W1 += calcHelper;
                                break;
                            case 2:
                                row.Xpress_W2 += calcHelper;
                                break;
                            case 3:
                                row.Xpress_W3 += calcHelper;
                                break;
                        }
                        break;
                }
                // Az adatbázis változtatások mentése
                try
                {
                    dbContext.SaveChanges();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error az adatbázisba mentés közben: {ex.Message}");
                }
                if (calcHelper == 1)
                {
                    logger.LogInformation("SZÁMLÁLÓ NÖVELVE | bolt: {shop}, év: {year}, hónap: {month}, súly: {weight}, taskId: {id}, taskName: {name}, taskDate: {date}",
                        shop, year, month, card.Weight, card.Id, card.Name, card.Date.ToString("yyyy-MM-dd"));
                }
                else
                {
                    logger.LogInformation("SZÁMLÁLÓ CSÖKKENTVE | bolt: {shop}, év: {year}, hónap: {month}, súly: {weight}, taskId: {id}, taskName: {name}, taskDate: {date}",
                        shop, year, month, card.Weight, card.Id, card.Name, card.Date.ToString("yyyy-MM-dd"));
                }

            }
        }
        /// <summary>
        /// (BEÉPÍTETT SEGÉDFÜGGVÉNY) A paraméterben megkapott hónapban, a kapott list-nek megfelelő SHOP-ban, a kapott labelId alapján frissíti a számlálót.
        /// </summary>
        /// <param name="date">A dátum, ahol frissítjük az értéket.</param>
        /// <param name="listId">A bolt Trello lista ID-ja.</param>
        /// <param name="cardName">A task neve</param>
        /// <param name="weight">A task súlya</param>
        /// <param name="operation">Az operáció típusa: "ADD" a hozzáadáshoz, bármilyen más érték a kivonáshoz.</param>
        public static void UpdateShopCounters(Completed date, string listId, string cardName, int weight, ILogger logger, string operation = "ADD")
        {
            // SHOP LEKÉRDEZÉSE TRELLO LISTA ID ALAPJÁN
            string? shop = GetShopByListId(listId);

            // HA A TRELLO LISTA ID ORDERSHEZ TARTOZIK -> TÖBB SHOP TASKJAI SZEREPELNEK BENNE ->
            // --> KÁRTYA NEVÉBEN SZEREPLŐ AZONOSÍTÓ ALAPJÁN KELL LEKÉRDEZNI A SHOPOT
            if (shop == "ORDERS")
                shop = getShopByCardName(cardName);

            // CALCHELPER === 1     -> KÁRTYA EL LETT FOGADVA   -> SZÁMLÁLÓT NÖVELNI KELL ||
            // CALCHELPER === -1    -> KÁRTYA ÚJRA LETT NYITVA  -> SZÁMLÁLÓT CSÖKKENTENI KELL
            int calcHelper = 1;
            // HA NEM "ADD" AZ OPERÁTOR, AKKOR KIVONNI KELL
            if (operation != "ADD")
                calcHelper = -1;

            // Csak akkor frissíthető az adat, ha valamelyik shop listájához tartozik
            if (shop != null)
            {
                // ÖSSZESÍTETT SZÁMLÁLÓ NÖVELÉSE
                date.AllCompleted += calcHelper;

                // A MEGFELELŐ SHOP MEGFELELŐ SÚLYOZÁSÁNAK FRISSÍTÉSE
                switch (shop)
                {
                    case "SHOPERIA":
                        date.ShoperiaAllCompleted += calcHelper;
                        switch (weight)
                        {
                            case 0:
                                date.Shoperia_UnWeighted += calcHelper;
                                break;
                            case 1:
                                date.Shoperia_W1 += calcHelper;
                                break;
                            case 2:
                                date.Shoperia_W2 += calcHelper;
                                break;
                            case 3:
                                date.Shoperia_W3 += calcHelper;
                                break;
                        }
                        return;
                    case "HOME12":
                        date.Home12AllCompleted += calcHelper;
                        switch (weight)
                        {
                            case 0:
                                date.Home12_UnWeighted += calcHelper;
                                break;
                            case 1:
                                date.Home12_W1 += calcHelper;
                                break;
                            case 2:
                                date.Home12_W2 += calcHelper;
                                break;
                            case 3:
                                date.Home12_W3 += calcHelper;
                                break;
                        }
                        return;
                    case "MATEBIKE":
                        date.MatebikeAllCompleted += calcHelper;
                        switch (weight)
                        {
                            case 0:
                                date.Matebike_UnWeighted += calcHelper;
                                break;
                            case 1:
                                date.Matebike_W1 += calcHelper;
                                break;
                            case 2:
                                date.Matebike_W2 += calcHelper;
                                break;
                            case 3:
                                date.Matebike_W3 += calcHelper;
                                break;
                        }
                        return;
                    case "XPRESS":
                        date.XpressAllCompleted += calcHelper;
                        switch (weight)
                        {
                            case 0:
                                date.Xpress_UnWeighted += calcHelper;
                                break;
                            case 1:
                                date.Xpress_W1 += calcHelper;
                                break;
                            case 2:
                                date.Xpress_W2 += calcHelper;
                                break;
                            case 3:
                                date.Xpress_W3 += calcHelper;
                                break;
                        }
                        return;
                }
                if (calcHelper == 1)
                {
                    logger.LogInformation("SZÁMLÁLÓ NÖVELVE | bolt: {shop}, év: {year}, hónap: {month}, súly: {weight}, taskName: {name}",
                        shop, date.Date.Year, date.Date.Month, weight, cardName);
                }
                else
                {
                    logger.LogInformation("SZÁMLÁLÓ CSÖKKENTVE | bolt: {shop}, év: {year}, hónap: {month}, súly: {weight}, taskName: {name}",
                        shop, date.Date.Year, date.Date.Month, weight, cardName);
                }
            }
        }
        /// <summary>
        /// Kap egy lista Id-t paraméterül, és megkeresi, hogy az melyik shophoz tartozik, ha egyikhez se, akkor nullal tér vissza
        /// </summary>
        /// <param name="list">Trello lista ID.</param>
        public static string? GetShopByListId(string list)
        {
            var listIDs = Settings.GetLists();
            if (listIDs.Shoperia != null && listIDs.Home12 != null && listIDs.Xpress != null && listIDs.Matebike != null)
            {
                foreach (string listID in listIDs.Shoperia)
                    if (listID == list)
                        return "SHOPERIA";
                foreach (string listID in listIDs.Home12)
                    if (listID == list)
                        return "HOME12";
                foreach (string listID in listIDs.Xpress)
                    if (listID == list)
                        return "XPRESS";
                foreach (string listID in listIDs.Matebike)
                    if (listID == list)
                        return "MATEBIKE";
                if (list == listIDs.Orders)
                    return "ORDERS";
            }
            return null;
        }
        /// <summary>
        /// Az alábbi függvény az ORDERS listának a segédfüggvénye, ahol a task nevének első szakasza azonosítja a boltot. A függvény megkeresi, hogy melyik shophoz tartozik az adott kártyanév, ha egyikhez se, akkor nullal tér vissza
        /// </summary>
        /// <param name="cardName">Task neve.</param>
        static string? getShopByCardName(string cardName)
        {
            string shopId = cardName.Split('-')[0];
            switch (shopId)
            {
                case "SH":
                    return "SHOPERIA";
                case "XP":
                    return "XPRESS";
                case "HOME12":
                    return "HOME12";
                case "MATE":
                    return "MATEBIKE";
            }
            return null;
        }

        /// <summary>
        /// A paraméterben megadott labelek alapján visszaadja, hogy milyen a projekt fontossági súlya
        /// </summary>
        /// <param name="labels">Task neve.</param>
        public static int GetWeightFromLabels(List<string> labels)
        {
            var weightedLabels = Settings.GetLabels();
            int tempMax = 0;
            foreach (string label in labels)
            {
                if (label == weightedLabels.Weight3)
                    return 3;
                else if (label == weightedLabels.Weight2)
                    tempMax = 2;
                else if (tempMax < 2 && label == weightedLabels.Weight1)
                    tempMax = 1;
            }
            return tempMax;

        }
        public static void RefreshCompletedTableByStoredTasks(ApplicationDbContext dbContext, ILogger logger) 
        {
            
            try
            {
                // ÖSSZES TASK LEKÉRDEZÉSE AZ ADATBÁZISBÓL
                var tasksInDB = dbContext.Cards ?? throw new Exception("Hiba történt a taskok lekérdezése közben az adatbázisból (null értékkel tért vissza)");
                foreach (var task in tasksInDB)
                {
                    if (task.IsComplete == true)
                    {

                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}

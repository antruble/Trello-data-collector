using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Trello
{
    public class ExcelColorAssist
    {
        public System.Drawing.Color Title;
        public System.Drawing.Color Default;
        public System.Drawing.Color W1;
        public System.Drawing.Color W2;
        public System.Drawing.Color W3;
    }
    public class ExcelColorList 
    {
        public List<ExcelColorAssist> ShopColors;
        //DÁTUM SZÍN
        public System.Drawing.Color DateColor = System.Drawing.Color.FromArgb(96, 229, 252);
        // SUMMARY SZÍN
        public System.Drawing.Color SummaryColor = System.Drawing.Color.FromArgb(54, 145, 163);
        public ExcelColorList() 
        {
            ShopColors = new List<ExcelColorAssist>()
            {
                //SHOPERIA
                new ExcelColorAssist
                {
                    Title = System.Drawing.Color.FromArgb(90, 0, 128), // Sötét lila
                    Default = System.Drawing.Color.FromArgb(204, 153, 255), // Világos lila
                    W1 = System.Drawing.Color.FromArgb(153, 102, 204), // Közepesen világos lila
                    W2 = System.Drawing.Color.FromArgb(102, 51, 153), // Közepesen sötét lila
                    W3 = System.Drawing.Color.FromArgb(73, 35, 112) // Sötét lila
                },
                //HOM12
                new ExcelColorAssist
                {
                    Title = System.Drawing.Color.FromArgb(0, 128, 0), // Sötét zöld
                    Default = System.Drawing.Color.FromArgb(144, 238, 144), // Világos zöld
                    W1 = System.Drawing.Color.FromArgb(34, 139, 34), // Közepesen sötét zöld
                    W2 = System.Drawing.Color.FromArgb(0, 128, 0), // Közepesen világos zöld
                    W3 = System.Drawing.Color.FromArgb(0, 100, 0) // Sötét zöld
                },
                //XPRESS
                new ExcelColorAssist
                {
                    Title = System.Drawing.Color.FromArgb(255, 140, 0), // Sötét narancssárga
                    Default = System.Drawing.Color.FromArgb(255, 215, 0), // Világos narancssárga
                    W1 = System.Drawing.Color.FromArgb(255, 165, 0), // Közepesen világos narancssárga
                    W2 = System.Drawing.Color.FromArgb(255, 140, 0), // Közepesen sötét narancssárga
                    W3 = System.Drawing.Color.FromArgb(255, 69, 0) // Sötét narancssárga
                },
                //MATEBIKE
                new ExcelColorAssist
                {
                    Title = System.Drawing.Color.FromArgb(105, 105, 105), // Sötét szürke
                    Default = System.Drawing.Color.FromArgb(192, 192, 192), // Világos szürke
                    W1 = System.Drawing.Color.FromArgb(169, 169, 169), // Közepesen világos szürke
                    W2 = System.Drawing.Color.FromArgb(128, 128, 128), // Közepesen sötét szürke
                    W3 = System.Drawing.Color.FromArgb(105, 105, 105) // Sötét szürke
                },
            };
        }
    }
}

using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TrelloDotNet.AutomationEngine.Model;

namespace Trello
{
    public class Settings
    {
        public string? ConnectionString { get; set; }
        public string? ApiKey { get; set; }
        public string? UserSecret { get; set; }
        public string? BoardId { get; set; }
        public static Settings LoadSettings()
        {
            // Load settings
            IConfiguration config = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: false)
                .Build();

            return config.GetRequiredSection("Settings").Get<Settings>() ??
                throw new Exception("Could not load app settings.");
        }
        public static Label GetLabels()
        {
            // Load settings
            IConfiguration config = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: false)
                .Build();

            return config.GetRequiredSection("Labels").Get<Label>() ??
                throw new Exception("Could not load labels.");
        }
        public static Lists GetLists()
        {
            // Load settings
            IConfiguration config = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: false)
                .Build();

            return config.GetRequiredSection("shop-list").Get<Lists>() ??
                throw new Exception("Could not load lists.");
        }
        public static ExcelData GetSheets()
        {
            // Load settings
            IConfiguration config = new ConfigurationBuilder()
                .AddJsonFile("appsettings.json", optional: false)
                .Build();

            return config.GetRequiredSection("excel-data").Get<ExcelData>() ??
                throw new Exception("Could not load lists.");
        }
    }
    public class Label 
    {
        public string? Weight1 { get; set; }
        public string? Weight2 { get; set; }
        public string? Weight3 { get; set; }
    }
    public class Lists
    {
        public List<string>? Shoperia { get; set; }
        public List<string>? Xpress { get; set; }
        public List<string>? Home12 { get; set; }
        public List<string>? Matebike { get; set; }
        public string? Orders {  get; set; }
    }
    public class ExcelData 
    {
        public string? Summary { get; set; }
        public string? Cards { get; set; }
    }
}

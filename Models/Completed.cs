using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Trello.Models
{
    public class Completed
    {
        [Key]
        public DateTime Date { get; set; }
        public int AllCompleted { get; set; }
        public int ShoperiaAllCompleted { get; set; }
        public int Shoperia_W1 { get; set; }
        public int Shoperia_W2 { get; set; }
        public int Shoperia_W3 { get; set; }
        public int Shoperia_UnWeighted { get; set; }
        public int XpressAllCompleted { get; set; }
        public int Xpress_W1 { get; set; }
        public int Xpress_W2 { get; set; }
        public int Xpress_W3 { get; set; }
        public int Xpress_UnWeighted { get; set; }
        public int Home12AllCompleted { get; set; }
        public int Home12_W1 { get; set; }
        public int Home12_W2 { get; set; }
        public int Home12_W3 { get; set; }
        public int Home12_UnWeighted { get; set; }
        public int MatebikeAllCompleted { get; set; }
        public int Matebike_W1 { get; set; }
        public int Matebike_W2 { get; set; }
        public int Matebike_W3 { get; set; }
        public int Matebike_UnWeighted { get; set; }

    }
}

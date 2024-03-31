using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Trello.Models
{
    [Table(name: "Cards")]
    public class CardModel
    {
        [Key]
        public string? Id { get; set; }
        public string? Name { get; set; }
        public DateTime Date { get; set; }
        public int Weight { get; set; }
        public virtual ListModel? List {  get; set; } 
        public string? ListId {  get; set; } 
        public bool? IsComplete { get; set; }
    }
}

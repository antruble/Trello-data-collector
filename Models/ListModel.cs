using Manatee.Trello;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Trello.Models
{
    public class ListModel
    {
        public string? Id { get; set; }
        public string? Name { get; set; }
        public virtual ICollection<CardModel>? Cards { get; set; } 
    }
}

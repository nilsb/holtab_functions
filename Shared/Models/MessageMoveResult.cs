using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shared.Models
{
    public class MessageMoveResult
    {
        public bool Success { get; set; } = false;
        public string MessageID { get; set; } = "";
        public List<string> Files { get; set; } = new List<string>();
    }
}

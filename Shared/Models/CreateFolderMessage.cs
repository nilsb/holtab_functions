using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shared.Models
{
    public class CreateFolderMessage
    {
        public string? GroupID { get; set; }
        public string? GroupURL { get; set; }
    }
}

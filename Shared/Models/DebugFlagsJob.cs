using System;
using System.Collections.Generic;
using System.Diagnostics.SymbolStore;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shared.Models
{
    public class DebugFlagsJob
    {
        public bool BGUnhandledOrders { get; set; }
        public bool CreateOrUpdateAMKundrekl { get; set; }
        public bool HandleEmail { get; set; }
        public bool PostProcessEmails { get; set; }
        public bool PostProcessOrders { get; set; }
    }
}

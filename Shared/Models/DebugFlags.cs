using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shared.Models
{
    public class DebugFlags
    {
        public DebugFlagsCustomer? Customer { get; set; }
        public DebugFlagsOrder? Order { get; set; }
        public DebugFlagsJob? Job { get; set; }
    }
}

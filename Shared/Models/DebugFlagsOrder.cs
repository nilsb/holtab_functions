using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shared.Models
{
    public class DebugFlagsOrder
    {
        public bool BGOrderInfo { get; set; }
        public bool BGOrderGroupTeamFolder { get; set; }
        public bool BGCreateProject { get; set; }
        public bool BGCopyOrderStructure { get; set; }
        public bool BGAssignPermission { get; set; }
    }
}

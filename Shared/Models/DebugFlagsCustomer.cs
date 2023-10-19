using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shared.Models
{
    public class DebugFlagsCustomer
    {
        public bool BGCustomerInfo { get; set; }
        public bool BGCreateTeam { get; set; }
        public bool BGGreateGroup { get; set; }
        public bool BGCreateColumn { get; set; }
        public bool BGCopyRootStructure { get; set; }
        public bool BGAssignPermissions { get; set; }
    }
}

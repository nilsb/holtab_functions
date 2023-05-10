using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace Shared.Models
{
    public class FindCustomerGroupResult
    {
        public bool Success { get; set; }
        public Group? group { get; set; }
        public Drive? groupDrive { get; set; }
        public DriveItem? rootFolder { get; set; }
        public DriveItem? generalFolder { get; set; }
        public List<DriveItem>? rootItems { get; set; }
        public Customer? customer { get; set; }
    }
}

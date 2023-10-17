using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace Shared.Models
{
    public class FindOrderGroupAndFolder
    {
        public bool Success { get; set; } = false;
        public string? orderTeamId { get; set; }
        public string? orderGroupId { get; set; }
        public string? orderDriveId { get; set; }
        public DriveItem? orderFolder { get; set; }
        public DriveItem? generalFolder { get; set; }
        public Customer? customer { get; set; }
    }
}

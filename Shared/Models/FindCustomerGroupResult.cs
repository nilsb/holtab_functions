using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace Shared.Models
{
    public class FindCustomerGroupResult
    {
        public bool Success { get; set; } = false;
        public string? groupId { get; set; }
        public string? groupDriveId { get; set; }
        public DriveItem? rootFolder { get; set; }
        public string? generalFolderId { get; set; }
        public List<DriveItem>? rootItems { get; set; }
        public Customer? customer { get; set; }
    }
}

using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace Shared.Models
{
    public class FindGroupResult
    {
        public bool Success { get; set; }
        public int Count { get; set; }
        public Group? group { get; set; }
        public List<Group>? groups { get; set; }
    }
}

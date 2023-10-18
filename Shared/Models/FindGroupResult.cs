using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace Shared.Models
{
    public class FindGroupResult
    {
        public bool Success { get; set; } = false;
        public int Count { get; set; } = 0;
        public string? group { get; set; }
        public List<string?>? groups { get; set; }
    }
}

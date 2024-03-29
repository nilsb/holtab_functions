﻿using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace Shared.Models
{
    public class CreateCustomerResult
    {
        public bool Success { get; set; }
        public string? group { get; set; }
        public Team? team { get; set; }
        public Customer? customer { get; set; }
    }
}

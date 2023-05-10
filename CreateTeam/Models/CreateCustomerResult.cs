using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace CreateTeam.Models
{
    public class CreateCustomerResult
    {
        public bool Success { get; set; }
        public Group group { get; set; }
        public Team team { get; set; }
        public Customer customer { get; set; }
    }
}

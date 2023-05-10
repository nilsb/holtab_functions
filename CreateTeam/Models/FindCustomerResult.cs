using System;
using System.Collections.Generic;
using System.Text;

namespace CreateTeam.Models
{
    public class FindCustomerResult
    {
        public FindCustomerResult()
        {
            Success = false;
            customers = new List<Customer>();
            customer = null;
        }

        public bool Success { get; set; }
        public List<Customer> customers { get; set; }
        public Customer customer { get; set; }
    }
}

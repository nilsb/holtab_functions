using System;
using System.Collections.Generic;
using System.Text;

namespace CreateTeam.Models
{
    public class Company
    {
        public Guid ID { get; set; }
        public string ExternalId { get; set; } = "";
        public string Name { get; set; } = "";
        public string Address { get; set; } = "";
        public string Address1 { get; set; } = "";
        public string ZipCode { get; set; } = "";
        public string City { get; set; } = "";
        public string State { get; set; } = "";
        public string Country { get; set; } = "";
        public string Phone { get; set; } = "";
        public string Fax { get; set; } = "";
        public DateTime Created { get; set; }
    }
}

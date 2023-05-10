using System;
using System.Collections.Generic;
using System.Text;

namespace CreateTeam.Models
{
    public class Contact
    {
        public Guid ID { get; set; }
        public string ExternalId { get; set; } = "";
        public Customer Customer { get; set; }
        public string Name { get; set; } = "";
        public string Email { get; set; } = "";
        public string Direct { get; set; } = "";
        public string Mobile { get; set; } = "";
        public string JobTitle { get; set; } = "";
        public Company Company { get; set; }
        public DateTime Created { get; set; }
    }
}

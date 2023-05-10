using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace CreateTeam.Models
{
    public class CustomerQueueMessage
    {
        public string ID { get; set; }
        public string ExternalId { get; set; }
        public string Type { get; set; }
        public string Name { get; set; }
    }
}

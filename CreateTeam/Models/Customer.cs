using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;
using System.Text.Json.Serialization;

namespace CreateTeam.Models
{
    public class Customer
    {
        [Key]
        public Guid ID { get; set; }
        [JsonPropertyName("CustomerNo")]
        public string ExternalId { get; set; } = "";
        public string Type { get; set; } = "";
        [JsonPropertyName("Responsible")]
        public string Seller { get; set; } = "";
        public string ProjectManager { get; set; } = "";
        [JsonPropertyName("CustomerName")]
        public string Name { get; set; } = "";
        [JsonPropertyName("ProspectNo")]
        public string Prospect { get; set; } = "";
        [JsonPropertyName("CustomerAddress")]
        public string Address { get; set; } = "";
        [JsonPropertyName("CustomerAddress1")]
        public string Address1 { get; set; } = "";
        [JsonPropertyName("CustomerZipCode")]
        public string ZipCode { get; set; } = "";
        [JsonPropertyName("CustomerCity")]
        public string City { get; set; } = "";
        [JsonPropertyName("CustomerState")]
        public string State { get; set; } = "";
        [JsonPropertyName("CustomerCountry")]
        public string Country { get; set; } = "";
        [JsonPropertyName("CustomerPhone")]
        public string Phone { get; set; } = "";
        [JsonPropertyName("CustomerFax")]
        public string Fax { get; set; } = "";
        public string TeamID { get; set; } = "";
        public string TeamUrl { get; set; } = "";
        public string GroupID { get; set; } = "";
        public string GroupURL { get; set; } = "";
        public string DriveID { get; set; } = "";
        public string GeneralFolderID { get; set; }
        public bool GroupCreated { get; set; } = false;
        public bool TeamCreated { get; set; } = false;
        public bool GeneralFolderCreated { get; set; } = false;
        public bool CopiedRootStructure { get; set; } = false;
        public bool CreatedColumnKundnummer { get; set; } = false;
        public bool CreatedColumnAdditionalInfo { get; set; } = false;
        public bool CreatedColumnNAVid { get; set; } = false;
        public bool CreatedColumnProduktionsdokument { get; set; } = false;
        public bool CreatedDefaultView { get; set; } = false;
        public bool InstalledApp { get; set; } = false;
        public string MembersAdded { get; set; } = "";
        public DateTime Created { get; set; }
        public DateTime Modified { get; set; }
        [NotMapped]
        [Newtonsoft.Json.JsonIgnore]
        public virtual ICollection<Order> Orders { get; set; }
        public string itemId { get; set; }
    }
}

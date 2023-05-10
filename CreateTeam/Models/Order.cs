using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;
using System.Text.Json.Serialization;

namespace CreateTeam.Models
{
    public class Order
    {
        [Key]
        public Guid ID { get; set; }
        [NotMapped]
        public string No { get; set; } = "";
        public string ExternalId { get; set; } = "";
        public Guid CustomerID { get; set; }
        [NotMapped]
        public Customer Customer { get; set; }
        public string Seller { get; set; } = "";
        public string ProjectManager { get; set; } = "";
        public string AdditionalInfo { get; set; } = "";
        [NotMapped]
        public string DriveID { get; set; }
        public string FolderID { get; set; }
        public bool GroupFound { get; set; }
        public bool DriveFound { get; set; }
        public bool GeneralFolderFound { get; set; }
        public bool OrdersFolderFound { get; set; }
        public bool OffersFolderFound { get; set; }
        public bool PurchaseFolderFound { get; set; }
        public bool CreatedFolder { get; set; }
        public bool StructureCreated { get; set; }
        public bool Handled { get; set; } = false;
        public string Status { get; set; }
        public DateTime Created { get; set; }
        public string Type { get; set; }
        [NotMapped]
        public string itemId { get; set; }
        [NotMapped]
        public string CustomerName { get; set; }
        public string CustomerNo { get; set; }
        public string CustomerType { get; set; }
    }
}

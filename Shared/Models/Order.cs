using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Text;
using System.Text.Json.Serialization;

namespace Shared.Models
{
    public class Order
    {
        public Order()
        {
        }

        public Order(Order? src)
        {
            if(src != null)
            {
                this.ID = src.ID;
                this.No = src.No;
                this.ExternalId = src.ExternalId;
                this.Seller = src.Seller;
                this.ProjectManager = src.ProjectManager;
                this.AdditionalInfo = src.AdditionalInfo;
                this.DriveID = src.DriveID;
                this.FolderID = src.FolderID;
                this.GroupFound = src.GroupFound;
                this.DriveFound = src.DriveFound;
                this.GeneralFolderFound = src.GeneralFolderFound;
                this.OrdersFolderFound = src.OrdersFolderFound;
                this.OffersFolderFound = src.OffersFolderFound;
                this.PurchaseFolderFound = src.PurchaseFolderFound;
                this.CreatedFolder = src.CreatedFolder;
                this.StructureCreated = src.StructureCreated;
                this.Handled = src.Handled;
                this.Status = src.Status;
                this.Created = src.Created;
                this.Type = src.Type;
                this.itemId = src.itemId;
                this.QueueCount = src.QueueCount;

                if (src.Customer != null)
                {
                    this.Customer = src.Customer;
                    this.CustomerID = src.Customer.ID;
                    this.CustomerName = src.Customer.Name;
                    this.CustomerNo = src.Customer.ExternalId;
                    this.CustomerType = src.Customer.Type;
                }
            }
        }

        [Key]
        public Guid ID { get; set; }
        [NotMapped]
        public string No { get; set; } = "";
        public string ExternalId { get; set; } = "";
        public Guid CustomerID { get; set; }
        [NotMapped]
        public Customer? Customer { get; set; }
        public string Seller { get; set; } = "";
        public string ProjectManager { get; set; } = "";
        public string AdditionalInfo { get; set; } = "";
        [NotMapped]
        public string DriveID { get; set; } = "";
        public string FolderID { get; set; } = "";
        public bool GroupFound { get; set; }
        public bool DriveFound { get; set; }
        public bool GeneralFolderFound { get; set; }
        public bool OrdersFolderFound { get; set; }
        public bool OffersFolderFound { get; set; }
        public bool PurchaseFolderFound { get; set; }
        public bool CreatedFolder { get; set; }
        public bool StructureCreated { get; set; }
        public bool Handled { get; set; } = false;
        public string Status { get; set; } = "";
        public DateTime Created { get; set; }
        public string Type { get; set; } = "";
        [NotMapped]
        public string itemId { get; set; } = "";
        [NotMapped]
        public string CustomerName { get; set; } = "";
        public string CustomerNo { get; set; } = "";
        public string CustomerType { get; set; } = "";
        public int QueueCount { get; set; } = 0;
    }
}

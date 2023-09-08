using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph.Models;

namespace Shared.Models
{
    public class OrderMessage
    {
        public string No { get; set; } = "";
        public string Type { get; set; } = "";
        public string Seller { get; set; } = "";
        public string ProjectManager { get; set; } = "";
        public string AdditionalInfo { get; set; } = "";
        public string CustomerName { get; set; } = "";
        public string CustomerNo { get; set; } = "";
        public string CustomerType { get; set; } = "";
        public string ExternalId { get; set; } = "";
        public string CustomerExternalId { get; set; } = "";
        public string CustomerGroupID { get; set; } = "";
        public string DriveID { get; set; } = "";
        public string GeneralFolderID { get; set; } = "";
        public string TeamID { get; set; } = "";
        public string OrderParentFolderID { get; set; } = "";
        public string OrderFolderID { get; set; } = "";
        public bool NeedStructureCopy { get; set; } = false;
    }
}

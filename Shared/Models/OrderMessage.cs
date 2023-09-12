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
        public OrderMessage()
        {
            
        }

        public OrderMessage(dynamic src)
        {
            if(src != null)
            {
                try
                {
                    this.No = src.No;
                }
                catch
                {
                }

                try
                {
                    this.Type = src.Type;
                }
                catch
                {
                }

                try
                {
                    this.Seller = src.Seller;
                }
                catch
                {
                }

                try
                {
                    this.ProjectManager = src.ProjectManager;
                }
                catch
                {
                }

                try
                {
                    this.AdditionalInfo = src.AdditionalInfo;
                }
                catch
                {
                }

                try
                {
                    this.CustomerName = src.CustomerName;
                }
                catch
                {
                }

                try
                {
                    this.CustomerNo = src.CustomerNo;
                }
                catch
                {
                }

                try
                {
                    this.CustomerType = src.CustomerType;
                }
                catch
                {
                }

                try
                {
                    this.ExternalId = src.ExternalId;
                }
                catch
                {
                }

                try
                {
                    this.CustomerExternalId = src.CustomerExternalId;
                }
                catch
                {
                }

                try
                {
                    this.CustomerGroupID = src.CustomerGroupID;
                }
                catch
                {
                }

                try
                {
                    this.DriveID = src.DriveID;
                }
                catch
                {
                }

                try
                {
                    this.GeneralFolderID = src.GeneralFolderID;
                }
                catch
                {
                }

                try
                {
                    this.TeamID = src.TeamID;
                }
                catch
                {
                }

                try
                {
                    this.OrderParentFolderID = src.OrderParentFolderID;
                }
                catch
                {
                }

                try
                {
                    this.OrderFolderID = src.OrderFolderID;
                }
                catch
                {
                }

                try
                {
                    this.NeedStructureCopy = src.NeedStructureCopy;
                }
                catch
                {
                }

            }
        }

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

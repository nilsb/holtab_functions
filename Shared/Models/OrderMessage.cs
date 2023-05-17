using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Shared.Models
{
    public class OrderMessage
    {
        public string? No { get; set; } = "";
        public string? Seller { get; set; } = "";
        public string? ProjectManager { get; set; } = "";
        public string? AdditionalInfo { get; set; } = "";
        public string? CustomerName { get; set; } = "";
        public string? CustomerNo { get; set; } = "";
        public string? CustomerType { get; set; } = "";

    }
}

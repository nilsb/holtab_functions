using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace CreateTeam.Models
{
    public class OrderFiles
    {
        public DriveItem file { get; set; }
        public List<DriveItem> associated { get; set; }
    }
}

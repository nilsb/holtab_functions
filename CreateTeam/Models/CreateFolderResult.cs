﻿using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace CreateTeam.Models
{
    public class CreateFolderResult
    {
        public DriveItem folder { get; set; }
        public bool Success { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Text;

namespace Shared.Models
{
    public class HandleEmailMessage
    {
        public string Title { get; set; } = "";
        public string Filename { get; set; } = "";
        public string Source { get; set; } = "";
        public string Sender { get; set; } = "";
    }
}

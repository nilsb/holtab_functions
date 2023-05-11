using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace Shared.Models
{
    public class MetadataDefaults
    {
        [XmlArray("a")]
        public List<Reference>? Refs { get; set; }
    }

    public class Reference
    {
        [XmlAttribute]
        public string? href { get; set; }
        [XmlArray("DefaultValue")]
        public List<DefaultValue>? DefaultValues { get; set; }
    }

    public class DefaultValue
    {
        [XmlAttribute]
        public string? FieldName { get; set; }
        [XmlText]
        public string? Value { get; set; }
    }
}

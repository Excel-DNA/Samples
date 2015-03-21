using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Serialization;

namespace AddInReloader
{
    // ---------------------------------------------------------------------------------------------------
    // Configuration types

    [Serializable]
    [XmlType(AnonymousType = true)]
    [XmlRoot(Namespace = "", IsNullable = false)]
    public class AddInReloaderConfiguration
    {
        [XmlElement("WatchedAddIn", typeof(WatchedAddIn))]
        public List<WatchedAddIn> WatchedAddIns { get; set; }
    }

    [Serializable]
    public class WatchedAddIn
    {
        [XmlAttribute]
        public string Path { get; set; }
        [XmlElement("WatchedFile", typeof(WatchedFile))]
        public List<WatchedFile> WatchedFiles { get; set; }
    }

    [Serializable]
    public class WatchedFile
    {
        [XmlAttribute]
        public string Path { get; set; }
    }
}

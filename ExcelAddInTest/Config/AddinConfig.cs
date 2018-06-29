using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddInTest.Config
{
    [DataContract]
    public class AddinConfig
    {
        [DataMember]
        public bool IsMark { get; set; }
    }
}

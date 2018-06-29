using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;

namespace TestOtherFile.Config
{
    [DataContract]
    public class TestConfig
    {
        [DataMember]
        public int IsMark { get; set; }
    }
}

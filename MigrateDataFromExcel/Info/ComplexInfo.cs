using MigrateDataFromExcel.Attribute;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MigrateDataFromExcel.Info
{
    public class ComplexInfo
    {
        [MigrateProperty]
        public int IntProperty { get; set; }

        [MigrateProperty(IsRequired = true, IsManualComposed = true)]
        public double DoubleProperty { get; set; }

        public long LongProperty { get; set; }

        public string StringProperty { get; set; }
    }
}

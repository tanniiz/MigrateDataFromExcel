using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MigrateDataFromExcel.Attribute
{
    public class MigrateProperty : System.Attribute
    {
        public string ColumnName { get; set; }

        public int ColumnIndex { get; set; }

        public bool IsRequired { get; set; }

        public bool IsManualComposed { get; set; }
    }
}

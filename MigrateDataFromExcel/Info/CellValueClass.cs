using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MigrateDataFromExcel.Info
{
    public delegate bool ValidateRule(object cellValue);

    public class CellValueClass
    {
        public string PropertyName { get; set; }

        public int ColumnIndex { get; set; }

        public string ColumnName { get; set; }

        public Type Type { get; set; }

        public bool IsRequire { get; set; }

        public bool IsManualComposed { get; set; }
    }
}

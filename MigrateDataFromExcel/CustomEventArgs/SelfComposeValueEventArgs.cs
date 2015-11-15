using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MigrateDataFromExcel.CustomEventArgs
{
    public class SelfComposeValueEventArgs
    {
        public string PropertyName { get; set; }

        public string CellValue { get; set; }

        public string ColumnName { get; set; }

        public int ColumnIndex { get; set; }

        public int RowIndex { get; set; }
    }
}

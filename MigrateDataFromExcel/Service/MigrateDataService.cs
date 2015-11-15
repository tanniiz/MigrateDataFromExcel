using MigrateDataFromExcel.Attribute;
using MigrateDataFromExcel.CustomEventArgs;
using MigrateDataFromExcel.Info;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MigrateDataFromExcel.Service
{
    public static class MigrateDataService
    {
        public static event EventHandler<SelfComposeValueEventArgs> OnSelfComposedProperty;

        public delegate bool AfterComposedInfoHandler(object sender, AfterComposedInfoEventArgs args);
        public static event AfterComposedInfoHandler AfterComposedInfo;

        public static ISheet GetSheet(string filePath)
        {
            using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                HSSFWorkbook workbook = new HSSFWorkbook(file);
                ISheet sheet = workbook.GetSheetAt(0);

                return sheet;
            }
        }

        public static List<PropertyInfo> GetPropertyListFromInfo(Type type)
        {
            var properties = type.GetProperties().Where(x => x.IsDefined(typeof(MigrateProperty)));

            return properties.ToList();
        }

        public static List<CellValueClass> GetColumnCellInfoFromProperties(
            List<PropertyInfo> properties,
            ISheet sheet,
            int columnRow = 0 
            )
        {
            if (properties == null)
                return null;

            Dictionary<string, int> columnNameToColumnIndexDict = new Dictionary<string, int>();
            IRow row = null;

            if (sheet != null)
                row = sheet.GetRow(columnRow);

            if (row != null)
            {
                foreach (var columnCell in row.Cells)
                {
                    try
                    {
                        columnNameToColumnIndexDict.Add(columnCell.StringCellValue, columnCell.ColumnIndex);
                    }
                    catch(Exception ex)
                    {
                        throw ex;
                    }
                }
            }

            CellValueClass cellObj;
            List<CellValueClass> cells = new List<CellValueClass>();

            foreach(var prop in properties)
            {
                cellObj = new CellValueClass();

                var propertyAttribute = prop.GetCustomAttribute(typeof(MigrateProperty)) as MigrateProperty;

                if (propertyAttribute != null)
                {
                    if(!String.IsNullOrEmpty(propertyAttribute.ColumnName))
                    {
                        cellObj.ColumnName = propertyAttribute.ColumnName;
                    }
                    else
                    {
                        if (columnNameToColumnIndexDict.ContainsKey(prop.Name))
                        {
                            cellObj.ColumnName = prop.Name;
                        }
                    }

                    if(propertyAttribute.ColumnIndex > 0)
                    {
                        cellObj.ColumnIndex = propertyAttribute.ColumnIndex;
                    }
                    else
                    {
                        if (columnNameToColumnIndexDict.ContainsKey(prop.Name))
                        {
                            cellObj.ColumnIndex = columnNameToColumnIndexDict[prop.Name];
                        }
                    }

                    cellObj.PropertyName = prop.Name;
                    cellObj.IsRequire = propertyAttribute.IsRequired;
                    cellObj.IsManualComposed = propertyAttribute.IsManualComposed;
                    cellObj.Type = prop.PropertyType;

                    cells.Add(cellObj);
                }
            }

            return cells;
        }

        public static Dictionary<bool, List<TInfo>> GetInfoes<TInfo>(ISheet sheet,  
            List<CellValueClass> columns,
            Dictionary<string, ValidateRule> validateDict = null, 
            int startRowIndex = 0, 
            int? endRowIndex = null) where TInfo : class, new()
        {
            if(endRowIndex <= startRowIndex)
                throw new ApplicationException("endRowIndex must greater than startRowIndex.");

            Type type = typeof(TInfo);

            if(endRowIndex == null)
                endRowIndex = sheet.LastRowNum;

            Dictionary<bool, List<TInfo>> infoes = new Dictionary<bool, List<TInfo>>();
            infoes[true] = new List<TInfo>();
            infoes[false] = new List<TInfo>();

            bool isRowValid = true;
            IRow row;

            for (int i = startRowIndex; i <= endRowIndex; i++ )
            {
                row = sheet.GetRow(i);
                isRowValid = true;
                TInfo info = new TInfo();

                foreach(var item in columns)
                {
                    try
                    {
                        if(item.IsManualComposed)
                        {
                            SelfComposeValueEventArgs args = new SelfComposeValueEventArgs
                            {
                                PropertyName = item.PropertyName,
                                ColumnIndex = item.ColumnIndex,
                                ColumnName = item.ColumnName,
                                RowIndex = i,
                                CellValue = GetCellValue<string>(row.GetCell(item.ColumnIndex))
                            };

                            OnSelfComposedProperty(info, args);
                        }

                        var prop = type.GetProperty(item.ColumnName);

                        if (prop != null)
                        {
                            var cellValue = GetCellValue<string>(row.GetCell(item.ColumnIndex));

                            if(item.IsRequire)
                            {
                                if(String.IsNullOrEmpty(cellValue))
                                {
                                    isRowValid = false;
                                }
                            }

                            if (item.Type == typeof(int) || item.Type == typeof(int?))
                            {
                                prop.SetValue(info, GetCellValue<int>(row.GetCell(item.ColumnIndex)));
                            }
                            else if (item.Type == typeof(double) || item.Type == typeof(double?))
                            {
                                prop.SetValue(info, GetCellValue<double>(row.GetCell(item.ColumnIndex)));
                            }
                            else if (item.Type == typeof(long) || item.Type == typeof(long?))
                            {
                                prop.SetValue(info, GetCellValue<long>(row.GetCell(item.ColumnIndex)));
                            }
                            else if (item.Type == typeof(string))
                            {
                                prop.SetValue(info, GetCellValue<string>(row.GetCell(item.ColumnIndex)));
                            }
                        }

                        if (validateDict != null)
                        {
                            if (validateDict.ContainsKey(item.ColumnName))
                            {
                                isRowValid = validateDict[item.ColumnName](prop.GetValue(info));
                            }
                        }
                    }
                    catch(Exception ex)
                    {
                        isRowValid = false;
                    }
                }

                if (AfterComposedInfo != null)
                {
                    isRowValid = AfterComposedInfo(null, new AfterComposedInfoEventArgs { Info = info });
                }
                
                infoes[isRowValid].Add(info);
            }

            return infoes;
        }

        private static T GetCellValue<T>(ICell cell)
        {
            if (cell == null && typeof(T) != typeof(string))
                throw new ApplicationException("Cell cannot be null.");

            if (cell == null && typeof(T) == typeof(string))
                return (T)(object)String.Empty;

            string cellValue = null;

            try
            {
                cellValue = cell.StringCellValue;
            }
            catch
            {
                cellValue = cell.NumericCellValue.ToString();
            }

            try
            {
                if (typeof(T) == typeof(int))
                {
                    return (T)(object)Convert.ToInt32(cellValue);
                }
                else if (typeof(T) == typeof(int?))
                {
                    int intValue;
                    return (T)(object)(Int32.TryParse(cellValue, out intValue) ? intValue : (int?)null);
                }
                else if (typeof(T) == typeof(double))
                {
                    return (T)(object)Convert.ToDouble(cellValue);
                }
                else if (typeof(T) == typeof(double?))
                {
                    double doubleValue;
                    return (T)(object)(Double.TryParse(cellValue, out doubleValue) ? doubleValue : (double?)null);
                }
                else if (typeof(T) == typeof(long))
                {
                    return (T)(object)Convert.ToInt64(cellValue);
                }
                else if (typeof(T) == typeof(long?))
                {
                    long longValue;
                    return (T)(object)(Int64.TryParse(cellValue, out longValue) ? longValue : (long?)null);
                }
                else if (typeof(T) == typeof(string))
                {
                    return (T)(object)cellValue;
                }
                else
                {
                    return default(T);
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
    }
}

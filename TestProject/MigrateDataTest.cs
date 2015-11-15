using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using MigrateDataFromExcel.Service;
using MigrateDataFromExcel.Info;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using MigrateDataFromExcel.CustomEventArgs;

namespace TestProject
{
    [TestClass]
    public class MigrateDataTest
    {
        [TestMethod]
        public void TestMethod1()
        {
            {
                ISheet sheet = MigrateDataService.GetSheet("C:/Users/Wongsiripiphat/Desktop/Test.xls");

                MigrateDataService.OnSelfComposedProperty += OnSelfCompose;
                MigrateDataService.AfterComposedInfo += ValidateInfo;

                var properties = MigrateDataService.GetPropertyListFromInfo(typeof(ComplexInfo));

                var columnInfo = MigrateDataService.GetColumnCellInfoFromProperties(properties, sheet);

                Dictionary<string, ValidateRule> delegateDict = new Dictionary<string, ValidateRule>()
                {
                    { "IntProperty", delegate(object obj)
                                            {
                                                try
                                                {
                                                    var param = (int)obj;

                                                    if(param < 0)
                                                    {
                                                        return false;
                                                    }

                                                    return true;
                                                }
                                                catch(Exception ex)
                                                {
                                                    return false;
                                                }
                                            }}
                };

                var infoes = MigrateDataService.GetInfoes<ComplexInfo>(sheet, columnInfo, delegateDict, 1);
            }
        }

        public void OnSelfCompose(object sender, MigrateDataFromExcel.CustomEventArgs.SelfComposeValueEventArgs e)
        {

        }

        private bool ValidateInfo(object sender, AfterComposedInfoEventArgs args)
        {
            return false;
        }
    }
}

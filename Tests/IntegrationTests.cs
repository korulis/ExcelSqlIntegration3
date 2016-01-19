using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using IntegrationApp;
using Microsoft.Win32.SafeHandles;
using Xunit;

namespace Tests
{
    public static class AssertEx
    {
        public static void PropertyValuesAreEquals(object actual, object expected)
        {
            PropertyInfo[] properties = expected.GetType().GetProperties();
            foreach (PropertyInfo property in properties)
            {
                object expectedValue = property.GetValue(expected, null);
                object actualValue = property.GetValue(actual, null);

                if (actualValue is IList)
                    AssertListsAreEquals(property, (IList)actualValue, (IList)expectedValue);
                else if (!Equals(expectedValue, actualValue))
                    Assert.True(false,string.Format("Property {0}.{1} does not match. Expected: {2} but was: {3}", 
                        property.DeclaringType.Name, property.Name, expectedValue, actualValue));
            }
        }

        private static void AssertListsAreEquals(PropertyInfo property, IList actualList, IList expectedList)
        {
            if (actualList.Count != expectedList.Count)
                Assert.True(false,string.Format("Property {0}.{1} does not match. Expected IList containing {2} elements but was IList containing {3} elements",
                    property.PropertyType.Name, property.Name, expectedList.Count, actualList.Count));

            for (int i = 0; i < actualList.Count; i++)
                if (!Equals(actualList[i], expectedList[i]))
                    Assert.True(false,string.Format("Property {0}.{1} does not match. Expected IList with element {1} equals to {2} but was IList with element {1} equals to {3}",
                        property.PropertyType.Name, property.Name, expectedList[i], actualList[i]));
        }
    }

    public class IntegrationTests
    {

        //        [WpfFact]
        //        public void Tamtamtamtaaaam()
        //        {
        //            var app = new App();
        //            app.InitializeComponent();
        //            app.Run();
        //
        //
        //            Assert.Equal(1,1);
        //        }

        [Theory]
        [InlineData("")]
        [InlineData("a")]
        public void ExcelRepoCollectsCorrectDataFromExcel(string filePath)
        {


            //setup
            var vv = new FileStream("test.xlsx", FileMode.Open, FileAccess.Read);
            var excelRepo = new ExcelRepo(vv);
            //excerise
            var actual = excelRepo.GetDataFromExcel("SpecialSheetName1");
            //assert

            var expected = new List<List<string>>()
            {
               new List<string> { "42" }
            };

            //var expected = new List<object>
            //{
            //new {Id = "Id1", Comment = "FirstEntry", Amount = 100, Date = new DateTime(2016, 1, 18)},
            //new {Id = "Id2", Comment = "SecondEntry", Amount = 200, Date = new DateTime(2016, 1, 20)},
            //new {Id = "Id3", Comment = "ThirdEntry", Amount = 300.25 , Date = new DateTime(2007, 3, 12)}
            //};



            Assert.Equal(expected,actual);
        }
//        ([0-9]{0,3}(,[0-9]{3})*(\.[0-9]+)?)
        //        [Fact] 
        //        public void BrainTransfersCorrectDataFromExcelToSql() { }

        //        [Fact]
        //        public void BrainWritesCorrectDataToSql() { }

    }


}

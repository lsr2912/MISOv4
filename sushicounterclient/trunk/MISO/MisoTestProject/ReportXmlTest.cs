using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Xml;
#pragma warning disable CS0234 // The type or namespace name 'VisualStudio' does not exist in the namespace 'Microsoft' (are you missing an assembly reference?)
using Microsoft.VisualStudio.TestTools.UnitTesting;
#pragma warning restore CS0234 // The type or namespace name 'VisualStudio' does not exist in the namespace 'Microsoft' (are you missing an assembly reference?)
using SushiLibrary;

namespace MisoTestProject
{
#pragma warning disable CS0246 // The type or namespace name 'TestClassAttribute' could not be found (are you missing a using directive or an assembly reference?)
#pragma warning disable CS0246 // The type or namespace name 'TestClass' could not be found (are you missing a using directive or an assembly reference?)
    /// <summary>
    /// Summary description for ReportXmlTest
    /// </summary>
    [TestClass]
#pragma warning restore CS0246 // The type or namespace name 'TestClass' could not be found (are you missing a using directive or an assembly reference?)
#pragma warning restore CS0246 // The type or namespace name 'TestClassAttribute' could not be found (are you missing a using directive or an assembly reference?)
    public class ReportXmlTest
    {
        private XmlDocument JR1SampleReport;

        public ReportXmlTest()
        {
            FileStream JR1SampleFile = new FileStream("D:\\sushicounterclient\\MISO\\MISO\\JR1v3SampleData.xml", FileMode.Open, FileAccess.Read);
            JR1SampleReport = new XmlDocument();
            JR1SampleReport.Load(JR1SampleFile);

            JR1SampleFile.Close();
        }

#pragma warning disable CS0246 // The type or namespace name 'TestContext' could not be found (are you missing a using directive or an assembly reference?)
        private TestContext testContextInstance;
#pragma warning restore CS0246 // The type or namespace name 'TestContext' could not be found (are you missing a using directive or an assembly reference?)

#pragma warning disable CS0246 // The type or namespace name 'TestContext' could not be found (are you missing a using directive or an assembly reference?)
        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
#pragma warning restore CS0246 // The type or namespace name 'TestContext' could not be found (are you missing a using directive or an assembly reference?)
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

#pragma warning disable CS0246 // The type or namespace name 'TestMethod' could not be found (are you missing a using directive or an assembly reference?)
#pragma warning disable CS0246 // The type or namespace name 'TestMethodAttribute' could not be found (are you missing a using directive or an assembly reference?)
        [TestMethod]
#pragma warning restore CS0246 // The type or namespace name 'TestMethodAttribute' could not be found (are you missing a using directive or an assembly reference?)
#pragma warning restore CS0246 // The type or namespace name 'TestMethod' could not be found (are you missing a using directive or an assembly reference?)
        public void TestLoadingReport()
        {
            ReportLoader.LoadCounterReport(JR1SampleReport);
        }
    }
}

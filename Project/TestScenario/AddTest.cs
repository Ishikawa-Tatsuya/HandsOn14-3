using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace TestScenario
{
    [TestClass]
    public class AddTest : TestBase<AddTest>
    {
        [ClassInitialize]
        public static void ClassInitialize(TestContext c)
        {
            NotifyClassInitialize();
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            NotifyClassCleanup();
        }

        [TestInitialize]
        public void TestInitialize()
        {
            NotifyTestInitialize();
        }

        [TestCleanup]
        public void TestCleanup()
        {
            NotifyTestCleanup();
        }

        [TestMethod]
        public void TestAdd()
        {
            var addForm = App.MainForm.ButtonAdd_EmulateClick();
            addForm.TextBoxName.EmulateChangeText("ishikawa-tatsuya");
            addForm.TextBoxAddress.EmulateChangeText("Japan");
            addForm.RadioButtonMan.EmulateCheck();
            addForm.ButtonEntry_EmulateClickAndClose();
            Assert.AreEqual("ishikawa-tatsuya(男) Japan", App.MainForm.ListBoxEmployee_GetItemText(0));
        }

        class ErrorParam
        {
            public string Name { get; set; }
            public string Address { get; set; }
            public bool Checked { get; set; }
            public string Message { get; set; }
        }

        [TestMethod]
        [DataSource("System.Data.OleDB",
            @"Provider=Microsoft.ACE.OLEDB.12.0; Data Source=Params.xlsx; Extended Properties='Excel 12.0;HDR=yes';",
            "AddTest$",
            DataAccessMethod.Sequential
        )]     
        public void TestError()
        {
            var param = GetParam<ErrorParam>();
            var addForm = App.MainForm.ButtonAdd_EmulateClick();
            addForm.TextBoxName.EmulateChangeText(param.Name);
            addForm.TextBoxAddress.EmulateChangeText(param.Address);
            if (param.Checked)
            {
                addForm.RadioButtonMan.EmulateCheck();
            }
            Assert.AreEqual(param.Message, addForm.ButtonEntry_EmulateClickAndGetMessage());
            addForm.Close();
        }
    }
}

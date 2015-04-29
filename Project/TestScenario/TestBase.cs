using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using EmployeeManagementDriver;
using System.Collections.Generic;
using EmployeeManagement;
using System.Linq;

namespace TestScenario
{
    public class TestBase<T>
    {
        static protected AppDriver App { get; set; }
        static Dictionary<string, bool> _tests;
        public TestContext TestContext { get; set; }

        public static void NotifyClassInitialize()
        {
            App = new AppDriver();
            _tests = typeof(T).GetMethods().Where(e => 0 < e.GetCustomAttributes(typeof(TestMethodAttribute), true).Length).ToDictionary(e => e.Name, e => true);
        }

        public static void NotifyClassCleanup()
        {
            App.EndProcess();
        }

        public void NotifyTestInitialize()
        {
            App.Attach();
        }

        public void NotifyTestCleanup()
        {
            if (TestContext.DataRow == null ||
                ReferenceEquals(TestContext.DataRow, TestContext.DataRow.Table.Rows[TestContext.DataRow.Table.Rows.Count - 1]))
            {
                _tests.Remove(TestContext.TestName);
            }
            App.Release(TestContext.CurrentTestOutcome == UnitTestOutcome.Passed && 0 < _tests.Count);
        }

        public Data GetParam<Data>() where Data : new()
        {
            Data data = new Data();
            foreach (var e in typeof(Data).GetProperties())
            {
                e.GetSetMethod().Invoke(data, new object[] { Convert(e.PropertyType, TestContext.DataRow[e.Name]) });
            }
            return data;
        }

        static object Convert(Type type, object obj)
        {
            string value = obj == null ? string.Empty : obj.ToString();
            if (type == typeof(int))
            {
                return int.Parse(value);
            }
            else if (type == typeof(bool))
            {
                return string.Compare(value, true.ToString(), true) == 0;
            }
            else if (type == typeof(string))
            {
                return value;
            }
            throw new NotSupportedException();
        }
    }
}

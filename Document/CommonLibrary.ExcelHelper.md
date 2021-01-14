# CommonLibrary.ExcelHelper使用示例

## 控制台应用代码

```C#
using CommonLibrary.ExcelHelper.ExportStyle;
using CommonLibrary.ExcelHelper.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;

namespace CommonLibrary.ExcelHelper.Demo
{
    internal class Program
    {
        private static DataSet dataSet;
        private static List<Person> list;
        private static void Main()
        {
            CreateDemoData();//创建示例数据
            {//导出示例，数据源是List<T>，并且自定义导出文件的工作表名称，工作表设置样式
                var helper = ExcelHelperFactory.CreateExporter(list, Enum.ExcelVersion.XLSX, "列表导出测试");
                //var stream = helper.ExportToStream(new DefaultStyle());//导出到流
                helper.ExportToFile(@"..\test1.xlsx", new DefaultStyle());//导出到文件
            }
            Thread.Sleep(1000);
            {//导出示例，数据源是DataSet
                var helper = ExcelHelperFactory.CreateExporter(dataSet);
                //var stream = helper.ExportToStream();//导出到流
                helper.ExportToFile(@"..\test2.xlsx");//导出到文件
            }
            Thread.Sleep(1000);
            {//导出示例，数据源是DataTable
                var helper = ExcelHelperFactory.CreateExporter(dataSet.Tables[0]);
                //var stream = helper.ExportToStream();//导出到流
                helper.ExportToFile(@"..\test3.xlsx");//导出到文件
            }
            Thread.Sleep(1000);
            {//导入示例，导入数据生成List<T>
                var helper = ExcelHelperFactory.CreateImporter(@"..\test1.xlsx");
                var data = helper.Import<Person>();
                ShowInConsole_List(data);
            }
            Thread.Sleep(1000);
            {//导入示例，导入数据生成DataSet
                var helper = ExcelHelperFactory.CreateImporter(@"..\test2.xlsx");
                var data = helper.Import();
                ShowInConsole_DataSet(data);
            }
            Thread.Sleep(1000);
            {//导入示例，导入数据生成List<T>,且只导入多个工作表中的指定的一个
                List<ImportSheetSetting> ImportSheets = new List<ImportSheetSetting>() {
                    new ImportSheetSetting(1,0)
                };
                var helper = ExcelHelperFactory.CreateImporter(@"..\test2.xlsx", ImportSheets);
                var data = helper.Import();
                ShowInConsole_DataSet(data);
            }
            Console.ReadKey();
        }

        private static void ShowInConsole_DataSet(DataSet data)
        {
            foreach (DataTable Table in data.Tables)
            {
                foreach (DataColumn col in Table.Columns)
                {
                    Console.Write(col.ColumnName);
                    Console.Write("\t\t");
                }
                Console.WriteLine();
                foreach (DataRow row in Table.Rows)
                {
                    foreach (DataColumn col in Table.Columns)
                    {
                        Console.Write(row[col]);
                        Console.Write("\t\t");
                    }
                    Console.WriteLine();
                }
                Console.WriteLine();
                Console.WriteLine();
            }
        }

        private static void ShowInConsole_List(List<Person> data)
        {
            Console.WriteLine("ID\t\tName\t\tIDCard\t\tAge");
            foreach (var item in data)
            {
                Console.WriteLine($"{item.ID}\t\t{item.Name}\t\t{item.IDCard}\t\t{item.Age}");
            }
            Console.WriteLine();
            Console.WriteLine();
        }

        private static void CreateDemoData()
        {
            list = new List<Person>() {
                new Person(){  ID=1, Name="张三", IDCard="41050218604173000",Age=10 },
                new Person(){  ID=2, Name="张三", IDCard="41050218604173000" },
                new Person(){  ID=3, Name="张三", IDCard="41050218604173000" },
                new Person(){  ID=4, Name="张三", IDCard="41050218604173000" },
                new Person(){  ID=5, Name="张三", IDCard="41050218604173000" },
            };

            dataSet = new DataSet();
            CreatTableWithData(5);
            CreatTableWithData(10);
        }

        private static void CreatTableWithData(int count)
        {
            DataTable table = new DataTable();
            DataColumn column;
            DataRow row;

            column = new DataColumn
            {
                DataType = System.Type.GetType("System.Int32"),
                ColumnName = "id"
            };
            table.Columns.Add(column);

            column = new DataColumn
            {
                DataType = System.Type.GetType("System.String"),
                ColumnName = "Name"
            };
            table.Columns.Add(column);

            column = new DataColumn
            {
                DataType = System.Type.GetType("System.String"),
                ColumnName = "IDCard"
            };
            table.Columns.Add(column);

            dataSet.Tables.Add(table);

            for (int i = 0; i <= count; i++)
            {
                row = table.NewRow();
                row["id"] = i;
                row["Name"] = "ParentItem " + i;
                row["IDCard"] = "111111111111111111111111";
                table.Rows.Add(row);
            }
        }
    }

    internal class Person
    {
        public int ID { get; set; }

        public string Name { get; set; }

        public string IDCard { get; set; }

        public int? Age { get; set; }
    }
}
```

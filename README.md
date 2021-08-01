# MH.Excel.Export

`MH.Excel.Export` is a Excel (.xlsx) generator for ASP.NET Core applications.

## Install via NuGet

To install `MH.Excel.Export`, run the following command in the Package Manager Console:

```
PM> Install-Package MH.Excel.Export
``` 
You can also view the [package page](https://www.nuget.org/packages/MH.Excel.Export) on NuGet.

## Usage:

After installing the MH.Excel.Export package, you can send any classes with (or without) `[Display(Name ="")]` attribute and get required data for pass to `return File();`

![dntcaptcha](https://raw.githubusercontent.com/mohsen2hasani/MH.Excel/master/MH.Excel.Export.png)

```
    public class TestController : BaseController
    {
        public async Task<IActionResult> Index()
        {
            var list = new TestClass().GetList();

            var excel = await ExportManager.ExportToXlsxAsync<TestClass, SubClassTest>(list, "Test Excel Class");

            return File(excel.FileContents, excel.ContentType, excel.FileDownloadName);
        }

        public class TestClass
        {
            [Display(Name = "Id")]
            public int Id { get; set; }

            [Display(Name = "First Name")]
            public string Name { get; set; }

            [Display(Name = "Last Name")]
            public string Name2 { get; set; }


            [Display(Name = "Items")]
            public List<SubClassTest> List { get; set; } = new();


            public List<TestClass> GetList()
            {
                var tests = new List<TestClass>();

                for (int i = 1; i < 10; i++)
                {
                    var test = new TestClass
                    {
                        Id = i,
                        Name = $"Name - {i}",
                        Name2 = $"Name2 - {i}",
                    };

                    for (int j = 1; j < 10; j++)
                    {
                        test.List.Add(new SubClassTest
                        {
                            Id = 50 * j,
                            Type = $"Type - {j}",
                            Type2 = $"Type2 - {j}",
                            Type3 = $"Type3 - {j}"
                        });
                    }

                    tests.Add(test);
                }

                return tests;
            }
        }

        public class SubClassTest
        {
            [Display(Name = "Id")]
            public int Id { get; set; }

            [Display(Name = "attr1")]
            public string Type { get; set; }

            [Display(Name = "attr2")]
            public string Type2 { get; set; }

            [Display(Name = "attr3")]
            public string Type3 { get; set; }
        }
    }
```

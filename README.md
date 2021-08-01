Generate Excel from generic class with sub class:

![dntcaptcha](/MH.Excel.Export.png)

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

using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace MH.Excel.Test.Models
{
    public class TestClass
    {
        public List<TestClass> GetList()
        {
            var tests = new List<TestClass>();

            for (var i = 1; i < 10; i++)
            {
                var test = new TestClass
                {
                    Id = i,
                    Name = $"Name - {i}",
                    LastName = $"LastName - {i}",
                    Description = "Lorem ipsum is placeholder text"
                };

                for (var j = 1; j < 10; j++)
                {
                    test.List.Add(new SubClassTest
                    {
                        Id = 50 * j,
                        Type = $"Type - {j}",
                        Type2 = $"Type2 - {j}",
                        Type3 = $"Lorem ipsum is placeholder text - {j}"
                    });
                }

                for (var j = 1; j < 10; j++)
                {
                    test.List2.Add(new SubClassTest2
                    {
                        Id = 50 * j,
                        Type = $"Type - {j}",
                        Type2 = $"Type2 - {j}",
                        Type3 = $"Lorem ipsum is placeholder text - {j}"
                    });
                }

                tests.Add(test);
            }

            return tests;
        }

        [Display(Name = "Id")]
        public int Id { get; set; }

        [Display(Name = "First Name")]
        public string Name { get; set; }

        [Display(Name = "Last Name")]
        public string LastName { get; set; }

        [Display(Name = "About user")]
        public string Description { get; set; }

        [Display(Name = "Items")]
        public List<SubClassTest> List { get; set; } = new();

        [Display(Name = "Items 2")]
        public List<SubClassTest2> List2 { get; set; } = new();
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

    public class SubClassTest2
    {
        [Display(Name = "Id 2")]
        public int Id { get; set; }

        [Display(Name = "attr1 2")]
        public string Type { get; set; }

        [Display(Name = "attr2 2")]
        public string Type2 { get; set; }

        [Display(Name = "attr3 2")]
        public string Type3 { get; set; }
    }
}
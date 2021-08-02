using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using MH.Excel.Export.Helper;
using OfficeOpenXml;
using LicenseContext = OfficeOpenXml.LicenseContext;

namespace MH.Excel.Export
{
    /// <summary>
    /// Export manager
    /// </summary>
    public sealed class ExportManager
    {
        private static int _itemRow = 1;

        #region Utility

        private static string GetDisplayName<T>(MemberDescriptor prop)
        {
            return typeof(T).GetProperty(prop.Name)?
                .GetCustomAttribute(typeof(DisplayAttribute)) is DisplayAttribute data
                ? data.Name
                : prop.DisplayName ?? prop.Name;
        }

        private static async Task SubClassHandler<T, TSubClass>(IEnumerable properties, ExcelWorksheet baseWorksheet, Color captionBackgroundColor, T item)
        {
            foreach (var prop in properties.Cast<PropertyDescriptor>().Where(prop => prop.PropertyType.GetInterfaces().Any(x => x == typeof(ICollection<TSubClass>))))
            {
                var subClass1S = (prop.GetValue(item) as IEnumerable<TSubClass> ?? Array.Empty<TSubClass>()).ToList();

                if (!subClass1S.Any())
                    continue;

                _itemRow += 1;

                var subItemsManager = new PropertyManager<TSubClass>();

                var subClassProperties = TypeDescriptor.GetProperties(typeof(TSubClass));

                foreach (var subProp in subClassProperties.Cast<PropertyDescriptor>()
                    .Where(subProp => subProp.PropertyType.GetInterfaces().All(x => x != typeof(ICollection<>))))
                    subItemsManager.Add(new PropertyByName<TSubClass>(GetDisplayName<TSubClass>(subProp), a => subProp.GetValue(a)));

                subItemsManager.WriteCaption(baseWorksheet, captionBackgroundColor, _itemRow, 1);
                baseWorksheet.Row(_itemRow).OutlineLevel = 1;
                baseWorksheet.Row(_itemRow).Collapsed = true;

                foreach (var subClass1 in subClass1S)
                {
                    _itemRow++;
                    subItemsManager.CurrentObject = subClass1;
                    await subItemsManager.WriteToXlsxAsync(baseWorksheet, _itemRow, 1);
                    baseWorksheet.Row(_itemRow).OutlineLevel = 1;
                    baseWorksheet.Row(_itemRow).Collapsed = true;
                }
            }
        }

        #endregion

        /// <summary>
        /// Get excel data from class with sub class for sub table
        /// </summary>
        /// <param name="list">List of object</param>
        /// <param name="fileName">Download file name</param>
        /// <param name="rightToLeft">false</param>
        /// <param name="captionBackgroundColor">Caption background color</param>
        /// <typeparam name="T">Type of object</typeparam>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static async Task<ExcelData> ExportToXlsxAsync<T>(ICollection<T> list, string fileName, bool rightToLeft = false, Color? captionBackgroundColor = null)
        {
            return await ExportToXlsxAsync<T, T>(list, fileName, rightToLeft, captionBackgroundColor);
        }

        /// <summary>
        /// Get excel data from class with sub class for sub table
        /// </summary>
        /// <param name="list">List of object</param>
        /// <param name="fileName">Download file name</param>
        /// <param name="rightToLeft">false</param>
        /// <param name="captionBackgroundColor">Caption background color</param>
        /// <typeparam name="T">Type of object</typeparam>
        /// <typeparam name="TSubClass1">Type of sub class object</typeparam>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static async Task<ExcelData> ExportToXlsxAsync<T, TSubClass1>(ICollection<T> list, string fileName, bool rightToLeft = false, Color? captionBackgroundColor = null)
        {
            return await ExportToXlsxAsync<T, TSubClass1, T>(list, fileName, rightToLeft, captionBackgroundColor);
        }

        /// <summary>
        /// Get excel data from class with sub class for sub table
        /// </summary>
        /// <param name="list">List of object</param>
        /// <param name="fileName">Download file name</param>
        /// <param name="rightToLeft">false</param>
        /// <param name="captionBackgroundColor">Caption background color</param>
        /// <typeparam name="T">Type of object</typeparam>
        /// <typeparam name="TSubClass1">Type of sub class 1 object</typeparam>
        /// <typeparam name="TSubClass2">Type of sub class 2 object</typeparam>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static async Task<ExcelData> ExportToXlsxAsync<T, TSubClass1, TSubClass2>(ICollection<T> list, string fileName, bool rightToLeft = false,
            Color? captionBackgroundColor = null)
        {
            return await ExportToXlsxAsync<T, TSubClass1, TSubClass2, T>(list, fileName, rightToLeft, captionBackgroundColor);
        }

        /// <summary>
        /// Get excel data from class with sub class for sub table
        /// </summary>
        /// <param name="list">List of object</param>
        /// <param name="fileName">Download file name</param>
        /// <param name="rightToLeft">false</param>
        /// <param name="captionBackgroundColor">Caption background color</param>
        /// <typeparam name="T">Type of object</typeparam>
        /// <typeparam name="TSubClass1">Type of sub class 1 object</typeparam>
        /// <typeparam name="TSubClass2">Type of sub class 2 object</typeparam>
        /// <typeparam name="TSubClass3">Type of sub class 3 object</typeparam>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static async Task<ExcelData> ExportToXlsxAsync<T, TSubClass1, TSubClass2, TSubClass3>(ICollection<T> list, string fileName, bool rightToLeft = false,
            Color? captionBackgroundColor = null)
        {
            return await ExportToXlsxAsync<T, TSubClass1, TSubClass2, TSubClass3, T>(list, fileName, rightToLeft, captionBackgroundColor);
        }

        /// <summary>
        /// Get excel data from class with sub class for sub table
        /// </summary>
        /// <param name="list">List of object</param>
        /// <param name="fileName">Download file name</param>
        /// <param name="rightToLeft">false</param>
        /// <param name="captionBackgroundColor">Caption background color</param>
        /// <typeparam name="T">Type of object</typeparam>
        /// <typeparam name="TSubClass1">Type of sub class 1 object</typeparam>
        /// <typeparam name="TSubClass2">Type of sub class 2 object</typeparam>
        /// <typeparam name="TSubClass3">Type of sub class 3 object</typeparam>
        /// <typeparam name="TSubClass4">Type of sub class 4 object</typeparam>
        /// <returns></returns>
        /// <exception cref="ArgumentNullException"></exception>
        public static async Task<ExcelData> ExportToXlsxAsync<T, TSubClass1, TSubClass2, TSubClass3, TSubClass4>(ICollection<T> list, string fileName, bool rightToLeft = false,
            Color? captionBackgroundColor = null)
        {
            if (list == null)
                throw new ArgumentNullException(nameof(list));

            if (!list.Any())
                return default;

            captionBackgroundColor ??= Color.FromArgb(71, 195, 99);

            var propertyManager = new PropertyManager<T>();

            var properties = TypeDescriptor.GetProperties(typeof(T));

            foreach (var prop in properties.Cast<PropertyDescriptor>()
                .Where(prop => !prop.PropertyType.GetInterfaces().Any(x => x.IsGenericType && x.GetGenericTypeDefinition() == typeof(ICollection<>))))
                propertyManager.Add(new PropertyByName<T>(GetDisplayName<T>(prop), a => prop.GetValue(a)));

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            await using var stream = new MemoryStream();
            using (var xlPackage = new ExcelPackage(stream))
            {
                var baseWorksheet = xlPackage.Workbook.Worksheets.Add("Excel");
                baseWorksheet.View.RightToLeft = rightToLeft;

                propertyManager.WriteCaption(baseWorksheet, captionBackgroundColor.Value);

                _itemRow = 1;

                foreach (var item in list)
                {
                    _itemRow += 1;
                    propertyManager.CurrentObject = item;
                    await propertyManager.WriteToXlsxAsync(baseWorksheet, _itemRow);

                    if (typeof(TSubClass1) != typeof(T))
                        await SubClassHandler<T, TSubClass1>(properties, baseWorksheet, captionBackgroundColor.Value, item);

                    if (typeof(TSubClass2) != typeof(T))
                        await SubClassHandler<T, TSubClass2>(properties, baseWorksheet, captionBackgroundColor.Value, item);

                    if (typeof(TSubClass3) != typeof(T))
                        await SubClassHandler<T, TSubClass3>(properties, baseWorksheet, captionBackgroundColor.Value, item);

                    if (typeof(TSubClass4) != typeof(T))
                        await SubClassHandler<T, TSubClass4>(properties, baseWorksheet, captionBackgroundColor.Value, item);
                }

                baseWorksheet.Cells[baseWorksheet.Dimension.Address].AutoFitColumns();

                await xlPackage.SaveAsync();
            }

            return new ExcelData
            {
                FileContents = stream.ToArray(),
                ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                FileDownloadName = $"{fileName}.xlsx"
            };
        }
    }
}
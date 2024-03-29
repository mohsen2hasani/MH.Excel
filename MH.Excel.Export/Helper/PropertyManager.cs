﻿using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace MH.Excel.Export.Helper;

/// <summary>
/// Class for working with PropertyByName object list
/// </summary>
/// <typeparam name="T">Object type</typeparam>
public sealed class PropertyManager<T>
{
    /// <summary>
    /// All properties
    /// </summary>
    private readonly Dictionary<string, PropertyByName<T>> _properties = new();

    /// <summary>
    /// Current object to access
    /// </summary>
    public T CurrentObject { get; set; }

    private int _poz = 1;

    public void Add(PropertyByName<T> propertyByName)
    {
        if (!propertyByName.Ignore)
        {
            propertyByName.PropertyOrderPosition = _poz;
            _poz++;
            _properties.Add(propertyByName.PropertyName, propertyByName);
        }
    }

    /// <summary>
    /// Write object data to XLSX worksheet
    /// </summary>
    /// <param name="worksheet">Data worksheet</param>
    /// <param name="row">Row index</param>
    /// <param name="cellOffset">Cell offset</param>
    /// <returns>A task that represents the asynchronous operation</returns>
    public async Task WriteToXlsxAsync(ExcelWorksheet worksheet, int row, int cellOffset = 0)
    {
        if (CurrentObject == null)
            return;

        foreach (var prop in _properties.Values)
        {
            var cell = worksheet.Cells[row, prop.PropertyOrderPosition + cellOffset];

            cell.Value = await prop.GetProperty(CurrentObject);
        }
    }

    /// <summary>
    /// Write caption (first row) to XLSX worksheet
    /// </summary>
    /// <param name="worksheet">worksheet</param>
    /// <param name="backgroundColor">Caption background color</param>
    /// <param name="row">Row number</param>
    /// <param name="cellOffset">Cell offset</param>
    public void WriteCaption(ExcelWorksheet worksheet, Color backgroundColor, int row = 1, int cellOffset = 0)
    {
        foreach (var caption in _properties.Values)
        {
            var cell = worksheet.Cells[row, caption.PropertyOrderPosition + cellOffset];
            cell.Value = caption;

            SetCaptionStyle(cell, backgroundColor);
            cell.Style.Hidden = false;
        }
    }

    /// <summary>
    /// Set caption style to excel cell
    /// </summary>
    /// <param name="cell">Excel cell</param>
    /// <param name="backgroundColor">Caption background color</param>
    public void SetCaptionStyle(ExcelRange cell, Color backgroundColor)
    {
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(backgroundColor);
        cell.Style.Font.Bold = true;
    }
}
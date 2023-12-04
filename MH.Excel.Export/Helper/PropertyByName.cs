using System;
using System.Threading.Tasks;

namespace MH.Excel.Export.Helper;

/// <summary>
/// A helper class to access the property by name
/// </summary>
/// <typeparam name="T">Object type</typeparam>
public class PropertyByName<T>
{
    /// <summary>
    /// Ctor
    /// </summary>
    /// <param name="propertyName">Property name</param>
    /// <param name="func">Feature property access</param>
    /// <param name="ignore">Specifies whether the property should be exported</param>
    public PropertyByName(string propertyName, Func<T, object> func = null, bool ignore = false)
    {
        PropertyName = propertyName;
            
        if(func != null)
            GetProperty = obj => Task.FromResult(func(obj));

        PropertyOrderPosition = 1;
        Ignore = ignore;
    }

    /// <summary>
    /// Property order position
    /// </summary>
    public int PropertyOrderPosition { get; set; }

    /// <summary>
    /// Feature property access
    /// </summary>
    /// <returns>A task that represents the asynchronous operation</returns>
    public Func<T, Task<object>> GetProperty { get; }

    /// <summary>
    /// Property name
    /// </summary>
    public string PropertyName { get; }

    /// <summary>
    /// To string
    /// </summary>
    /// <returns>String</returns>
    public override string ToString()
    {
        return PropertyName;
    }

    /// <summary>
    /// Specifies whether the property should be exported
    /// </summary>
    public bool Ignore { get; set; }
}
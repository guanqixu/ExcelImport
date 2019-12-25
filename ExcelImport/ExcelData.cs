using System;
using System.Collections.Generic;

namespace ExcelImport
{
    /// <summary>
    /// Excel 的数据
    /// </summary>
    public abstract class ExcelData
    {

        /// <summary>
        /// 标题属性
        /// </summary>
        private Dictionary<string, string> _header_property;

        /// <summary>
        /// 获取标题属性
        /// </summary>
        /// <returns></returns>
        public virtual Dictionary<string, string> GetHeaderProperty()
        {
            if (_header_property == null)
            {
                _header_property = new Dictionary<string, string>();
                var properties = GetType().GetProperties();
                foreach (var prop in properties)
                {
                    var attribute = Attribute.GetCustomAttribute(prop, typeof(ExcelHeaderAttribute));
                    string name = (attribute as ExcelHeaderAttribute).Header;
                    _header_property.Add(name, prop.Name);
                }
            }
            return _header_property;
        }
    }

    /// <summary>
    /// Excel 标题
    /// </summary>
    public class ExcelHeaderAttribute : Attribute
    {
        /// <summary>
        /// 头
        /// </summary>
        public string Header { get; set; }

        /// <summary>
        /// 注释
        /// </summary>
        public string Comment { get; set; }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="header"></param>
        public ExcelHeaderAttribute(string header)
        {
            Header = header;
        }

    }
}

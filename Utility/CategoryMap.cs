// ***********************************************************************
// Assembly         : OppmUtility
// Author           : ggreiff
// Created          : 10-09-2014
//
// Last Modified By : ggreiff
// Last Modified On : 10-09-2014
// ***********************************************************************
// <copyright file="CategoryMap.cs" company="">
//     Copyright (c) . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OppmUtility.Utility
{
    /// <summary>
    /// Class CategoryMap.
    /// </summary>
    public class CategoryMap
    {
        /// <summary>
        /// Gets or sets the name of the column.
        /// </summary>
        /// <value>The name of the column.</value>
        public String ColumnName { get; set; }

        /// <summary>
        /// Gets or sets the column letter.
        /// </summary>
        /// <value>The column letter.</value>
        public String ColumnLetter { get; set; }

        /// <summary>
        /// Gets the column number.
        /// </summary>
        /// <value>The column number.</value>
        public Int32 ColumnNumber
        {
            get { return ExcelColumnNameToNumber(ColumnLetter); }
        }

        /// <summary>
        /// Gets or sets the name of the category.
        /// </summary>
        /// <value>The name of the category.</value>
        public String CategoryName { get; set; }

        /// <summary>
        /// Gets or sets the name of the value list.
        /// </summary>
        /// <value>The name of the value list.</value>
        public String ValueListName { get; set; }

        /// <summary>
        /// Gets or sets the item name flag.
        /// </summary>
        /// <value>The item name flag.</value>
        public String ColumnFlagType { get; set; }


        /// <summary>
        /// Gets the data column number. Excel start with 1 DataTables with 0
        /// </summary>
        /// <value>The data column number.</value>
        public Int32 DataColumnNumber => ColumnNumber - 1;


        /// <summary>
        /// Determines whether [is item name map].
        /// </summary>
        public Boolean IsItemNameMap
        {
            get
            {
                if (ColumnFlagType.IsTrimEqualTo("UseAsSubItemID", true)) return false;
                return ColumnFlagType.IsTrimEqualTo("Yes", true) || ColumnFlagType.IsTrimEqualTo("UseAsName", true) || CategoryName.IsNullOrEmpty();
            }
        }

        /// <summary>
        /// Gets the use uci map.
        /// </summary>
        /// <value>The use uci map.</value>
        public Boolean UseUciMap => ColumnFlagType.IsTrimEqualTo("UseAsUCI", true);

        public Boolean IsSubItemKey => ColumnFlagType.IsTrimEqualTo("SubItemKey", true) || ColumnFlagType.IsTrimEqualTo("UseAsSubItemID", true);


        public CategoryMap()
        {
            ColumnFlagType = String.Empty;
            CategoryName = String.Empty;
        }

        /// <summary>
        /// Excels the column name to number.
        /// </summary>
        /// <param name="columnName">Name of the column.</param>
        /// <returns>System.Int32.</returns>
        /// <exception cref="System.ArgumentNullException">columnName</exception>
        private static Int32 ExcelColumnNameToNumber(string columnName)
        {
            if (String.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");
            
            // Already a numbere
            var columnNumber = columnName.ToInt();
            if (columnNumber.HasValue) return columnNumber.Value;
            
            // Try the letter version
            columnName = columnName.ToUpperInvariant();
            var sum = 0;
            foreach (var t in columnName)
            {
                sum *= 26;
                sum += (t - 'A' + 1);
            }
            return sum;
        }

        public override String ToString()
        {
            var sb = new StringBuilder();
            sb.AppendFormat("IsItemNameMap {0}\r\n", IsItemNameMap);
            sb.AppendFormat("IsSubItemKey {0}\r\n", IsSubItemKey);
            sb.AppendFormat("UseUciMap {0}\r\n", UseUciMap);
            sb.AppendFormat("ColumnNumber {0}\r\n", ColumnNumber);
            sb.AppendFormat("DataColumnNumber {0}\r\n", DataColumnNumber);
            sb.AppendFormat("CategoryName {0}\r\n", CategoryName);
            sb.AppendFormat("ColumnFlagType {0}\r\n", ColumnFlagType);
            sb.AppendFormat("ColumnLetter {0}\r\n", ColumnLetter);
            sb.AppendFormat("ValueListName {0}\r\n", ValueListName);
            return sb.ToString();
        }
    }

    /// <summary>
    /// Class DataColumns.
    /// </summary>
    public class DataColumns
    {
        /// <summary>
        /// Gets or sets the column maps.
        /// </summary>
        /// <value>The column maps.</value>
        public List<CategoryMap> CategoryMapList { get; set; }

        /// <summary>
        /// Gets the item name map exits.
        /// </summary>
        /// <value>The item name map exits.</value>
        public Boolean ItemNameMapExits
        {
            get
            {
                return CategoryMapList.Find(x => x.IsItemNameMap) != null;
            }
        }

        /// <summary>
        /// Gets the uci map exits.
        /// </summary>
        /// <value>The uci map exits.</value>
        public Boolean UciMapExits
        {
            get
            {
                return CategoryMapList.Find(x => x.UseUciMap) != null;
            }
        }

        /// <summary>
        /// Gets the item name map.
        /// </summary>
        /// <returns>CategoryMap.</returns>
        public CategoryMap GetItemNameUciMap
        {
            get { return CategoryMapList.Find(x => x.UseUciMap) ?? CategoryMapList.Find(x => x.IsItemNameMap); }
        }

        public List<String> GetSubItemKeys
        {
            get
            {
                var subItemKeys = CategoryMapList.FindAll(x => x.IsSubItemKey);
                if (subItemKeys.Count == 1 && subItemKeys[0].ColumnFlagType.IsTrimEqualTo("UseAsSubItemID", true))
                {
                    return new List<String> { "UseAsSubItemID" };
                }
                return CategoryMapList.FindAll(x => x.IsSubItemKey).Select(x => x.CategoryName).ToList();
             }
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="DataColumns"/> class.
        /// </summary>
        public DataColumns()
        {
            CategoryMapList = new List<CategoryMap>();
        }

        /// <summary>
        /// Adds the specified column map.
        /// </summary>
        /// <param name="categoryMap">The column map.</param>
        /// <returns>Boolean.</returns>
        public Boolean Add(CategoryMap categoryMap)
        {
            CategoryMapList.Add(categoryMap);
            return true;
        }

        /// <summary>
        /// Finds the name of the by column.
        /// </summary>
        /// <param name="columnName">Name of the column.</param>
        /// <returns>CategoryMap.</returns>
        public CategoryMap FindByColumnName(String columnName)
        {
            return CategoryMapList.Find(x => x.ColumnName.IsTrimEqualTo(columnName, true));
        }

        /// <summary>
        /// Finds the by column letter.
        /// </summary>
        /// <param name="columnLetter">The column letter.</param>
        /// <returns>CategoryMap.</returns>
        public CategoryMap FindByColumnLetter(String columnLetter)
        {
            return CategoryMapList.Find(x => x.ColumnLetter.IsTrimEqualTo(columnLetter, true));
        }

        /// <summary>
        /// Finds the by column number.
        /// </summary>
        /// <param name="columnNumber">The column number.</param>
        /// <returns>CategoryMap.</returns>
        public CategoryMap FindByColumnNumber(Int32 columnNumber)
        {
            return CategoryMapList.Find(x => x.ColumnNumber.Equals(columnNumber));
        }

        /// <summary>
        /// Finds the by column number.
        /// </summary>
        /// <param name="columnLetter">The column letter.</param>
        /// <returns>CategoryMap.</returns>
        public CategoryMap FindByColumnNumber(String columnLetter)
        {
            return CategoryMapList.Find(x => x.ColumnLetter.IsTrimEqualTo(columnLetter, true));
        }
        
        /// <summary>
        /// Finds the name of the by category.
        /// </summary>
        /// <param name="categoryName">Name of the category.</param>
        /// <returns>CategoryMap.</returns>
        public CategoryMap FindByCategoryName(String categoryName)
        {
            return CategoryMapList.Find(x => x.CategoryName.IsTrimEqualTo(categoryName, true));
        }


    }
}

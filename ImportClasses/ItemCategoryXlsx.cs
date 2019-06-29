using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using ClosedXML.Excel;
using NLog;
using OppmUtility.Utility;
using wsPortfoliosCategory;
using Cell = wsPortfoliosCell;
using SubItem = wsPortfoliosSubItem;
using wsPortfoliosPortfolio;


namespace OppmUtility.ImportClasses
{
    public class ItemCategoryXlsx
    {

        public static Logger NLogger = LogManager.GetCurrentClassLogger();

        public X509Certificate Certificate { get; set; }

        public Boolean RunImport(Options options)
        {
            var importPortfolio = options.ImportPortfolio;
            var retVal = false;
            var checkStatus = false;
            var badSubItemType = false;
            var subItemKeys = new List<String>();

            try
            {
                if (Properties.Settings.Default.Properties["ImportPortfolio"] != null)
                    importPortfolio = Properties.Settings.Default.ImportPortfolio;

                var goodToGo = true;
                var xlsxDataTable = new DataTable();
                using (var wb = new XLWorkbook(options.XlsxDataFileName))
                {
                NLogger.Info("Processing workbook {0}", options.XlsxDataFileName);

                var wsData = wb.Worksheets.ToList().Find(w => w.Name.IsEqualTo(options.XlsxDataSheetName, true));
                if (wsData == null)
                {
                    NLogger.Error("Unable to find a Data worksheet named {0}", options.XlsxDataSheetName);
                        goodToGo = false;
                }
                if (goodToGo) xlsxDataTable = wsData.ToDataTable();                  
                }
                if (!goodToGo) return false;

                var xlsxMapTable = new DataTable();
                using (var wb = new XLWorkbook(options.XlsxMapFileName))
                {
                    var wsMap = wb.Worksheets.ToList().Find(w => w.Name.IsEqualTo(options.XlsxMapSheetName, true));
                    if (wsMap == null)
                    {
                        NLogger.Error("Unable to find a Map worksheet named {0}", options.XlsxMapSheetName);
                        goodToGo = false;
                    }

                    if (goodToGo) xlsxMapTable = wsMap.ToDataTable();

                }
                if (!goodToGo) return false;

                //
                // Datatables are easier to use than worksheets so just convert it to a datatable
                //
                if (!xlsxDataTable.Rows.HasItems())
                {
                    NLogger.Error("Either there was no rows on {0} or we had an error converting the Data worksheet to a datatable.", options.XlsxDataSheetName);
                    return false;
                }

                
                if (!xlsxMapTable.Rows.HasItems())
                {
                    NLogger.Error("Either there was no rows on {0} or we had an error converting the Map worksheet to a datatable.", options.XlsxMapSheetName);
                    return false;
                }

                //
                // Build our map list
                //
                var dataColumns = new DataColumns();
                foreach (DataRow row in xlsxMapTable.Rows)
                {
                    var columnMap = new CategoryMap
                    {
                        ColumnName = row.Field<String>(0),
                        ColumnLetter = row.Field<String>(1),
                        CategoryName = row.Field<String>(2),
                        ColumnFlagType = row.Field<String>(3)
                    };
                    if (columnMap.ColumnName.IsNullOrEmpty()) continue;
                    if (columnMap.CategoryName.IsTrimEqualTo("status", true)) checkStatus = true;
                    dataColumns.Add(columnMap);
                }

                //
                // OK lets log into OPPM
                //
                var oppm = new Oppm(options.OppmUser, options.OppmPassword, options.OppmHost, options.UseSsl);
                if (options.UseCert)
                {
                    oppm = new Oppm(options.OppmUser, options.OppmPassword, options.OppmHost, options.UseCert, Certificate);
                }
                var loggedin = oppm.Login();

                if (!loggedin)
                {
                    NLogger.Error("Unable to log into {0} with {1}", options.OppmHost, options.OppmUser);
                    return false;
                }
                NLogger.Info("Successful login with {0}", options.OppmUser);

                //
                // Check to make sure all the categories exists and we have an item name column
                //
                if (!(dataColumns.ItemNameMapExits || dataColumns.UciMapExits))
                {
                    NLogger.Error("Unable to find a ItemName either by marking the row with a Yes or UCI (blank mapping) in the mapping worksheet: {0}", options.XlsxMapSheetName);
                    return false;
                }

                //
                // Now check to make sure all the categories exists
                //
                var categoryInfoList = new List<psPortfoliosCategoryInfo>();
                var valueListDictionary = new Dictionary<String, List<String>>();
                var badCategoryList = new List<String>();
                foreach (var categoryMap in dataColumns.CategoryMapList)
                {
                    if (categoryMap.IsItemNameMap || categoryMap.CategoryName.IsNullOrEmpty()) continue; // A blank mapping uses this column as the name
                    var portfolioCategoryInfo = oppm.SeCategory.GetCategoryInfo(categoryMap.CategoryName);
                    if (portfolioCategoryInfo != null)
                    {
                        categoryInfoList.Add(portfolioCategoryInfo);
                        if (portfolioCategoryInfo.ValueListName.IsNotNullOrEmpty())
                        {
                            categoryMap.ValueListName = portfolioCategoryInfo.ValueListName;
                            var values = oppm.SeValueList.GetValueListText(categoryMap.ValueListName);
                            if (valueListDictionary.ContainsKey(categoryMap.ValueListName)) continue;
                            valueListDictionary.Add(categoryMap.ValueListName, values);
                        }
                        continue;
                    }
                    badCategoryList.Add(categoryMap.CategoryName);
                }

                //
                // Now check the values list values
                //
                var valueListCategories = dataColumns.CategoryMapList.FindAll(x => x.ValueListName.IsNotNullOrEmpty());
                var badValuesList = new List<String>();
                if (valueListCategories.HasItems())
                {
                    foreach (DataRow row in xlsxDataTable.Rows)
                    {
                        foreach (var valueListCategory in valueListCategories)
                        {
                            // Get our category value for a value list category
                            var value = row[valueListCategory.DataColumnNumber].ToString();
                            if (value.IsNullOrEmpty()) continue;
                            if (!valueListDictionary.ContainsKey(valueListCategory.ValueListName)) continue; // should never happen

                            // Get our value list values
                            var validValueList = valueListDictionary[valueListCategory.ValueListName];
                            if (validValueList.Contains(value)) continue;
                            var badValueString = String.Format("ValueList {0} on Category {1} does not contain value: {2}", valueListCategory.ValueListName, valueListCategory.CategoryName, value);
                            if (badValuesList.Contains(badValueString)) continue;
                            badValuesList.Add(badValueString);
                        }
                    }
                }

                //
                // Now check to see if 
                //
                if (options.SubItemType.IsNotNullOrEmpty())
                {
                    badSubItemType = !oppm.SeValueList.HasValue("Dynamic List Types", options.SubItemType);
                    subItemKeys = dataColumns.GetSubItemKeys;
                }

                //
                // See if we have bad values for the value list categories
                //
                var foundError = false;

                //
                // Did we get any bad categories
                //
                if (badCategoryList.HasItems())
                {
                    foreach (var badCategory in badCategoryList)
                    {
                        NLogger.Warn("Category {0} doesn't exists.", badCategory);
                    }
                    NLogger.Warn("Correct category mapping errors in the Map worksheet: {0}", options.XlsxMapSheetName);
                    foundError = true;
                }

                //
                // Did we get any bad values
                //
                if (badValuesList.HasItems())
                {
                    foreach (var value in badValuesList)
                    {
                        NLogger.Warn(value);
                    }
                    NLogger.Warn("Correct value list values in the Data worksheet: {0}", options.XlsxDataSheetName);
                    foundError = true;
                }

                //
                // Bad subitem type list value -- didn't find it in Dynamic List Types
                //
                if (badSubItemType)
                {
                    NLogger.Warn("Unable to find {0} subitem type in Dynamic List Types valuelist", options.SubItemType);
                    foundError = true;
                }

                if (options.SubItemType.IsNotNullOrEmpty() && !subItemKeys.HasItems())
                {
                    NLogger.Warn("SubItem loading required at least one SubItemKey column to be defined on the map sheet for {0}", options.SubItemType);
                    foundError = true;
                }

                //
                // The user needs to fix the errors we found.
                //
                if (foundError) return false;


                //
                // Look for our import portfolio.  If we don't find it create it.
                //
                var msg = "Found";
                var importPortfolioId = oppm.SeItem.GetItemIdByName(importPortfolio);
                if (importPortfolioId.IsNullOrEmpty())
                {
                    var importPortfolioInfo = new psPortfoliosItemInfo
                    {
                        Name = importPortfolio,
                        Status = psOPEN_CLOSED_STATUS.OCSTS_OPEN,
                        PortfolioType = psPORTFOLIO_TYPE.PTYP_PROJECTS,
                        Description = "This is the import portfolio that was created by OppmUtility",
                        IsContainerRoot = true,
                        CalculationLevel = psCALCULATION_LEVEL.CL_NOT_CALCULATED
                    };
                    var psReturnValues = oppm.SePorfolio.AddPortfolio(importPortfolioInfo);

                    if (psReturnValues == wsPortfoliosPortfolio.psRETURN_VALUES.ERR_OK) importPortfolioId = oppm.SeItem.GetItemIdByName(importPortfolio);
                    msg = "Created ";
                }
                if (importPortfolioId.IsNullOrEmpty())
                {
                    NLogger.Fatal("Unable to create import portfolio {0}", importPortfolio);
                    return false;
                }
                NLogger.Info("{0} import porfolio {1}", msg, importPortfolio);

                //
                // OK lets do some work
                //
                var rowCnt = 1;
                var itemNameUciMap = dataColumns.GetItemNameUciMap;
                var categories = dataColumns.CategoryMapList.Where(x => x.IsItemNameMap != true && x.CategoryName.IsNotNullOrEmpty()).Select(x => x.CategoryName).ToList();
                NLogger.Trace("itemNameUciMap = {0}", itemNameUciMap);
                foreach (DataRow row in xlsxDataTable.Rows)
                {
                    var itemNameUciValue = row.Field<String>(itemNameUciMap.DataColumnNumber);
                    NLogger.Trace("itemNameUciValue = {0}", itemNameUciValue);
                    var status = wsPortfoliosItem.psOPEN_CLOSED_STATUS.OCSTS_OPEN;

                    var importItemCellInfoList = new List<Cell.psPortfoliosCellInfo>();
                    var importSubItemCellInfoList = new List<SubItem.psPortfoliosCellInfo>();
                    var subItemID = String.Empty;

                    foreach (var categoryMap in dataColumns.CategoryMapList)
                    {
                        //var columnValue = row.Field<String>(categoryMap.DataColumnNumber);
                        //NLogger.Trace("categoryMap = {0}", categoryMap.ToString());
                        if (categoryMap.IsItemNameMap || categoryMap.UseUciMap) continue; //we already got the name

                        if (categoryMap.ColumnFlagType.IsEqualTo("UseAsSubItemId", true))
                        {
                            subItemID = row.Field<String>(categoryMap.DataColumnNumber);
                            continue;
                        }

                        //
                        // check to see if we need to set the item status.
                        //
                        NLogger.Trace("categoryMap.CategoryName = {0}", categoryMap.CategoryName);
                        if (checkStatus && categoryMap.CategoryName.IsTrimEqualTo("status", true))
                        {
                            var columnStatus = row.Field<String>(categoryMap.DataColumnNumber);
                            if (columnStatus.IsTrimEqualTo("CLOSED", true)) status = wsPortfoliosItem.psOPEN_CLOSED_STATUS.OCSTS_CLOSED;
                            if (columnStatus.IsTrimEqualTo("CANDIDATE", true)) status = wsPortfoliosItem.psOPEN_CLOSED_STATUS.OCSTS_CANDIDATE;
                            continue;
                        }

                        //
                        // Find our category create our cellinfo and add it to the cellinfo list depending if item or subitem.
                        //
                        var portfolioCategoryInfo = categoryInfoList.Find(x => x.Name.IsTrimEqualTo(categoryMap.CategoryName, true));
                        if (portfolioCategoryInfo == null)
                        {
                            NLogger.Trace("portfolioCategoryInfo is null");
                            continue;
                        }


                        NLogger.Trace("portfolioCategoryInfo.Name = {0}", portfolioCategoryInfo.Name);
                        //if (portfolioCategoryInfo.ValueListName.IsNotNullOrEmpty() && columnValue.IsNullOrEmpty()) continue;
                        if (options.SubItemType.IsNotNullOrEmpty())
                        {
                            importSubItemCellInfoList.Add(HelperFunctions.BuildSubItemCellInfo(portfolioCategoryInfo, row.Field<String>(categoryMap.DataColumnNumber)));
                            NLogger.Trace("importSubItemCellInfoList = {0}", "Added");
                            continue;
                        }
                        importItemCellInfoList.Add(HelperFunctions.BuildItemCellInfo(portfolioCategoryInfo, row.Field<String>(categoryMap.DataColumnNumber)));
                        NLogger.Trace("importItemCellInfoList = {0}", "Added");
                    }

                    //
                    // Make sure we have a set of subitem keys
                    //
                    /*
                    if (importItemCellInfoList.HasItems())
                    {
                        var validKeyCount = 0;
                        foreach (var subItemKey in subItemKeys)
                        {
                            var cell = importItemCellInfoList.Find(x => x.CategoryName.IsTrimEqualTo(subItemKey, true));
                            if (cell == null || cell.CellDisplayValue.IsNullOrEmpty()) continue;
                            validKeyCount++;
                        }
                        if (!validKeyCount.Equals(subItemKeys.Count))
                        {
                            NLogger.Warn("Invalid subItem key on line {0} for item {1}", rowCnt++, itemNameUciValue);
                        }
                    }
                     */


                    //
                    // Check to see if we need to add the item.
                    //
                    var msgOne = String.Format("Process row {0} out of {1} ", rowCnt++, xlsxDataTable.Rows.Count);
                    msgOne = String.Format(itemNameUciMap.IsItemNameMap ? "{0} with ItemName {1}" : "{0} with {2} {1}", msgOne, itemNameUciValue, itemNameUciMap.CategoryName);
                    NLogger.Info(msgOne);

                    if (!options.Commit) continue;

                    //
                    // OK lets get our item that we want to write this data to.
                    //
                    NLogger.Trace("Check Existance for {0}", itemNameUciValue);
                    wsPortfoliosItem.psPortfoliosItemInfo itemInfo = null;

                    // if we are using item name
                    if (itemNameUciMap.IsItemNameMap)
                    {
                        itemInfo = oppm.SeItem.GetItemInfoByName(itemNameUciValue);
                        if (itemInfo == null)
                        {
                            if (!options.NewItem)
                            {
                                NLogger.Info("Item {0} doesn't exists", itemNameUciValue);
                                continue;
                            }

                            // Create a new item.
                            var itemId = oppm.SeItem.AddNewEx(itemNameUciValue, importPortfolio);
                            itemInfo = oppm.SeItem.GetItemInfo(itemId);
                            NLogger.Info("Added {0} to portfolio {1} with id {2}", itemInfo.Name, importPortfolio, itemInfo.ProSightID);
                        }
                    }

                    // if we are using uci
                    if (itemNameUciMap.UseUciMap)
                    {
                        itemInfo = oppm.SeItem.GetItemInfo(itemNameUciMap.CategoryName, itemNameUciValue);
                        if (itemInfo == null)
                        {
                            NLogger.Info("UCI {0} doesn't exists", itemNameUciValue);
                            continue;
                        }
                    }

                    // We didn't find the item -- bummer!
                    if (itemInfo == null)
                    {
                        NLogger.Warn("Unable to find an item base on either the ItemName or UCI for row {0}", rowCnt - 1);
                        continue;
                    }

                    //
                    // See if the status has changed
                    //
                    if (checkStatus && itemInfo.Status != status)
                    {
                        NLogger.Trace("Checking status of {0}", itemInfo.Name);
                        itemInfo.Status = status;
                        oppm.SeItem.UpdateEx(itemInfo, 32);
                    }

                    //
                    // If subitem type is null then update our import item cells.
                    //
                    if (options.SubItemType.IsNullOrEmpty())
                    {
                        // Process it as an Item
                        NLogger.Trace("Doing Item UpdateMultipleCellsEx");
                        var updateStatusList = oppm.SeCell.UpdateMultipleCellsEx(itemInfo.ProSightID, importItemCellInfoList);

                        NLogger.Info("Update item cells on {0}", itemNameUciValue);
                        if (!updateStatusList.HasItems()) continue;

                        //
                        // Let see if the update when well
                        //
                        NLogger.Trace("Logging update status");
                        foreach (var portfolioCellUpdateStatus in updateStatusList)
                        {
                            NLogger.Warn("On item {0}, the following cell didn't update {1}.  Error {2}", itemNameUciValue, portfolioCellUpdateStatus.CategoryName, portfolioCellUpdateStatus.ErrorText);
                        }

                        //
                        // Do the next item cell import
                        //
                        continue;
                    }

                    //
                    // Process subitem cells based on the subItemKeys
                    //
                    NLogger.Trace("Doing subitem SyncSubItemsAsOfToday");

                    //
                    // Get our subitems base on the type from our current item.
                    //
                    var asOf = DateTime.Now;
                    var subItemInfos = oppm.SeSubItem.GetSubItemListAsOf(String.Empty, itemInfo.ProSightID.ToString(CultureInfo.InvariantCulture), options.SubItemType, 0, categories, false, DateTime.Now);
                    NLogger.Info("Updating {0} subItemInfos on {1}", subItemInfos.Count, itemInfo.Name);


                    //
                    // Check for an existing subitme based on the subitems keys
                    //
                    var doSubItemSync = true;
                    SubItem.psPortfoliosSubItemInfo foundKeyedSubItem = null;
                    var subItemUciList = new List<String>(subItemInfos.Select(x => x.SubItemUCI));

                    if (subItemID.IsNotNullOrEmpty())
                    {
                        NLogger.Trace("Doing UseAsSubItemID Sync");
                        
                        doSubItemSync = false;
                        var subItemInfo = subItemInfos.Find(x => x.SubItemUCI.IsTrimEqualTo(subItemID, true));
                        if (subItemInfo != null)
                        {
                            NLogger.Trace("Found subItemInfo {0}", subItemInfo.SubItemUCI);
                            foundKeyedSubItem = subItemInfo;
                        }
                    }
                    else
                    {

                        foreach (var subItemUci in subItemUciList)
                        {
                            var subItemInfo = subItemInfos.Find(x => x.SubItemUCI.IsTrimEqualTo(subItemUci, true));

                            var subItemKeyValueList = new MapList();
                            foreach (var subItemKey in subItemKeys)
                            {
                                var subItemKeyMap = new SubKeyMap {Category = subItemKey, KeyValue = String.Empty, CellValue = String.Empty};
                                var subItemKeyCategory = subItemInfo.CategoryValues.ToList().Find(x => x.CategoryName.IsTrimEqualTo(subItemKey, true));
                                if (subItemKeyCategory != null)
                                    subItemKeyMap.KeyValue = subItemKeyCategory.CellDisplayValue.IsNotNullOrEmpty() ? subItemKeyCategory.CellDisplayValue : String.Empty;

                                var existItemKeyCategory = importSubItemCellInfoList.Find(x => x.CategoryName.IsTrimEqualTo(subItemKey, true));
                                if (existItemKeyCategory != null)
                                    subItemKeyMap.CellValue = existItemKeyCategory.CellDisplayValue.IsNotNullOrEmpty() ? existItemKeyCategory.CellDisplayValue : String.Empty;

                                subItemKeyValueList.KeyMaps.Add(subItemKeyMap);
                            }

                            if (!subItemKeyValueList.CheckMap) continue;
                            foundKeyedSubItem = subItemInfo;
                            break;
                        }
                    }

                    //
                    // if we didn't find a matching subitem just add a new one
                    //
                    if (foundKeyedSubItem != null)
                    {
                        doSubItemSync = false;
                        foreach (var subItemCell in foundKeyedSubItem.CategoryValues)
                        {
                            var importSubItemCell = importSubItemCellInfoList.Find(x => x.CategoryName.IsTrimEqualTo(subItemCell.CategoryName, true));
                            if (importSubItemCell == null) continue;
                            if (!subItemCell.CellDisplayValue.IsNotEqualTo(importSubItemCell.CellDisplayValue, false)) continue;
                            subItemCell.CellDisplayValue = importSubItemCell.CellDisplayValue;
                            subItemCell.CellAsOf = DateTime.Now.ToString("MM/dd/yyyy");
                            doSubItemSync = true;
                        }
                    }
                    else
                    {
                        var newSerialNumber = 1;
                        if (subItemInfos.HasItems()) newSerialNumber = subItemInfos.Max(x => x.SubItemSerial) + 1;

                        var newSubItem = new SubItem.psPortfoliosSubItemInfo
                        {
                            SubItemName = Guid.NewGuid().ToString(),
                            SubItemSerial = newSerialNumber,
                            CategoryValues = importSubItemCellInfoList.ToArray()
                        };
                        subItemInfos.Add(newSubItem);
                    }

                    //
                    // Sync the subitems as of today
                    //
                    if (!doSubItemSync)
                    {
                        NLogger.Trace("Update item cells on {0}", itemNameUciValue);
                        continue;
                    }
                    var subItemUpdateStatuses = oppm.SeSubItem.SyncSubItemsAsOf(itemInfo.ProSightID, options.SubItemType, 0, subItemInfos, asOf, false);

                    NLogger.Info("Update item cells on {0}", itemNameUciValue);
                    if (!subItemUpdateStatuses.HasItems()) continue;

                    //
                    // Let see if the subitem sync when well
                    //
                    NLogger.Trace("Logging subItem update status");
                    foreach (var subItemUpdateStatus in subItemUpdateStatuses)
                    {
                        NLogger.Warn("On item {0}, the following subItem {1}.  Error {2}", itemNameUciValue, subItemUpdateStatus.SubItemSerial, subItemUpdateStatus.ErrorText);
                    }

                }
                retVal = true;
            }
            catch (Exception ex)
            {
                NLogger.Fatal(ex.Message);
            }
            return retVal;
        }
    }
}

// ***********************************************************************
// Assembly         : OppmUtility
// Author           : ggreiff
// Created          : 09-01-2014
//
// Last Modified By : ggreiff
// Last Modified On : 09-01-2014
// ***********************************************************************
// <copyright file="Options.cs" company="">
//     Copyright (c) . All rights reserved.
// </copyright>
// <summary></summary>
// ***********************************************************************

using CommandLine;
using CommandLine.Text;

namespace OppmUtility.Utility
{
    /// <summary>
    /// Class Options.
    /// </summary>
    public class Options
    {

        /// <summary>
        /// Gets or sets the name of the XLXS Data file.
        /// </summary>
        /// <value>The name of the XLXS file.</value>
        [Option('x', "xlxsdatafilename", Required = false, HelpText = "The Excel XLXS data filename to use for importing data.")]
        public string XlsxDataFileName { get; set; }

        /// <summary>
        /// Gets or sets the name of the XLXS sheet.
        /// </summary>
        /// <value>The name of the XLXS sheet.</value>
        [Option('d', "dataSheetName", 
            HelpText = "The import worksheet name that contains the data to import. This will default to a sheet named Data if not specified.", DefaultValue = "Data")]
        public string XlsxDataSheetName { get; set; }

        /// <summary>
        /// Get or Set the name of the XLXS Map file
        /// </summary>
        ///  <value>The name of the XLXS file.</value>
        [Option('v', "xlxsmapfilename", Required = false, 
            HelpText = "The Excel XLXS map filename to use for importing data.  This will default to the xlxsdatafilename spreadsheet if not specified.", DefaultValue = "XlsxDataFileName")]
        public string XlsxMapFileName { get; set; }

        /// <summary>
        /// Gets or sets the name of the XLXS sheet.
        /// </summary>
        /// <value>The name of the XLXS sheet.</value>
        [Option('m', "mapSheetName", 
            HelpText = "The import worksheet name that contains the category map values. This will default to a sheet named Map if not specified..", DefaultValue = "Map")]
        public string XlsxMapSheetName { get; set; }

        /// <summary>
        /// Gets or sets the opp server.
        /// </summary>
        /// <value>The opp server.</value>
        [Option('h', "oppmHost", HelpText = "The Oppm host name.")]
        public string OppmHost { get; set; }

        /// <summary>
        /// Gets or sets the name of the oppm user.
        /// </summary>
        /// <value>The name of the oppm user.</value>
        [Option('u', "oppmUser", HelpText = "The Oppm username.")]
        public string OppmUser { get; set; }

        /// <summary>
        /// Gets or sets the oppm password.
        /// </summary>
        /// <value>The oppm password.</value>
        [Option('p', "oppmPassword", HelpText = "The Oppm username's password.")]
        public string OppmPassword { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether [use SSL].
        /// </summary>
        /// <value><c>true</c> if [use SSL]; otherwise, <c>false</c>.</value>
        [Option('l', "useSSL", HelpText = "Use SSL for the web services", DefaultValue = false)]
        public bool UseSsl { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether [use SSL].
        /// </summary>
        /// <value><c>true</c> if [use SSL]; otherwise, <c>false</c>.</value>
        [Option('o', "useCert", HelpText = "Use a client certficate for the web services", DefaultValue = false)]
        public bool UseCert { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="Options"/> is commit.
        /// </summary>
        /// <value><c>true</c> if commit; otherwise, <c>false</c>.</value>
        [Option('c', "commit", HelpText = "Commit the data import otherwise just run a import check on the import spreadsheet.", DefaultValue = false)]
        public bool Commit { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this <see cref="Options"/> is commit.
        /// </summary>
        /// <value><c>true</c> if commit; otherwise, <c>false</c>.</value>
        [Option('n', "newitem", HelpText = "Create new items if the item name in the spreadsheet doesn't exits.", DefaultValue = false)]
        public bool NewItem { get; set; }
        
        /// <summary>
        /// Gets or sets the import portfolio.
        /// </summary>
        /// <value>The import portfolio.</value>
        [Option('i', "importPortfolioName", HelpText = "The name of the portfolio in which to create new items.", DefaultValue = "OPPM IMPORTED ITEMS")]
        public string ImportPortfolio { get; set; }

        /// <summary>
        /// Gets or sets the import portfolio.
        /// </summary>
        /// <value>The import portfolio.</value>
        [Option('s', "subItemType", HelpText = "The name (dynamic list) of the subitem type to import.")]
        public string SubItemType { get; set; }

        /*
        /// <summary>
        /// Gets or sets the name of the category.
        /// </summary>
        /// <value>The name of the category.</value>
        [Option('e', "categoryName", HelpText = "The category name to empty on the import portfolio")]
        public string CategoryName { get; set; }
        */

        /// <summary>
        /// Gets the usage.
        /// </summary>
        /// <returns>System.String.</returns>
        [HelpOption]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this, current => HelpText.DefaultParsingErrorsHandler(this, current));
        }
    }
}

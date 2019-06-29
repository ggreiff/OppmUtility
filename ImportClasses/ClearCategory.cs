using System;
using System.Security.Cryptography.X509Certificates;
using NLog;
using OppmUtility.Utility;

namespace OppmUtility.ImportClasses
{
    class ClearCategory
    {
        public static Logger NLogger = LogManager.GetCurrentClassLogger();

        public X509Certificate Certificate { get; set; }

        public Boolean RunClear(Options options)
        {
            try
            {
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

                var portfolioInfo = oppm.SePorfolio.GetPortfolioInfoByName(options.ImportPortfolio);
                var itemsToClearList = oppm.SePorfolio.GetPortfolioListIdsByCommondId(portfolioInfo.ProSightID);
                NLogger.Info("Clearing category {0} on all items in {1}", options.CategoryName, portfolioInfo.Name);
                foreach (var itemId in itemsToClearList)
                {
                    var itemInfo = oppm.SeItem.GetItemInfo(itemId);
                    oppm.SeCell.UpdateCellValue(itemId, options.CategoryName, String.Empty, 7);
                    NLogger.Info("Cleared {0} on {1}", options.CategoryName, itemInfo.Name);
                }
            }
            catch (Exception ex)
            {
                NLogger.Error(ex.Message);
            }
            return false;
        }
    }
}

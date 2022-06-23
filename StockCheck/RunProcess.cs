using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Text;
using System.Windows;

namespace StockCheck
{
    class RunProcess
    {
        public RunProcess()
        {
            try
            {
                DataModel.errMsg = new StringBuilder();

                // Load default parameters
                string referFile = Helper.getLatestMotherFile();
                string currentFile = Helper.getCurrentFile();

                if(string.IsNullOrEmpty(referFile) || string.IsNullOrEmpty(currentFile))
                {
                    MessageBox.Show("找不到檔名包含[StockDefault]的參考檔案或檔名包含[Stocktake]的作業檔案");
                    Environment.Exit(0);
                }
                // Load reference excel and current excel
                DataTable dtM = Helper.LoadTable(referFile);
                DataTable dtC = Helper.LoadTable(currentFile);

                // Check stock amount
                List<DataTable> dtRL = Function.StockCheck(dtM, dtC);

                // name output excel file
                string storagePath = ConfigurationManager.AppSettings["storagePath"];
                string storageExcelName = $"Stockcheck-{DateTime.Today.Year}-{DateTime.Today.Month}-{DateTime.Today.Day}.xlsx";
                string fullPathName = storagePath + storageExcelName;

                // Output result to excel file
                Helper.DT2Excel(fullPathName, dtRL);

                // Send e-mail
                Helper.SendMailByGmail(@"//郵件內文", fullPathName);

                if (!string.IsNullOrEmpty(DataModel.errMsg.ToString()))
                    MessageBox.Show(DataModel.errMsg.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}

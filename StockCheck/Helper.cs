using System;
using System.Configuration;
using System.IO;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using NPOI.HSSF.UserModel;
using System.Linq;
using System.Net.Mail;

namespace StockCheck
{
    class Helper
    {
        #region get latest mother excel file
        public static string getLatestMotherFile()
        {
            string filePath = ConfigurationManager.AppSettings["motherExcelPath"];
            filePath = filePath.Replace(filePath.Split('\\')[filePath.Split('\\').Length - 1], "");
            string fileKeyword = "StockDefault";

            string latestFile = null;
            foreach (string file in Directory.GetFiles(filePath, "*.*"))
                if (file.Contains(fileKeyword))
                    if (string.Compare(file, latestFile) > 0)
                        latestFile = file;

            return latestFile;
        }
        #endregion

        #region get current excel file
        public static string getCurrentFile()
        {
            string filePath = ConfigurationManager.AppSettings["currentExcelPath"];
            string fileName = filePath.Split('\\').Last();
            filePath = filePath.Replace(filePath.Split('\\')[filePath.Split('\\').Length - 1], "");
            fileName = fileName.Replace("YYYY-MM-DD", $"{DateTime.Today.Year}-{DateTime.Today.Month.ToString("00")}-{DateTime.Today.Day.ToString("00")}");

            foreach (string file in Directory.GetFiles(filePath, "*.*"))
                if (file.Contains(fileName))
                    return filePath + fileName;

            return null;
        }
        #endregion

        #region load table
        public static DataTable LoadTable(string fileName)
        {
            DataTable dt = new DataTable();

            try
            {
                using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.ReadWrite))
                {
                    IWorkbook workbook = null;

                    if (fileName.IndexOf(".xlsx") > 0) // 2007版本 
                    {
                        workbook = new XSSFWorkbook(fileStream); //xlsx數據讀入workbook 
                    }
                    else if (fileName.IndexOf(".xls") > 0) // 2003版本 
                    {
                        workbook = new HSSFWorkbook(fileStream); //xls數據讀入workbook 
                    }

                    ISheet sheet = workbook.GetSheetAt(0);
                    if (sheet.GetRow(0) != null)
                    {
                        IRow row = null;
                        row = sheet.GetRow(0);

                        foreach (ICell cellVal in row.Cells)
                        {
                            dt.Columns.Add(cellVal.ToString());
                        }
                        dt.PrimaryKey = new DataColumn[] { dt.Columns["Barcode"] };
                    }

                    for (int i = 1; i <= sheet.LastRowNum; i++)
                    {
                        DataRow dr = dt.NewRow();
                        IRow row = null;
                        row = sheet.GetRow(i);

                        if (row != null)
                        {
                            for (int j = 0; j < row.LastCellNum; j++)
                            {
                                string cellValue = row.GetCell(j).ToString();
                                dr[j] = cellValue;
                            }
                            if (!string.IsNullOrEmpty(dr["Barcode"].ToString()))
                                dt.Rows.Add(dr);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DataModel.errMsg.AppendLine(ex.Message).ToString();
            }
            return dt;
        }
        #endregion

        #region save table to excel
        public static void DT2Excel(string fileFullName, List<DataTable> dtRL)
        {
            try
            {
                if (dtRL != null)
                {
                    XSSFWorkbook workbook = new XSSFWorkbook();
                    foreach (DataTable dt in dtRL)
                    {
                        if (dt != null && dt.Rows.Count > 0)
                        {
                            ISheet sheet = workbook.CreateSheet(dt.TableName);
                            int rowCount = dt.Rows.Count;
                            int columnCount = dt.Columns.Count;

                            IRow row = sheet.CreateRow(0);
                            ICell cell = null;

                            XSSFColor xssfColor = new XSSFColor();
                            byte[] rgb;

                            #region standard font
                            var font = workbook.CreateFont();
                            font.FontName = "Calibri";
                            font.IsBold = false;
                            #endregion

                            #region exception font
                            var fontEx = workbook.CreateFont();
                            fontEx.FontName = "Calibri";
                            fontEx.IsBold = true;
                            #endregion

                            #region reminder font
                            var fontReminder = workbook.CreateFont();
                            fontReminder.FontName = "Calibri";
                            fontReminder.IsBold = true;
                            fontReminder.Color = IndexedColors.Red.Index;
                            #endregion

                            #region title style
                            XSSFCellStyle titleStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                            titleStyle.FillPattern = FillPattern.SolidForeground;
                            rgb = new byte[] { 142, 169, 219 };
                            xssfColor.SetRgb(rgb);
                            titleStyle.SetFillForegroundColor(xssfColor);
                            #endregion

                            #region standard style
                            XSSFCellStyle stdStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                            #endregion

                            #region default stock exception style
                            XSSFCellStyle defaultStockExStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                            defaultStockExStyle.FillPattern = FillPattern.SolidForeground;
                            xssfColor = new XSSFColor();
                            rgb = new byte[] { 248, 203, 173 };
                            xssfColor.SetRgb(rgb);
                            defaultStockExStyle.SetFillForegroundColor(xssfColor);
                            #endregion

                            #region Stock balance exception style
                            XSSFCellStyle outOfStockExStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                            outOfStockExStyle.FillPattern = FillPattern.SolidForeground;
                            xssfColor = new XSSFColor();
                            rgb = new byte[] { 255, 75, 33 };
                            xssfColor.SetRgb(rgb);
                            outOfStockExStyle.SetFillForegroundColor(xssfColor);
                            #endregion

                            #region reminder style
                            XSSFCellStyle reminderStyle = (XSSFCellStyle)workbook.CreateCellStyle();
                            reminderStyle.FillPattern = FillPattern.SolidForeground;
                            xssfColor = new XSSFColor();
                            rgb = new byte[] { 255, 255, 0 };
                            xssfColor.SetRgb(rgb);
                            reminderStyle.SetFillForegroundColor(xssfColor);
                            #endregion

                            for (int c = 0; c < columnCount; c++)
                            {
                                cell = row.CreateCell(c);
                                cell.SetCellValue(dt.Columns[c].ColumnName);
                                cell.CellStyle = titleStyle;
                                cell.CellStyle.SetFont(font);
                            }

                            for (int i = 0; i < rowCount; i++)
                            {
                                row = sheet.CreateRow(i + 1);
                                for (int j = 0; j < columnCount; j++)
                                {
                                    cell = row.CreateCell(j);
                                    cell.SetCellValue(dt.Rows[i][j].ToString());
                                    if (Int32.Parse(dt.Rows[i][2].ToString()) == 0)
                                    {
                                        if (j == 2)
                                            cell.CellStyle = defaultStockExStyle;
                                    }

                                    if (Int32.Parse(dt.Rows[i][4].ToString()) < 0)
                                    {
                                        if (j == 4)
                                            cell.CellStyle = outOfStockExStyle;
                                        else
                                            cell.CellStyle = stdStyle;

                                        if (dt.TableName != "StockException")
                                            cell.CellStyle.SetFont(fontEx);
                                    }
                                    else
                                        cell.CellStyle.SetFont(font);
                                }
                            }

                            // 自動列寬
                            int numberOfColumns = sheet.GetRow(0).PhysicalNumberOfCells;
                            for (int i = 0; i < numberOfColumns; i++)
                            {
                                sheet.AutoSizeColumn(i);
                            }

                            if (dt.TableName == "StockException")
                            {
                                row = sheet.CreateRow(rowCount + 2);
                                cell = row.CreateCell(0);
                                cell.SetCellValue(@"Please check ""囤貨數量參考表""的更新狀況。");
                                cell.CellStyle.SetFont(fontReminder);
                                cell.CellStyle = reminderStyle;
                                cell = row.CreateCell(1);
                                cell.CellStyle = reminderStyle;
                                cell.CellStyle.SetFont(fontReminder);
                            }
                        }
                    }
                    using (FileStream fs = new FileStream(fileFullName, FileMode.Create, FileAccess.ReadWrite))
                    {
                        workbook.Write(fs);
                    }
                }
            }
            catch (Exception ex)
            {
                DataModel.errMsg.AppendLine(ex.Message).ToString();
            }
        }
        #endregion

        #region send e-mail
        public static void SendMailByGmail(string body, string attach)
        {
            MailMessage msg = new MailMessage();

            try
            {
                List<string> receivers = ConfigurationManager.AppSettings["emailReceiver"].Split(';').ToList();
                string subject = ConfigurationManager.AppSettings["emailSubject"];
                string senderAcc = ConfigurationManager.AppSettings["emailAcc"];
                string pwd = ConfigurationManager.AppSettings["emailPwd"];

                //收件者，以分號分隔不同收件者 ex "test@gmail.com,test2@gmail.com"
                msg.To.Add(string.Join(";", receivers.ToArray()));
                msg.From = new MailAddress(senderAcc, "【AutoSystem】HibicusPharmacy", System.Text.Encoding.UTF8);
                //郵件標題
                msg.Subject = subject;
                //郵件標題編碼
                msg.SubjectEncoding = System.Text.Encoding.UTF8;
                //郵件內容
                msg.Body = body;
                msg.IsBodyHtml = true;
                msg.BodyEncoding = System.Text.Encoding.UTF8;//郵件內容編碼
                msg.Priority = MailPriority.Normal;//郵件優先級

                if (attach != null)
                    msg.Attachments.Add(new Attachment(attach));

                                                   //建立 SmtpClient 物件 並設定 Gmail的smtp主機及Port
                #region 其它 Host
                /*
                 *  outlook.com smtp.live.com port:25
                 *  yahoo smtp.mail.yahoo.com.tw port:465
                */
                #endregion

                using (SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587))
                {
                    //設定你的帳號密碼
                    smtp.Credentials = new System.Net.NetworkCredential(senderAcc, pwd);
                    //Gmial 的 smtp 使用 SSL
                    smtp.EnableSsl = true;
                    //smtp.UseDefaultCredentials = true;
                    smtp.Send(msg);
                }
            }
            catch(Exception ex)
            {
                DataModel.errMsg.AppendLine(ex.Message).ToString();
            }
        }
        #endregion
    }
}

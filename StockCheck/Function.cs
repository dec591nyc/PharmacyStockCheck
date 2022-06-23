using System;
using System.Collections.Generic;
using System.Data;

namespace StockCheck
{
    class Function
    {
        public Function()
        {

        }

        public static List<DataTable> StockCheck(DataTable dtM, DataTable dtC)
        {
            List<DataTable> result = new List<DataTable>();

            try
            {
                DataTable dtR = dtC;
                dtR.TableName = "StockResult";
                DataTable dtTmp;
                dtR.Columns.RemoveAt(7);
                dtR.Columns.RemoveAt(6);
                dtR.Columns.RemoveAt(5);
                dtR.Columns.RemoveAt(2);
                dtR.Columns.RemoveAt(1);
                dtTmp = dtM;
                dtR.Columns.Add("Stock Default").SetOrdinal(2);
                dtR.Columns.Add("Stock balance");

                bool isMomLargerSon = false;
                if (dtM.Rows.Count > dtC.Rows.Count)
                    isMomLargerSon = true;

                DataTable lostDT = new DataTable("StockException");

                foreach (DataColumn col in dtR.Columns)
                    lostDT.Columns.Add(col.ColumnName, col.DataType);
                lostDT.PrimaryKey = new DataColumn[] { lostDT.Columns["Barcode"] };

                int stockDefault, stockOnHand;
                foreach (DataRow drR in dtR.Rows)
                {
                    var barcode = drR["Barcode"].ToString();

                    if (dtTmp.Rows.Contains(drR["Barcode"].ToString()))
                    {
                        DataRow foundRow = dtTmp.Rows.Find(barcode);
                        drR["Stock Default"] = foundRow["Stock Default"];
                        stockDefault = Int32.Parse(drR["Stock Default"].ToString());
                        stockOnHand = Int32.Parse(drR["Stock on Hand"].ToString());
                        drR["Stock balance"] = (stockOnHand - stockDefault).ToString();
                    }
                    else
                    {
                        drR["Stock Default"] = 0;
                        stockDefault = Int32.Parse(drR["Stock Default"].ToString());
                        stockOnHand = Int32.Parse(drR["Stock on Hand"].ToString());
                        drR["Stock balance"] = (stockOnHand - stockDefault).ToString();

                        // sheet: Exception add one new row.
                        DataRow dr = lostDT.NewRow();
                        dr["Description"] = drR["Description"];
                        dr["Barcode"] = barcode;
                        dr["Stock Default"] = 0;
                        dr["Stock on Hand"] = drR["Stock on Hand"];
                        stockDefault = Int32.Parse(dr["Stock Default"].ToString());
                        stockOnHand = Int32.Parse(dr["Stock on Hand"].ToString());
                        dr["Stock balance"] = (stockOnHand - stockDefault).ToString();
                        lostDT.Rows.Add(dr);
                    }
                }

                foreach (DataRow drT in dtTmp.Rows)
                {
                    var barcode = drT["Barcode"].ToString();

                    if (!dtR.Rows.Contains(barcode))
                    {
                        // sheet: Exception add one new row.
                        DataRow dr = lostDT.NewRow();
                        dr["Description"] = drT["Description"];
                        dr["Barcode"] = barcode;
                        dr["Stock Default"] = drT["Stock Default"];
                        dr["Stock on Hand"] = 0;
                        stockDefault = Int32.Parse(dr["Stock Default"].ToString());
                        stockOnHand = Int32.Parse(dr["Stock on Hand"].ToString());
                        dr["Stock balance"] = (stockOnHand - stockDefault).ToString();
                        lostDT.Rows.Add(dr);
                    }
                }

                result.Add(dtR);
                if (lostDT.Rows.Count > 0)
                    result.Add(lostDT);
            }
            catch (Exception ex)
            {
                DataModel.errMsg.AppendLine(ex.Message).ToString();
            }
            return result;
        }
    }
}

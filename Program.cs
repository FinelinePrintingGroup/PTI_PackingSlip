using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using System.Net;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Collections;
using System.Diagnostics;

namespace PTI_Integration_Console
{
    class Program
    {
        static void Main(string[] args)
        {
            XmlDocument queryDoc = new XmlDocument();
            XmlDocument reqDoc = new XmlDocument();
            XmlDocument respDoc = new XmlDocument();
            

            methods me = new methods();
            ArrayList al = new ArrayList();
            ArrayList fileInfo = new ArrayList();
            ArrayList shipInfo = new ArrayList();
            bool test = false;
            string ShipmentNumber = "";
            string ShipmentLineNumber = "";
            string PTI_lineItemID = "";

            me.preProcessNonPTI();
            al = me.getShipmentItemsForProcessing();

            //Remove id's for non-finalshipyn items
            foreach (int row in al)
            {
                shipInfo = me.getShipmentInfo(row);

                test = me.checkFinalShipyn(Convert.ToInt32(shipInfo[0].ToString()), Convert.ToInt32(shipInfo[1].ToString()));

                if (!test)
                {
                    al.Remove(row.ToString());
                }
            }

            foreach (int id in me.processDoubleCheck())
            {
                if (!al.Contains(id))
                    al.Add(id);
            }

            foreach (int column in al)
            {
                queryDoc = me.queryDatabase(column.ToString());
                
                Console.WriteLine("\n");
                me.CreatePackingSlipByLineItem(queryDoc).Save(Console.Out);

                reqDoc = me.CreatePackingSlipByLineItem(queryDoc);
                
                respDoc = me.sendXmlRequest(reqDoc);
                me.respCreatePackingSlipByLineItem(respDoc);

                try
                {
                    //fileInfo = me.get_fileInfo(column.ToString());
                    ShipmentNumber = "ShipNum"; 
                    //ShipmentNumber = fileInfo[0].ToString();
                    ShipmentLineNumber = column.ToString();
                    PTI_lineItemID = column.ToString();
                    //ShipmentLineNumber = fileInfo[1].ToString();
                    //PTI_lineItemID = fileInfo[2].ToString();

                    reqDoc.Save("../../XML/" + column.ToString() + "_req.xml");
                    respDoc.Save("../../XML/" + column.ToString() + "_resp.xml");
                    me.responseError("1", "File Write", "File Created", PTI_lineItemID, ShipmentNumber + " - " + ShipmentLineNumber);
                }
                catch (Exception e)
                {
                    me.responseError("0", "File Write", e.ToString(), PTI_lineItemID, ShipmentNumber + " - " + ShipmentLineNumber);
                    Console.WriteLine(e.ToString());
                }

                me.updatePrintableProcessed(column.ToString());
                Console.WriteLine("");
                Console.WriteLine("=========================================\n");
            }
            

                    
                    
        }
    }
}

using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Xml;
using System.IO;
using System.Net;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;

namespace PTI_Integration_Console
{
    class methods
    {
        public ArrayList get_fileInfo(string eod_ID)
        {
            string queryGetIDs = @"SELECT shipmentNumber, shipmentLineNum, OrderDetail_ID
                                   FROM vw_PTI_integration
                                   WHERE eod_id = " + eod_ID;

            ArrayList fileInfo = new ArrayList();

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(queryGetIDs, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            object[] values = new object[reader.FieldCount];
                            reader.GetValues(values);
                            fileInfo.Add(values);
                        }

                        reader.Close();
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            return fileInfo;
        }

        public void preProcessNonPTI()
        {
            string queryNonPTI = @"update CT_UPS_EODitems
                                   set PTI_processed = 1
                                   where PTI_lineItemID = '0'
                                   or (pti_lineItemID is null
                                       and (not Logic_FGNum > 0
                                            or Logic_FGNum = ''
                                            or Logic_FGNum is null))";


            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(queryNonPTI, conn);
                    try
                    {
                        int rowsUpdated = 0;
                        command.Connection.Open();
                        rowsUpdated = command.ExecuteNonQuery();
                        command.Dispose();
                        command = null;
                        Console.WriteLine("========================================");
                        Console.WriteLine("=    NON-PRINTABLE ORDERS PROCESSED    =");
                        Console.WriteLine("========================================");
                        Console.WriteLine("");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        public ArrayList getShipmentItemsForProcessing()
        {
            Console.WriteLine("=========================================");
            Console.WriteLine("= GATHERING TODAY'S PRINTABLE SHIPMENTS =");
            Console.WriteLine("=========================================");

            ArrayList al = new ArrayList();
            al.Clear();

            string queryGetIDs = @"select id 
                                   from CT_UPS_EODitems 
                                   where PTI_processed = 0 
                                   and (PTI_lineItemID > 0
                                        or Logic_FGNum > 0)";

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(queryGetIDs, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            if (!al.Contains(Convert.ToInt32(reader[0].ToString())))
                                al.Add(Convert.ToInt32(reader[0].ToString()));
                        }

                        reader.Close();
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            return al;
        }

        public XmlDocument queryDatabase(string pID)
        {
            string id = pID;
            string queryGetInfo = @"select *
                                   from vw_PTI_integration
                                   where ShipmentNumber = (select shipmentNum 
                                                           from CT_UPS_EODitems
                                                           where id = " + id + @")
                                   and ShipmentLineNum = (select shipmentLineNum
                                                             from CT_UPS_EODitems
                                                             where id = " + id + @")";

            //Console.WriteLine(queryGetInfo);

            Console.WriteLine("=========================================");
            Console.WriteLine("= QUERYING THE DATABASE FOR ID#: " + id + "  =");
            Console.WriteLine("=========================================");

            XmlDocument doc = new XmlDocument();
            doc.LoadXml("<queryReturn></queryReturn>");

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(queryGetInfo, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                XmlElement el = doc.CreateElement(reader.GetName(i).ToString());
                                XmlText txt = doc.CreateTextNode(reader.GetValue(i).ToString());
                                doc.DocumentElement.AppendChild(el);
                                doc.DocumentElement.LastChild.AppendChild(txt);
                            }
                        }
                        XmlDeclaration xmlDecl = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
                        XmlElement root = doc.DocumentElement;
                        doc.InsertBefore(xmlDecl, root);

                        doc.Save(Console.Out);

                        reader.Close();
                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            return doc;
        }

        public ArrayList getShipmentInfo(int eod_ID)
        {
            ArrayList al = new ArrayList();

            string q = @"SELECT ShipmentNum, ShipmentLineNum
                         FROM printable.dbo.CT_UPS_EODitems
                         WHERE id = " + eod_ID;

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            al.Add(Convert.ToInt32(reader[0].ToString()));
                            al.Add(Convert.ToInt32(reader[1].ToString()));
                        }

                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            return al;
        }

        public bool checkFinalShipyn(int shipmentNum, int shipmentLineNum)
        {
            int ret = 0;

            string q = @"SELECT ISNULL(FinalShipyn, 0)
                         FROM printable.dbo.vw_SHIP_FinalShipynTest
                         WHERE ShipmentNumber = " + shipmentNum + @"
                         AND LineN = " + shipmentLineNum;

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            ret = Convert.ToInt32(reader[0].ToString());
                        }

                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            if (ret > 0)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public ArrayList processDoubleCheck()
        {
            ArrayList temp = new ArrayList();

            string q = @"SELECT id
                        FROM printable.dbo.CT_UPS_EODitems
                        WHERE PTI_lineItemID IN (SELECT DISTINCT printableID
				                                FROM printable.dbo.CT_PTI_ErrorLog
				                                WHERE errNum LIKE 'PS-167%'
				                                AND printableID != '-1'
				                                AND doubleCheck = 0)";
            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        SqlDataReader reader = command.ExecuteReader();

                        while (reader.Read())
                        {
                            temp.Add(Convert.ToInt32(reader[0].ToString()));
                        }

                        conn.Close();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            updateDoubleCheck();

            return temp;            
        }

        protected void updateDoubleCheck()
        {
            int rowsUpdated = 0;

            string q = @"UPDATE printable.dbo.CT_PTI_ErrorLog
                        SET doubleCheck = 1
                        WHERE errNum LIKE 'PS-167%'
                        AND printableID != '-1'
                        AND doubleCheck = 0";

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_logicConnString))
                {
                    SqlCommand command = new SqlCommand(q, conn);
                    try
                    {
                        command.Connection.Open();
                        rowsUpdated = command.ExecuteNonQuery();
                        command.Dispose();
                        command = null;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            Console.WriteLine(rowsUpdated + " rows updated via DoubleCheck");
        }

        /*public XmlDocument CreatePackingSlipByOrder(XmlDocument pDoc)
        {
            XmlDocument inDoc = pDoc;
            XmlDocument outDoc = new XmlDocument();

            try
            {
                #region DECLARATIONS
                //XMLNODELIST CARRIES THE ENTIRE TAG, USE [0].INNERTEXT TO PULL OUT VALUES
                string token = Globals.get_printableToken;
                XmlNodeList packingSlipNum = inDoc.GetElementsByTagName("ShipmentNumber");
                XmlNodeList shipDate = inDoc.GetElementsByTagName("ShipDate");
                XmlNodeList carrierName = inDoc.GetElementsByTagName("CarrierName");
                XmlNodeList trackingNum = inDoc.GetElementsByTagName("trackingNum"); 
                //XmlNodeList shipCost = inDoc.GetElementsByTagName("CostOnShipmnt");
                XmlNodeList ship_atten = inDoc.GetElementsByTagName("ShiptoAttn");
                XmlNodeList ship_name = inDoc.GetElementsByTagName("Addressee");
                XmlNodeList ship_add1 = inDoc.GetElementsByTagName("AddrLine1");
                XmlNodeList ship_add2 = inDoc.GetElementsByTagName("AddrLine2");
                XmlNodeList ship_add3 = inDoc.GetElementsByTagName("AddrLine3");
                XmlNodeList ship_city = inDoc.GetElementsByTagName("City");
                XmlNodeList ship_state = inDoc.GetElementsByTagName("StateProv");
                XmlNodeList ship_postalCode = inDoc.GetElementsByTagName("PostalCode");
                XmlNodeList ship_country = inDoc.GetElementsByTagName("CountryCode");
                XmlNodeList printable_id = inDoc.GetElementsByTagName("PrintableID");
                #endregion

                //PREFIX DECLARATIONS
                string soapPrefix = "soapenv";
                string ptiPrefix = "pac";
                string soapNamespace = @"http://schemas.xmlsoap.org/soap/envelope/";
                string ptiNamespace = @"http://www.printable.com/WebService/PackingSlip";

                //SOAP ENVELOPE CREATION
                XmlElement root = outDoc.CreateElement(soapPrefix, "Envelope", soapNamespace);
                root.SetAttribute("xmlns:soapenv", soapNamespace);
                root.SetAttribute("xmlns:pac", ptiNamespace);
                outDoc.AppendChild(root);

                //SOAP EMPTY HEADER CREATION
                XmlElement header = outDoc.CreateElement(soapPrefix, "Header", soapNamespace);
                root.AppendChild(header);

                #region START SOAP BODY CREATION
                XmlElement body = outDoc.CreateElement(soapPrefix, "Body", soapNamespace);

                XmlElement CreatePackingSlipByOrder = outDoc.CreateElement(ptiPrefix, "CreatePackingSlipByOrder", ptiNamespace);
                body.AppendChild(CreatePackingSlipByOrder);

                XmlElement pRequest = outDoc.CreateElement(ptiPrefix, "pRequest", ptiNamespace);
                CreatePackingSlipByOrder.AppendChild(pRequest);

                #region START PARTNER CREDENTIALS BLOCK
                XmlElement partnerCredentials = outDoc.CreateElement("PartnerCredentials");
                pRequest.AppendChild(partnerCredentials);

                XmlElement tokenTag = outDoc.CreateElement("Token");
                XmlText txt = outDoc.CreateTextNode(token);
                tokenTag.AppendChild(txt);
                partnerCredentials.AppendChild(tokenTag);
                #endregion END PARTNER CREDENTIALS BLOCK

                #region START PACKING SLIP NODE BLOCK
                XmlElement packingSlipNode = outDoc.CreateElement("PackingSlipNode");
                pRequest.AppendChild(packingSlipNode);

                XmlElement packingSlipNumberTag = outDoc.CreateElement("PackingSlipNumber");
                txt = outDoc.CreateTextNode(packingSlipNum[0].InnerText);
                packingSlipNumberTag.AppendChild(txt);
                packingSlipNode.AppendChild(packingSlipNumberTag);

                XmlElement shipDateTag = outDoc.CreateElement("ShipDate");
                txt = outDoc.CreateTextNode(shipDate[0].InnerText);
                shipDateTag.AppendChild(txt);
                packingSlipNode.AppendChild(shipDateTag);

                XmlElement carrierNameTag = outDoc.CreateElement("CarrierName");
                txt = outDoc.CreateTextNode(carrierName[0].InnerText);
                carrierNameTag.AppendChild(txt);
                packingSlipNode.AppendChild(carrierNameTag);

                XmlElement trackingNumberTag = outDoc.CreateElement("TrackingNumber");
                txt = outDoc.CreateTextNode(trackingNum[0].InnerText);
                trackingNumberTag.AppendChild(txt);
                packingSlipNode.AppendChild(trackingNumberTag);

                //SHIP COST IS ALREADY ONLINE

                XmlElement shipToTag = outDoc.CreateElement("ShipToAddress");
                packingSlipNode.AppendChild(shipToTag);

                if (ship_atten[0].InnerText.Length > 0)
                {
                    XmlElement attentionTag = outDoc.CreateElement("Attention");
                    txt = outDoc.CreateTextNode(ship_atten[0].InnerText);
                    attentionTag.AppendChild(txt);
                    shipToTag.AppendChild(attentionTag);
                }

                XmlElement nameTag = outDoc.CreateElement("Name");
                txt = outDoc.CreateTextNode(ship_name[0].InnerText);
                nameTag.AppendChild(txt);
                shipToTag.AppendChild(nameTag);

                XmlElement addr1Tag = outDoc.CreateElement("Address1");
                txt = outDoc.CreateTextNode(ship_add1[0].InnerText);
                addr1Tag.AppendChild(txt);
                shipToTag.AppendChild(addr1Tag);

                if (ship_add2[0].InnerText.Length > 0)
                {
                    XmlElement addr2Tag = outDoc.CreateElement("Address2");
                    txt = outDoc.CreateTextNode(ship_add2[0].InnerText);
                    addr2Tag.AppendChild(txt);
                    shipToTag.AppendChild(addr2Tag);
                }


                if (ship_add3[0].InnerText.Length > 0)
                {
                    XmlElement addr3Tag = outDoc.CreateElement("Address3");
                    txt = outDoc.CreateTextNode(ship_add3[0].InnerText);
                    addr3Tag.AppendChild(txt);
                    shipToTag.AppendChild(addr3Tag);
                }

                XmlElement cityTag = outDoc.CreateElement("City");
                txt = outDoc.CreateTextNode(ship_city[0].InnerText);
                cityTag.AppendChild(txt);
                shipToTag.AppendChild(cityTag);

                XmlElement stateTag = outDoc.CreateElement("State");
                txt = outDoc.CreateTextNode(ship_state[0].InnerText);
                stateTag.AppendChild(txt);
                shipToTag.AppendChild(stateTag);

                XmlElement postalTag = outDoc.CreateElement("PostalCode");
                txt = outDoc.CreateTextNode(ship_postalCode[0].InnerText);
                postalTag.AppendChild(txt);
                shipToTag.AppendChild(postalTag);

                XmlElement countryTag = outDoc.CreateElement("Country");
                txt = outDoc.CreateTextNode(ship_country[0].InnerText);
                countryTag.AppendChild(txt);
                shipToTag.AppendChild(countryTag);
                #endregion END PACKING SLIP NODE BLOCK

                #region START ORDERS BLOCK
                XmlElement orders = outDoc.CreateElement("Orders");
                pRequest.AppendChild(orders);

                XmlElement idTag = outDoc.CreateElement("ID");
                idTag.SetAttribute("type", "Printable");
                txt = outDoc.CreateTextNode(printable_id[0].InnerText);
                idTag.AppendChild(txt);
                orders.AppendChild(idTag);
                #endregion END ORDERS BLOCK
                #endregion

                root.AppendChild(body);
                return outDoc;
            }

            catch (Exception e)
            {
                Console.WriteLine("Error: " + e.ToString());
                return outDoc;
            }
        }*/

        public XmlDocument CreatePackingSlipByLineItem(XmlDocument pDoc)
        {
            XmlDocument inDoc = pDoc;
            XmlDocument outDoc = new XmlDocument();

            try
            {
                #region DECLARATIONS
                //XMLNODELIST CARRIES THE ENTIRE TAG, USE [0].INNERTEXT TO PULL OUT VALUES
                XmlNodeList packingSlipNum = inDoc.GetElementsByTagName("ShipmentNumber");
                string pSlipNum = packingSlipNum.Item(0).InnerText;
                XmlNodeList shipDate = inDoc.GetElementsByTagName("ShipDate");
                XmlNodeList carrierName = inDoc.GetElementsByTagName("CarrierName");
                XmlNodeList trackingNum = inDoc.GetElementsByTagName("trackingNum");
                XmlNodeList ship_atten = inDoc.GetElementsByTagName("ShiptoAttn");
                XmlNodeList ship_name = inDoc.GetElementsByTagName("Addressee");
                XmlNodeList ship_add1 = inDoc.GetElementsByTagName("AddrLine1");
                XmlNodeList ship_add2 = inDoc.GetElementsByTagName("AddrLine2");
                XmlNodeList ship_add3 = inDoc.GetElementsByTagName("AddrLine3");
                XmlNodeList ship_city = inDoc.GetElementsByTagName("City");
                XmlNodeList ship_state = inDoc.GetElementsByTagName("StateProv");
                XmlNodeList ship_postalCode = inDoc.GetElementsByTagName("PostalCode");
                XmlNodeList ship_country = inDoc.GetElementsByTagName("CountryCode");
                XmlNodeList order_id = inDoc.GetElementsByTagName("Order_ID");
                XmlNodeList orderDetail_id = inDoc.GetElementsByTagName("OrderDetail_ID");
                XmlNodeList orderDetail_id2 = inDoc.GetElementsByTagName("Logic_FGNum");
                XmlNodeList ship_qty = inDoc.GetElementsByTagName("Quantity");
                XmlNodeList fg_itemNum = inDoc.GetElementsByTagName("FGItemNum");

                if (orderDetail_id[0].InnerText.Length == 0)
                {
                    orderDetail_id = orderDetail_id2;
                }

                #endregion

                //PREFIX DECLARATIONS
                string soapPrefix = "soapenv";
                string ptiPrefix = "pac";
                string soapNamespace = @"http://schemas.xmlsoap.org/soap/envelope/";
                string ptiNamespace = @"http://www.printable.com/WebService/PackingSlip";

                //SOAP ENVELOPE CREATION
                XmlElement root = outDoc.CreateElement(soapPrefix, "Envelope", soapNamespace);
                root.SetAttribute("xmlns:soapenv", soapNamespace);
                root.SetAttribute("xmlns:pac", ptiNamespace);
                outDoc.AppendChild(root);

                //SOAP EMPTY HEADER CREATION
                XmlElement header = outDoc.CreateElement(soapPrefix, "Header", soapNamespace);
                root.AppendChild(header);

                //SOAP BODY CREATION
                #region START SOAP BODY CREATION
                XmlElement body = outDoc.CreateElement(soapPrefix, "Body", soapNamespace);

                XmlElement CreatePackingSlipByLineItem = outDoc.CreateElement(ptiPrefix, "CreatePackingSlipByLineItem", ptiNamespace);
                body.AppendChild(CreatePackingSlipByLineItem);

                XmlElement pRequest = outDoc.CreateElement(ptiPrefix, "pRequest", ptiNamespace);
                CreatePackingSlipByLineItem.AppendChild(pRequest);

                #region START PARTNER CREDENTIALS BLOCK
                XmlElement partnerCredentials = outDoc.CreateElement("PartnerCredentials");
                pRequest.AppendChild(partnerCredentials);

                XmlElement tokenTag = outDoc.CreateElement("Token");
                XmlText txt = outDoc.CreateTextNode(Globals.get_printableToken);
                tokenTag.AppendChild(txt);
                partnerCredentials.AppendChild(tokenTag);
                #endregion END PARTNER CREDENTIALS BLOCK

                #region START PACKING SLIP NODE BLOCK
                XmlElement packingSlipNode = outDoc.CreateElement("PackingSlipNode");
                pRequest.AppendChild(packingSlipNode);

                XmlElement packingSlipNumberTag = outDoc.CreateElement("PackingSlipNumber");
                txt = outDoc.CreateTextNode(packingSlipNum[0].InnerText);
                packingSlipNumberTag.AppendChild(txt);
                packingSlipNode.AppendChild(packingSlipNumberTag);

                XmlElement shipDateTag = outDoc.CreateElement("ShipDate");
                txt = outDoc.CreateTextNode(shipDate[0].InnerText);
                shipDateTag.AppendChild(txt);
                packingSlipNode.AppendChild(shipDateTag);

                XmlElement carrierNameTag = outDoc.CreateElement("CarrierName");
                txt = outDoc.CreateTextNode(carrierName[0].InnerText);
                carrierNameTag.AppendChild(txt);
                packingSlipNode.AppendChild(carrierNameTag);

                XmlElement trackingNumberTag = outDoc.CreateElement("TrackingNumber");
                txt = outDoc.CreateTextNode(trackingNum[0].InnerText);
                trackingNumberTag.AppendChild(txt);
                packingSlipNode.AppendChild(trackingNumberTag);

                //NO SHIP COST INCLUDED BECAUSE IT'S ALREADY THERE

                XmlElement adjustInventoryTag = outDoc.CreateElement("SkipAdjustInventory");
                txt = outDoc.CreateTextNode("false"); // Changed to false because it wasn't adjusting FG Order inventories
                adjustInventoryTag.AppendChild(txt);
                packingSlipNode.AppendChild(adjustInventoryTag);

                XmlElement shipToTag = outDoc.CreateElement("ShipToAddress");
                packingSlipNode.AppendChild(shipToTag);

                if (ship_atten[0].InnerText.Length > 0)
                {
                    XmlElement attentionTag = outDoc.CreateElement("Attention");
                    txt = outDoc.CreateTextNode(ship_atten[0].InnerText);
                    attentionTag.AppendChild(txt);
                    shipToTag.AppendChild(attentionTag);
                }

                XmlElement nameTag = outDoc.CreateElement("Name");
                txt = outDoc.CreateTextNode(ship_name[0].InnerText);
                nameTag.AppendChild(txt);
                shipToTag.AppendChild(nameTag);

                XmlElement addr1Tag = outDoc.CreateElement("Address1");
                txt = outDoc.CreateTextNode(ship_add1[0].InnerText);
                addr1Tag.AppendChild(txt);
                shipToTag.AppendChild(addr1Tag);

                if (ship_add2[0].InnerText.Length > 0)
                {
                    XmlElement addr2Tag = outDoc.CreateElement("Address2");
                    txt = outDoc.CreateTextNode(ship_add2[0].InnerText);
                    addr2Tag.AppendChild(txt);
                    shipToTag.AppendChild(addr2Tag);
                }


                if (ship_add3[0].InnerText.Length > 0)
                {
                    XmlElement addr3Tag = outDoc.CreateElement("Address3");
                    txt = outDoc.CreateTextNode(ship_add3[0].InnerText);
                    addr3Tag.AppendChild(txt);
                    shipToTag.AppendChild(addr3Tag);
                }

                XmlElement cityTag = outDoc.CreateElement("City");
                txt = outDoc.CreateTextNode(ship_city[0].InnerText);
                cityTag.AppendChild(txt);
                shipToTag.AppendChild(cityTag);

                XmlElement stateTag = outDoc.CreateElement("State");
                txt = outDoc.CreateTextNode(ship_state[0].InnerText);
                stateTag.AppendChild(txt);
                shipToTag.AppendChild(stateTag);

                XmlElement postalTag = outDoc.CreateElement("PostalCode");
                txt = outDoc.CreateTextNode(ship_postalCode[0].InnerText);
                postalTag.AppendChild(txt);
                shipToTag.AppendChild(postalTag);

                XmlElement countryTag = outDoc.CreateElement("Country");
                txt = outDoc.CreateTextNode(ship_country[0].InnerText);
                countryTag.AppendChild(txt);
                shipToTag.AppendChild(countryTag);
                #endregion END PACKING SLIP NODE BLOCK

                #region START LINE ITEMS BLOCK
                               
                XmlElement lineItems = outDoc.CreateElement("LineItems");
                pRequest.AppendChild(lineItems);

                XmlElement lineItem = outDoc.CreateElement("LineItem");
                lineItems.AppendChild(lineItem);

                XmlElement idTag = outDoc.CreateElement("ID");
                idTag.SetAttribute("type", "Printable");
                txt = outDoc.CreateTextNode(orderDetail_id[0].InnerText);
                idTag.AppendChild(txt);
                lineItem.AppendChild(idTag);

                if (Convert.ToInt32(fg_itemNum[0].InnerText) > 0)
                {
                    XmlElement qtyTag = outDoc.CreateElement("Quantity");
                    txt = outDoc.CreateTextNode(ship_qty[0].InnerText);
                    qtyTag.AppendChild(txt);
                    lineItem.AppendChild(qtyTag);
                }

                #endregion END ORDERS BLOCK
                #endregion

                root.AppendChild(body);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
                        
            return outDoc;
        }

        #region OTHER REQUEST FORMATS
        //public XmlDocument CreatePackingSlipByWorkOrder(string PK)
        //{
        //}

        //public XmlDocument CreateUpdatePackingSlip(XmlDocument DOC)
        //{
        //}
        #endregion



        public void respCreatePackingSlipByOrder(XmlDocument pDoc)
        {
            XmlDocument doc = pDoc;

            string errNum = "";
            string errDesc = "";
            string errStatus = "";
            string printableID = "";

            XmlNodeList respStatus = doc.GetElementsByTagName("Status");

            try
            {
                for (int i = 0; i < respStatus.Count; i++)
                {
                    errNum = respStatus[i].Attributes["Code"].Value;
                    errDesc = respStatus[i].Attributes["Message"].Value;
                    errStatus = respStatus[i].Attributes["Status"].Value;
                    printableID = "";
                }

                responseError(errNum, errStatus, errDesc, printableID, "");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }

        public void respCreatePackingSlipByLineItem(XmlDocument pDoc)
        {
            XmlDocument doc = pDoc;

            doc.Save(Console.Out);

            string errNum = "";
            string errDesc = "";
            string errStatus = "";
            string printableID = "";
            string ps_num = "";

            try
            {
                XmlNodeList respStatus = doc.GetElementsByTagName("Status");
                XmlNodeList respID = doc.GetElementsByTagName("LineItemID");
                XmlNodeList respAction = doc.GetElementsByTagName("Action");

                for (int i = 0; i < respStatus.Count; i++)
                {
                    errNum = respStatus[i].Attributes["Code"].Value;
                    errDesc = respStatus[i].Attributes["Message"].Value;
                    errStatus = respStatus[i].Attributes["Status"].Value;
                    printableID = respID[0].InnerText.ToString();

                    try
                    {
                        ps_num = respAction[0].Attributes["ReferenceId"].Value;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                        ps_num = "Exception";
                    }

                    if ((printableID == "") | (printableID == null))
                    {
                        printableID = "No Identifier";
                    }

                    if ((ps_num == "") | (ps_num == null))
                    {
                        ps_num = "No identifier";
                    }
                    else if (ps_num == printableID)
                    {
                        ps_num = "x";
                    }

                    if (errNum == "167")
                    {
                        Globals globals = new Globals();
                        globals.sendEmail("ShawnH@Finelink.com", "PTI Inventory Conflict<br/>" + printableID + "<br/>" + errDesc +"<br/>" + ps_num);
                        globals.sendEmail("GVreeman@Finelink.com", "PTI Inventory Conflict<br/>" + printableID + "<br/>" + errDesc + "<br/>" + ps_num);
                    }

                }

                responseError(errNum, errStatus, errDesc, printableID, ps_num);
            }
            catch (Exception e)
            {
                responseError("0", "Parse Read Error", e.ToString(), printableID, "");
            }
        }

        #region OTHER RESPONSE FORMATS
        //public void respCreatePackingSlipByWorkOrder(XmlDocument DOC)
        //{
        //}

        //public void respCreateUpdatePackingSlip(XmlDocument DOC)
        //{
        //}
        #endregion



        public XmlDocument sendXmlRequest(XmlDocument pDoc)
        {
            XmlDocument docResp = null;
            XmlDocument docReq = pDoc;

            HttpWebRequest objHttpWebRequest;
            HttpWebResponse objHttpWebResponse = null;
            
            Stream objRequestStream = null;
            Stream objResponseStream = null;
            
            XmlTextReader objXMLReader;

            objHttpWebRequest = (HttpWebRequest)WebRequest.Create(Globals.get_printableURI_packingSlip);

            try
            {
                byte[] bytes;
                bytes = System.Text.Encoding.ASCII.GetBytes(docReq.InnerXml);
                objHttpWebRequest.Method = "POST";
                objHttpWebRequest.ContentLength = bytes.Length;
                objHttpWebRequest.ContentType = "text/xml; encoding='utf-8'";

                objRequestStream = objHttpWebRequest.GetRequestStream();

                objRequestStream.Write(bytes, 0, bytes.Length);

                objRequestStream.Close();

                objHttpWebResponse = (HttpWebResponse)objHttpWebRequest.GetResponse();

                if (objHttpWebResponse.StatusCode == HttpStatusCode.OK)
                {
                    objResponseStream = objHttpWebResponse.GetResponseStream();

                    objXMLReader = new XmlTextReader(objResponseStream);

                    XmlDocument xmldoc = new XmlDocument();
                    xmldoc.Load(objXMLReader);

                    docResp = xmldoc;

                    objXMLReader.Close();
                }
                objHttpWebResponse.Close();
            }
            catch (WebException we)
            {
                Console.WriteLine(we.ToString());
            }
            catch (Exception e)
            {
                Console.WriteLine( e.ToString());
            }
            finally
            {
                objRequestStream.Close();
                objResponseStream.Close();
                objHttpWebResponse.Close();

                objXMLReader = null;
                objRequestStream = null;
                objResponseStream = null;
                objHttpWebRequest = null;
                objHttpWebResponse = null;
            }

            return docResp;
        }

        public void updatePrintableProcessed(string pID)
        {
            string id = pID;
            string queryString = @"update CT_UPS_EODitems
                                   set PTI_processed = 1
                                   where id = " + id;

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(queryString, conn);
                    try
                    {
                        int rowsUpdated = 0;
                        command.Connection.Open();
                        rowsUpdated = command.ExecuteNonQuery();
                        command.Dispose();
                        command = null;
                        Console.WriteLine(rowsUpdated.ToString() + " records processed and updated");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }



        public void responseError(string pErrNum, string pErr, string pErrDesc, string pPrintableID, string pPackingSlip)
        {
            string errNum =  "PS-" + pErrNum;
            string errDesc = pErrDesc;
            string err = pErr;
            string errDate = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss");
            string packingSlipNum = pPackingSlip;
            string printableID = pPrintableID;

            string queryString = @"INSERT INTO CT_PTI_errorLog (errNum, errDesc, errMsg, errDate, printableID, packingSlipNum )
                                   VALUES ('" + errNum + "','" + errDesc + "','" + err + "','" + errDate + "','" + printableID + "','" + packingSlipNum + "')";

            try
            {
                using (SqlConnection conn = new SqlConnection(Globals.get_printableConnString))
                {
                    SqlCommand command = new SqlCommand(queryString, conn);
                    try
                    {
                        int rowsInserted = 0;
                        command.Connection.Open();
                        rowsInserted = command.ExecuteNonQuery();
                        command.Dispose();
                        command = null;
                        Console.WriteLine("");
                        Console.WriteLine(rowsInserted.ToString() + " row inserted into error log");
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }    
}

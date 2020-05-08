using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.ServiceModel.Description;
using System.Text;
using System.Xml;
using OfficeOpenXml;

namespace ExportData
{
    class Program
    {
        static IOrganizationService orgService;

        static void Main(string[] args)
        {
            Console.WriteLine("Start : " + DateTime.Now);
            string url = System.Configuration.ConfigurationManager.AppSettings["url"].ToString();
            string username = System.Configuration.ConfigurationManager.AppSettings["username"].ToString();
            string Password = System.Configuration.ConfigurationManager.AppSettings["Password"].ToString();
            Console.WriteLine("Connection  URL: " + url);
            Console.WriteLine("Connection UserName : " + username);
            Console.WriteLine("Connection Password : " + Password);

            Console.WriteLine("Connecting to CRM service : " + DateTime.Now);
            //System.Threading.Thread.Sleep(2000);
            ConnectToMSCRM(username, Password, url);
            //Test connection
            Guid userid = ((WhoAmIResponse)orgService.Execute(new WhoAmIRequest())).UserId;
            if (userid != Guid.Empty)
            {
                Console.WriteLine("Connection Established : " + DateTime.Now);
            }

        EName:
            Console.WriteLine("Please specify entity name");
            var entityname = Console.ReadLine();
            if (entityname.Length == 0) goto EName;

            Efolder:

            Console.WriteLine("Please select destination file type: \n 1: for .txt \n 2: for .xlsx \n 3: for .csv");
            var filepath = Console.ReadLine();
            if (filepath.Length == 0) goto Efolder;
            else if (filepath != "1" && filepath != "2" && filepath != "3")
            {
                Console.WriteLine("Incorrect file extension selected");
                goto Efolder;
            }
           
            Console.WriteLine("Fetching Data : " + DateTime.Now);

            // Set the number of records per page to retrieve.
            int fetchCount = 5000;
            // Initialize the page number.
            int pageNumber = 1;
            // Initialize the number of records.
            int recordCount = 0;
            // Specify the current paging cookie. For retrieving the first page, 
            // pagingCookie should be null.
            string pagingCookie = null;

            string fetchXML = @"<fetch version='1.0' mapping='logical' output-format='xml - platform'><entity name = '" + entityname + "'><all-attributes/></entity></fetch> ";
            DataTable dtData = new DataTable();

            while (true)
            {
                // Build fetchXml string with the placeholders.
                string xml = CreateXml(fetchXML, pagingCookie, pageNumber, fetchCount);

                // Excute the fetch query and get the xml result.
                RetrieveMultipleRequest fetchRequest1 = new RetrieveMultipleRequest
                {
                    Query = new FetchExpression(xml)
                };

                EntityCollection returnCollection = ((RetrieveMultipleResponse)orgService.Execute(fetchRequest1)).EntityCollection;


                if (dtData.Columns.Count == 0)
                {
                    if (returnCollection.Entities.Count > 0)
                    {
                        foreach (var item in returnCollection.Entities[0].Attributes)
                        {
                            dtData.Columns.Add(item.Key);
                        }
                    }
                }
                foreach (Entity record in returnCollection.Entities)
                {
                    DataRow dr = dtData.NewRow();
                    foreach (KeyValuePair<string, object> item in returnCollection.Entities[0].Attributes)
                    {
                        if (!dtData.Columns.Contains(item.Key)) dtData.Columns.Add(item.Key);
                        dr[item.Key] = SantizeData(record, item);
                    }
                    dtData.Rows.Add(dr);
                }

                // Check for morerecords, if it returns 1.
                if (returnCollection.MoreRecords)
                {
                    Console.WriteLine("pageNumber: " + pageNumber + " Record Count " + dtData.Rows.Count + " : " + DateTime.Now.ToString());
                    Console.WriteLine("\n********************************");

                    // Increment the page number to retrieve the next page.
                    pageNumber++;

                    // Set the paging cookie to the paging cookie returned from current results.                            
                    pagingCookie = returnCollection.PagingCookie;
                }
                else
                {
                    // If no more records in the result nodes, exit the loop.
                    break;
                }
            }
            Console.WriteLine("Total Record Fetch :" + dtData.Rows.Count);
            Console.WriteLine("End : Data Fetch" + DateTime.Now);

            Console.WriteLine("Start : File Creation " + DateTime.Now);

            switch (filepath)
            {
                case "1":
                    Write(dtData, entityname + ".txt");
                    break;

                case "2":
                    string xlsxpath = entityname + ".xlsx";
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (ExcelPackage pck = new ExcelPackage())
                    {
                        pck.File = new FileInfo(xlsxpath);
                        ExcelWorksheet ws = pck.Workbook.Worksheets.Add(entityname);
                        ws.Cells["A1"].LoadFromDataTable(dtData, true);
                        pck.Save();
                    }
                    break;

                case "3":
                    StringBuilder sb = new StringBuilder();

                    IEnumerable<string> columnNames = dtData.Columns.Cast<DataColumn>().
                                                      Select(column => column.ColumnName);
                    sb.AppendLine(string.Join(",", columnNames));

                    foreach (DataRow row in dtData.Rows)
                    {
                        IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                        sb.AppendLine(string.Join(",", fields));
                    }

                    File.WriteAllText(entityname + ".csv", sb.ToString());
                    break;
            }

            Console.WriteLine("End : File Creation " + DateTime.Now);

            Console.ReadLine();
        }

        static void Write(DataTable dt, string outputFilePath)
        {
            int[] maxLengths = new int[dt.Columns.Count];

            for (int i = 0; i < dt.Columns.Count; i++)
            {
                maxLengths[i] = dt.Columns[i].ColumnName.Length;

                foreach (DataRow row in dt.Rows)
                {
                    if (!row.IsNull(i))
                    {
                        int length = row[i].ToString().Length;

                        if (length > maxLengths[i])
                        {
                            maxLengths[i] = length;
                        }
                    }
                }
            }

            using (StreamWriter sw = new StreamWriter(outputFilePath, false))
            {
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    sw.Write(dt.Columns[i].ColumnName.PadRight(maxLengths[i] + 2));
                }

                sw.WriteLine();

                foreach (DataRow row in dt.Rows)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        if (!row.IsNull(i))
                        {
                            sw.Write(row[i].ToString().PadRight(maxLengths[i] + 2));
                        }
                        else
                        {
                            sw.Write(new string(' ', maxLengths[i] + 2));
                        }
                    }

                    sw.WriteLine();
                }

                sw.Close();
            }
        }
        public static string SantizeData(Entity c, KeyValuePair<string, object> field)
        {
            string data = "";
            if (c.Attributes.ContainsKey(field.Key))
            {
                if (field.ToString().Contains("OptionSetValue"))
                {
                    data = ((Microsoft.Xrm.Sdk.OptionSetValue)field.Value).Value.ToString();
                }
                else if (field.ToString().Contains("EntityReference"))
                {
                    data = ((Microsoft.Xrm.Sdk.EntityReference)field.Value).LogicalName + " : " + ((Microsoft.Xrm.Sdk.EntityReference)field.Value).Name + " : " + ((Microsoft.Xrm.Sdk.EntityReference)field.Value).Id;
                }
                else { data = c.Attributes[field.Key].ToString(); }
            }
            return data;
        }
        public static void ConnectToMSCRM(string UserName, string Password, string SoapOrgServiceUri)
        {
            try
            {
                ClientCredentials credentials = new ClientCredentials();
                credentials.UserName.UserName = UserName;
                credentials.UserName.Password = Password;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                Uri serviceUri = new Uri(SoapOrgServiceUri);
                OrganizationServiceProxy proxy = new OrganizationServiceProxy(serviceUri, null, credentials, null);
                proxy.EnableProxyTypes();
                orgService = (IOrganizationService)proxy;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while connecting to CRM " + ex.Message);
                Console.ReadKey();
            }
        }
        public static string CreateXml(string xml, string cookie, int page, int count)
        {
            StringReader stringReader = new StringReader(xml);
            XmlTextReader reader = new XmlTextReader(stringReader);

            // Load document
            XmlDocument doc = new XmlDocument();
            doc.Load(reader);

            return CreateXml(doc, cookie, page, count);
        }

        public static string CreateXml(XmlDocument doc, string cookie, int page, int count)
        {
            XmlAttributeCollection attrs = doc.DocumentElement.Attributes;

            if (cookie != null)
            {
                XmlAttribute pagingAttr = doc.CreateAttribute("paging-cookie");
                pagingAttr.Value = cookie;
                attrs.Append(pagingAttr);
            }

            XmlAttribute pageAttr = doc.CreateAttribute("page");
            pageAttr.Value = System.Convert.ToString(page);
            attrs.Append(pageAttr);

            XmlAttribute countAttr = doc.CreateAttribute("count");
            countAttr.Value = System.Convert.ToString(count);
            attrs.Append(countAttr);

            StringBuilder sb = new StringBuilder(1024);
            StringWriter stringWriter = new StringWriter(sb);

            XmlTextWriter writer = new XmlTextWriter(stringWriter);
            doc.WriteTo(writer);
            writer.Close();

            return sb.ToString();
        }
    }
}

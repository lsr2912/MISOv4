/*
Copyright (c) 2009, Serials Solutions
All rights reserved.

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

    * Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
    * Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
    * Neither the name of the <ORGANIZATION> nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace SushiLibrary
{
    public static class ReportLoader
    {
        private static string GetValue(XmlNode node)
        {
            if (node == null)
            {
                return null;
            }
            return node.Value;
        }

        private static string GetInnerText(XmlNode node)
        {
            if (node == null)
            {
                return null;
            }
            return node.InnerText;
        }

        //-----------------------------------------------
        // debugging -lsr - 6/15
        // function only used to try to clean title names
        //-----------------------------------------------
        static string StripExtended(string arg)
        {
            StringBuilder buffer = new StringBuilder(arg.Length); //Max length
            foreach (char ch in arg)
            {
                UInt16 num = Convert.ToUInt16(ch);//In .NET, chars are UTF-16
                //The basic characters have the same code points as ASCII, and the extended characters are bigger
                if ((num >= 32u) && (num <= 126u)) buffer.Append(ch);
            }
            return buffer.ToString();
        }
        //-- End -lsr - 6/15  --
        //-----------------------------------------------


        public static SushiReport LoadCounterReport(XmlDocument reportXml)
        {

            //----------------------
            // debugging -lsr - 6/15
            // var _dc = 0;
            //----------------------
            SushiReport sushiReport = new SushiReport();

            XmlNamespaceManager xmlnsManager = new XmlNamespaceManager(reportXml.NameTable);

            xmlnsManager.AddNamespace("c", "http://www.niso.org/schemas/counter");
            xmlnsManager.AddNamespace("s", "http://www.niso.org/schemas/sushi");

            sushiReport.ReportType =
                (CounterReportType)
                Enum.Parse(typeof (CounterReportType),
                           GetValue(
                            reportXml.SelectSingleNode("//s:ReportDefinition", xmlnsManager).Attributes["Name"]),
                           true);

            sushiReport.Release = GetValue(reportXml.SelectSingleNode("//s:ReportDefinition", xmlnsManager).Attributes["Release"]);

            XmlNodeList reports = reportXml.SelectNodes("//c:Report", xmlnsManager);

            if (reports != null)
            {
                sushiReport.CounterReports = new List<CounterReport>();

                foreach (XmlNode report in reports)
                {
                    DateTime created;
                    DateTime.TryParse(GetValue(report.Attributes["Created"]), out created);

                    var counterReport = new CounterReport();

                    counterReport.ID = GetValue(report.Attributes["ID"]);
                    counterReport.Name = GetValue(report.Attributes["Name"]);
                    counterReport.Title = GetValue(report.Attributes["Title"]);
                    counterReport.Version = GetValue(report.Attributes["Version"]);
                    counterReport.Created = created;
                    counterReport.Vendor_Id = GetInnerText(report.SelectSingleNode("c:Vendor/c:ID", xmlnsManager));
                    counterReport.Vendor_Name = GetInnerText(report.SelectSingleNode("c:Vendor/c:Name", xmlnsManager));
                    counterReport.Vendor_ContactEmail =
                        GetInnerText(report.SelectSingleNode("c:Vendor/c:Contact/c:E-mail", xmlnsManager));
                    counterReport.Vendor_WebSiteUrl =
                        GetInnerText(report.SelectSingleNode("c:Vendor/c:WebSiteUrl", xmlnsManager));
                    counterReport.Vendor_LogoUrl = GetInnerText(report.SelectSingleNode("c:Vendor/c:LogoUrl", xmlnsManager));

                    XmlNodeList customers = report.SelectNodes("c:Customer", xmlnsManager);
                    if (customers != null)
                    {
                        counterReport.CustomerReports = new List<CounterCustomerReport>();

                        foreach (XmlNode customer in customers)
                        {
                            var customerReport = new CounterCustomerReport();
                            
                            customerReport.Customer_ID = GetInnerText(customer.SelectSingleNode("c:ID", xmlnsManager));
                            customerReport.Customer_Name = GetInnerText(customer.SelectSingleNode("c:Name", xmlnsManager));
                            customerReport.Customer_Consortium_Code =
                                GetInnerText(customer.SelectSingleNode("c:Consortium/c:Code", xmlnsManager));
                            customerReport.Customer_Consortium_WellKnownName =
                                GetInnerText(customer.SelectSingleNode("c:Consortium/c:WellKnownName", xmlnsManager));

                            XmlNodeList reportItems = customer.SelectNodes("c:ReportItems", xmlnsManager);

                            if (reportItems != null)
                            {
                                customerReport.ReportItems = new List<CounterReportItem>();
                                foreach (XmlNode reportItem in reportItems)
                                {
                                    CounterReportItem counterReportItem;
                                    switch (sushiReport.ReportType)
                                    {
                                        case CounterReportType.JR1:
                                        case CounterReportType.JR1GOA:
                                        case CounterReportType.JR2:
                                        case CounterReportType.JR3:
                                        case CounterReportType.JR4:
                                        case CounterReportType.JR5:
                                            JournalReportItem journalReportItem = new JournalReportItem();
                                            XmlNodeList jr_identifiers = reportItem.SelectNodes("c:ItemIdentifier", xmlnsManager);

                                            if (jr_identifiers != null)
                                            {
                                                foreach (XmlNode identifier in jr_identifiers)
                                                {
                                                    string value = identifier.SelectSingleNode("c:Value", xmlnsManager).InnerText;
                                                    switch (identifier.SelectSingleNode("c:Type", xmlnsManager).InnerText.ToLower())
                                                    {
                                                        // see http://www.niso.org/workrooms/sushi/values/#item
                                                        case "issn":
                                                            journalReportItem.Print_ISSN = value;
                                                            break;
                                                        case "print_issn":
                                                            journalReportItem.Print_ISSN = value;
                                                            break;
                                                        case "online_issn":
                                                            journalReportItem.Online_ISSN = value;
                                                            break;
                                                        case "doi":
                                                            journalReportItem.DOI = value;
                                                            break;
                                                        case "proprietary":
                                                            journalReportItem.Proprietary = value;
                                                            break;
                                                    }
                                                }
                                            }
                                            counterReportItem = journalReportItem;
                                            break;
                                        case CounterReportType.BR1:
                                        case CounterReportType.BR2:
                                        case CounterReportType.BR3:
                                        case CounterReportType.BR4:
                                        case CounterReportType.BR5:
                                            BookReportItem bookReportItem = new BookReportItem();
                                            XmlNodeList br_identifiers = reportItem.SelectNodes("c:ItemIdentifier", xmlnsManager);

                                            if (br_identifiers != null)
                                            {
                                                foreach (XmlNode identifier in br_identifiers)
                                                {
                                                    string value = identifier.SelectSingleNode("c:Value", xmlnsManager).InnerText;
                                                    switch (identifier.SelectSingleNode("c:Type", xmlnsManager).InnerText.ToLower())
                                                    {
                                                        // see http://www.niso.org/workrooms/sushi/values/#item
                                                        case "isbn":
                                                            bookReportItem.Print_ISBN = value;
                                                            break;
                                                        case "print_isbn":
                                                            bookReportItem.Print_ISBN = value;
                                                            break;
                                                        case "online_isbn":
                                                            bookReportItem.Online_ISSN = value;
                                                            break;
                                                        case "online_issn":
                                                            bookReportItem.Online_ISSN = value;
                                                            break;
                                                        case "doi":
                                                            bookReportItem.DOI = value;
                                                            break;
                                                        case "proprietary":
                                                            bookReportItem.Proprietary = value;
                                                            break;
                                                    }
                                                }
                                            }
                                            counterReportItem = bookReportItem;
                                            break;
                                        case CounterReportType.CR1:
                                            ConsortiumReportItem consortiumReportItem = new ConsortiumReportItem();
                                            XmlNodeList ci_identifiers = reportItem.SelectNodes("c:ItemIdentifier", xmlnsManager);

                                            if (ci_identifiers != null)
                                            {
                                                foreach (XmlNode identifier in ci_identifiers)
                                                {
                                                    string value = identifier.SelectSingleNode("c:Value", xmlnsManager).InnerText;
                                                    switch (identifier.SelectSingleNode("c:Type", xmlnsManager).InnerText.ToLower())
                                                    {
                                                        // see http://www.niso.org/workrooms/sushi/values/#item
                                                        case "issn":
                                                            consortiumReportItem.Print_ISSN = value;
                                                            break;
                                                        case "print_issn":
                                                            consortiumReportItem.Print_ISSN = value;
                                                            break;
                                                        case "online_issn":
                                                            consortiumReportItem.Online_ISSN = value;
                                                            break;
                                                        case "doi":
                                                            consortiumReportItem.DOI = value;
                                                            break;
                                                        case "proprietary":
                                                            consortiumReportItem.Proprietary = value;
                                                            break;
                                                    }
                                                }
                                            }
                                            counterReportItem = consortiumReportItem;
                                            break;
                                        default:
                                            counterReportItem = new CounterReportItem();
                                            break;
                                    }

                                    counterReportItem.ItemName = reportItem.SelectSingleNode("c:ItemName", xmlnsManager).InnerText;
                                    //------------------------------------------------------------------------------------------------
                                    // debugging -lsr - 6/15
                                    //------------------------------------------------------------------------------------------------
                                    // counterReportItem.ItemName = reportItem.SelectSingleNode("c:ItemName", xmlnsManager).InnerText;
                                    // string _name = reportItem.SelectSingleNode("c:ItemName", xmlnsManager).InnerText;
                                    // counterReportItem.ItemName = StripExtended(_name);
                                    // _dc += 1;
                                    // Console.WriteLine("(Rec:" + Convert.ToString(_dc) + ") Name: " + counterReportItem.ItemName);
                                    //
                                    //--  End -lsr - 6/15  ---------------------------------------------------------------------------
                                    counterReportItem.ItemPublisher = reportItem.SelectSingleNode("c:ItemPublisher", xmlnsManager).InnerText;
                                    counterReportItem.ItemPlatform = reportItem.SelectSingleNode("c:ItemPlatform", xmlnsManager).InnerText;

                                    XmlNodeList PerfItems = reportItem.SelectNodes("c:ItemPerformance", xmlnsManager);

                                    if (PerfItems != null)
                                    {
                                        counterReportItem.PerformanceItems = new List<CounterPerformanceItem>();

                                        foreach (XmlNode perfItem in PerfItems)
                                        {
                                            var performanceItem = new CounterPerformanceItem();
                                            if (perfItem.Attributes["PubYr"] != null)
                                            {
                                                performanceItem.YOP = perfItem.Attributes["PubYr"].Value;
                                            }
                                            else if (perfItem.Attributes["PubYrTo"] != null) 
                                            {
                                                if (perfItem.Attributes["PubYrFrom"] != null)
                                                {
                                                    performanceItem.YOP = perfItem.Attributes["PubYrFrom"].Value + "-" + perfItem.Attributes["PubYrTo"].Value;
                                                }
                                                else
                                                {
                                                    performanceItem.YOP = "<= " + perfItem.Attributes["PubYrTo"].Value;
                                                }
                                            }
                                            else
                                            {
                                                performanceItem.YOP = "";
                                            }

                                            DateTime start, end;
                                            DateTime.TryParse(perfItem.SelectSingleNode("c:Period/c:Begin", xmlnsManager).InnerText, out start);
                                            DateTime.TryParse(perfItem.SelectSingleNode("c:Period/c:End", xmlnsManager).InnerText, out end);
                                            CounterMetricCategory category = CounterMetricCategory.Invalid;
                                            try
                                            {
                                                category = (CounterMetricCategory)Enum.Parse(typeof(CounterMetricCategory), perfItem.SelectSingleNode("c:Category", xmlnsManager).InnerText, true);
                                            }
                                            catch (ArgumentException ex)
                                            {
                                                Console.WriteLine(ex.Message);
                                                Console.WriteLine(string.Format("WARNING - Found Invalid Metric Category Type: {0}", perfItem.SelectSingleNode("c:Category", xmlnsManager).InnerText));
                                            }

                                            CounterMetric counterMetric = performanceItem.GetMetric(start, end, category);
                                            XmlNodeList instances = perfItem.SelectNodes("c:Instance", xmlnsManager);

                                            if (instances != null)
                                            {
                                                if (counterMetric.Instances == null)
                                                {
                                                    counterMetric.Instances = new List<CounterMetricInstance>();
                                                }
                                                foreach (XmlNode instance in instances)
                                                {
                                                    CounterMetricInstance metricInstance = new CounterMetricInstance();
                                                    metricInstance.Type =
                                                        (CounterMetricType)
                                                        Enum.Parse(typeof(CounterMetricType),
                                                                   instance.SelectSingleNode("c:MetricType",
                                                                                           xmlnsManager).InnerText, true);

                                                    // return exception if can't parse count, since it's important to process properly
                                                    metricInstance.Count = Int32.Parse(instance.SelectSingleNode("c:Count", xmlnsManager).InnerText);
                                                    counterMetric.Instances.Add(metricInstance);
                                                }
                                            }
                                            counterReportItem.PerformanceItems.Add(performanceItem);
                                        }
                                    }
                                    customerReport.ReportItems.Add(counterReportItem);
                                }
                            }
                            counterReport.CustomerReports.Add(customerReport);
                        }
                   }
                   sushiReport.CounterReports.Add(counterReport);
                }
            }
            return sushiReport;
        }

    }
}

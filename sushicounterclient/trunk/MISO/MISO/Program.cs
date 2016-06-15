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
using System.Globalization;
using System.IO;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Xml;
using System.Xml.Schema;
using SushiLibrary;
using SushiLibrary.Validation;

namespace MISO
{
    class Program
    {
        // constants
        private const string CounterSchemaURL = "http://www.niso.org/schemas/sushi/counter_sushi4_1.xsd";
        private static readonly string[] DateTimeFormats = { "yyyyMM" };
        private const string CommandParameterMessage = @"Parameters are:
MISO.EXE [-v] [filename] [-d] [start] [end] [-l] [Library codes separated by commas]
-v: Validate Mode: Validates a specified file or reports from the sushi server with the counter-sushi schema

-s: Strict Validation Mode: Validates a specified file or reports from the sushi server with the schema and custom validation rules

-d: Specify Sushi request start and end dates
    [start]: start date in yyyymm format
    [end]: end date in yyyymm format
    By default, it will automatically request the previous month from the current date.

-l: Specify Library Codes

-x: Save Request Response as XML without converting to csv";

        private static readonly char[] DELIM = { ',' };

        // global variables to MISO
        private static DateTime StartDate;
        private static DateTime EndDate;
        private static string RequestTemplate;
        private static bool XmlMode = false;
        private static bool ValidateMode = false;
        private static bool StrictValidateMode = false;
        private static readonly XmlReaderSettings XmlReaderSettings = new XmlReaderSettings();

        // lookup table to find month data
        private static Dictionary<string, string> MonthData;

        private static TextWriter _errorFile = null;
        private static TextWriter ErrorFile
        {
            get
            {
                // create the file
                if (_errorFile == null)
                {
                    _errorFile = new StreamWriter(string.Format("Error_{0}.txt", ErrorDate));
                }
                return _errorFile;
            }
        }

        //validation mode stuff
        private static bool IsValid = true;

        private static void CounterV41ValidationEventHandler(object sender, ValidationEventArgs args)
        {
            IsValid = false;
            Console.WriteLine("Validation event\n" + args.Message);

        }

        private static string ErrorDate
        {
            get { return DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss"); }
        }

        static void Main(string[] args)
        {
            #region Process Arguements
            string validateFile = string.Empty;
            bool specifiedLibCodes = false;
            string start = string.Empty;
            string end = string.Empty;
            string libCodeStr = string.Empty;

            try
            {

                for (int i = 0; i < args.Length; i++)
                {
                    switch (args[i])
                    {

                        case "-v":
                            ValidateMode = true;
                            if (i + 1 < args.Length)
                            {
                                validateFile = args[i + 1];
                            }
                            break;
                        case "-d":
                            if (i + 1 < args.Length)
                            {
                                start = args[i + 1];
                            }
                            if (i + 2 < args.Length)
                            {
                                end = args[i + 2];
                            }
                            break;
                        case "-l":
                            specifiedLibCodes = true;
                            if (i + 1 >= args.Length)
                            {
                                throw new ArgumentException("File not specified for validation mode");
                            }
                            libCodeStr = args[i + 1];
                            break;
                        case "-x":
                            XmlMode = true;
                            break;
                        case "-h":
                            Console.WriteLine(CommandParameterMessage);
                            System.Environment.Exit(-1);
                            break;
                        case "-s":
                            StrictValidateMode = true;
                            if (i + 1 < args.Length)
                            {
                                validateFile = args[i + 1];
                            }
                            break;
                    }
                }
            }
            catch (ArgumentException ex)
            {
                Console.WriteLine(ex.Message);
            }
            #endregion

            FileStream sushiConfig = new FileStream("sushiconfig.csv", FileMode.Open, FileAccess.Read);
            StreamReader sr = new StreamReader(sushiConfig);

            //initiate validation
            XmlSchema CounterSushiSchema = XmlSchema.Read(new XmlTextReader(CounterSchemaURL), new ValidationEventHandler(CounterV41ValidationEventHandler));
            XmlReaderSettings.ValidationType = ValidationType.Schema;
            XmlReaderSettings.Schemas.Add(CounterSushiSchema);
            XmlReaderSettings.ValidationEventHandler += new ValidationEventHandler(CounterV41ValidationEventHandler);

            try
            {

            #region Initialize
                // validate file mode
                if ((ValidateMode || StrictValidateMode) && !string.IsNullOrEmpty(validateFile))
                {
                    using (XmlReader xmlReader = XmlReader.Create(new XmlTextReader(validateFile), XmlReaderSettings))
                    {

                        // Read XML to the end
                        try
                        {
                            while (xmlReader.Read())
                            {
                                // just read through file to trigger any validation errors
                            }

                            Console.WriteLine("\nFinished validating XML file....");
                        }
                        catch (FileNotFoundException e)
                        {
                            Console.WriteLine(e.Message);
                            System.Environment.Exit(-1);
                        }
                    }

                    if (StrictValidateMode)
                    {
                        var file = new FileStream(validateFile, FileMode.Open, FileAccess.Read);
                        var xmlFile = new XmlDocument();
                        xmlFile.Load(file);

                        SushiReport report = ReportLoader.LoadCounterReport(xmlFile);

                        var validator = new CounterValidator();
                        if (!validator.Validate(report))
                        {
                            IsValid = false;

                            foreach (var error in validator.ErrorMessage)
                            {
                                Console.WriteLine(error);
                            }
                        }

                    }

                    if (IsValid)
                    {
                        Console.WriteLine("Document is valid");
                    }
                    else
                    {
                        Console.WriteLine("Document is invalid");
                    }
                }

                else
                {

                    string[] header = sr.ReadLine().Split(DELIM);

                    DateTime startMonth = DateTime.Now.AddMonths(-1);
                    DateTime endMonth = DateTime.Now.AddMonths(-1);

                    
                    if (!string.IsNullOrEmpty(start))
                    {
                        startMonth = DateTime.ParseExact(start, DateTimeFormats, null,
                                                                  DateTimeStyles.None);
                    }
                    if (!string.IsNullOrEmpty(end))
                    {
                        endMonth = DateTime.ParseExact(end, DateTimeFormats, null,
                                                                DateTimeStyles.None);
                    }

                    if (endMonth < startMonth)
                    {
                        throw new ArgumentException("End date is before start date.");
                    }

                    StartDate = new DateTime(startMonth.Year, startMonth.Month, 1);
                    EndDate = new DateTime(endMonth.Year, endMonth.Month,
                                           DateTime.DaysInMonth(endMonth.Year, endMonth.Month));

                    FileStream requestTemplate = new FileStream("SushiSoapEnv.xml", FileMode.Open, FileAccess.Read);
                    StreamReader reader = new StreamReader(requestTemplate);
                    RequestTemplate = reader.ReadToEnd();
                    reader.Close();

                    #endregion

                    Dictionary<string, string> libCodeMap = null;
                    if (specifiedLibCodes)
                    {
                        libCodeMap = new Dictionary<string, string>();

                        string[] libCodes = libCodeStr.Split(DELIM);
                        foreach (string libCode in libCodes)
                        {
                            libCodeMap.Add(libCode.ToUpper(), string.Empty);
                        }
                    }

                    string buffer;
                    for (int lineNum = 1; (buffer = sr.ReadLine()) != null; lineNum++)
                    {
                        string[] fields = buffer.Split(DELIM);

                        if (libCodeMap == null || libCodeMap.ContainsKey(fields[0].ToUpper()))
                        {
                            if (fields.Length < 25)
                            {
                                ErrorFile.WriteLine(string.Format("{0}: Line {1} has insufficient data", ErrorDate,
                                                                  lineNum));
                            }
                            else
                            {
                                //loop through report types in header
                                for (int i = 9; i < 26; i++)
                                {
                                    try
                                    {
                                        if (fields[i].ToLower().StartsWith("y"))
                                        {
                                            ProcessSushiRequest(header[i], fields);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        ErrorFile.WriteLine(
                                            string.Format(
                                                "{0}: Exception occurred processing line {1} for report type {2}",
                                                ErrorDate,
                                                lineNum, header[i]));
                                        ErrorFile.WriteLine(ex.Message);
                                        ErrorFile.Write(ex.StackTrace);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (FormatException ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                sr.Close();
                sushiConfig.Close();
                if (_errorFile != null)
                {
                    _errorFile.Close();
                }
            }
        }

        /// <summary>
        /// Make the request to the sushi server with the given request
        /// </summary>
        /// <param name="reqDoc"></param>
        /// <param name="url"></param>
        /// <returns></returns>
        private static string CallSushiServer(XmlDocument reqDoc, string url)
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Headers.Add("SOAPAction", "\"SushiService:GetReportIn\"");

            ServicePointManager.ServerCertificateValidationCallback += customXertificateValidation;

            req.ContentType = "text/xml;charset=\"utf-8\"";
            req.Accept = "text/xml";
            req.Method = "POST";
            req.Timeout = Int32.MaxValue;
            Stream stm = req.GetRequestStream();
            reqDoc.Save(stm);
            stm.Close();

            WebResponse resp = req.GetResponse();

            StreamReader SReader= new StreamReader(resp.GetResponseStream());
            return SReader.ReadToEnd();
        }

        private static bool customXertificateValidation(object sender, X509Certificate cert, X509Chain chain, System.Net.Security.SslPolicyErrors error)
        {
            //TODO just accept the cert for now since they usually tend to be unregistered
            return true;

        }
        
        /// <summary>
        /// Send sushi request and convert to csv
        /// </summary>
        /// <param name="reportType"></param>
        /// <param name="fields"></param>
        private static void ProcessSushiRequest(string reportType, string[] fields)
        {
            string fileName = string.Format("{0}_{1}_{2}_{3}_{4}", fields[1], fields[0], StartDate.ToString("yyyyMM"), EndDate.ToString("yyyyMM"), reportType);

            fileName = XmlMode ? fileName + ".xml" : fileName + ".csv";

            string startDateStr = StartDate.Date.ToString("yyyy-MM-dd");
            string endDateStr = EndDate.Date.ToString("yyyy-MM-dd");
            string createdDateStr = DateTime.Now.ToString("s");

            XmlDocument reqDoc = new XmlDocument();
            if (fields.Length == 27 && !string.IsNullOrEmpty(fields[25]) && !string.IsNullOrEmpty(fields[26]))
            {
                // Load WSSE fields for Proquest
                FileStream wsSecurityFile = new FileStream("WSSecurityPlainText.xml", FileMode.Open, FileAccess.Read);
                StreamReader reader = new StreamReader(wsSecurityFile);
                string wsSecuritySnippet = string.Format(reader.ReadToEnd(), fields[20], fields[21]);
                reader.Close();
                reqDoc.LoadXml(
                    string.Format(RequestTemplate, fields[4], fields[5], fields[6], fields[7],
                                  fields[8], reportType, fields[2], startDateStr, endDateStr, createdDateStr, wsSecuritySnippet));
            }
            else
            {
                reqDoc.LoadXml(
                    string.Format(RequestTemplate, fields[4], fields[5], fields[6], fields[7],
                                  fields[8], reportType, fields[2], startDateStr, endDateStr, createdDateStr, string.Empty));
            }

            string resonseString = CallSushiServer(reqDoc, fields[3]);
            if (XmlMode)
            {
                StreamWriter sw = new StreamWriter(fileName);
                sw.Write(resonseString);
                sw.Flush();
                sw.Close();
                return; // parsing and conversion to csv is unnecessary
            }

            XmlDocument sushiDoc = new XmlDocument();
            sushiDoc.LoadXml(resonseString);

            XmlNamespaceManager xmlnsManager = new XmlNamespaceManager(sushiDoc.NameTable);
            // Proquest Error
            xmlnsManager.AddNamespace("s", "http://www.niso.org/schemas/sushi");
            XmlNode exception = sushiDoc.SelectSingleNode("//s:Exception", xmlnsManager);

            if (exception != null && exception.HasChildNodes)
            {
                Console.WriteLine(string.Format("Exception detected for report of type {0} for Provider: {1}\nPlease see error log for more details.", reportType, fields[1]));
                throw new XmlException(
                    string.Format("Report returned Exception: Number: {0}, Severity: {1}, Message: {2}",
                    exception.SelectSingleNode("s:Number", xmlnsManager).InnerText, exception.SelectSingleNode("s:Severity", xmlnsManager).InnerText, exception.SelectSingleNode("s:Message", xmlnsManager).InnerText));
            }

            SushiReport sushiReport = ReportLoader.LoadCounterReport(sushiDoc);

            if (ValidateMode || StrictValidateMode)
            {
                IsValid = true;

                MemoryStream ms = new MemoryStream();
                sushiDoc.Save(ms);
                ms.Position = 0;
                using (XmlReader xmlReader = XmlReader.Create(new XmlTextReader(ms), XmlReaderSettings))
                {

                    // Read XML to the end
                    while (xmlReader.Read())
                    {
                        // just read through file to trigger any validation errors
                    }
                }

                if (StrictValidateMode)
                {
                    var validator = new CounterValidator();
                    if (!validator.Validate(sushiReport))
                    {
                        IsValid = false;

                        foreach (var error in validator.ErrorMessage)
                        {
                            Console.WriteLine(error);
                        }
                    }
                }

                Console.WriteLine(string.Format("Finished validation Counter report of type {0} for Provider: {1}", reportType, fields[1]));
                if (IsValid)
                {
                    Console.WriteLine("Document is valid");
                }
                else
                {
                    Console.WriteLine("Document is invalid");
                }
            }
//            XmlWriter xw = new 

            TextWriter tw = new StreamWriter(fileName);
            StringBuilder header;

            Console.WriteLine("Parsing report of type: " + reportType);

            switch(reportType)
            {
                case "JR1":
                    tw.WriteLine("Journal Report 1 (R4),Number of Successful Full-Text Article Requests by Month and Journal");
                    tw.WriteLine(fields[0]);
                    tw.WriteLine("Period covered by report");
                    tw.WriteLine(startDateStr + " to " + endDateStr);
                    tw.WriteLine("Date run:");
                    tw.WriteLine(DateTime.Now.ToString("yyyy-M-d"));

                    // construct header
                    header = new StringBuilder(",Publisher,Platform,Journal DOI,Proprietary Identifier,Print ISSN,Online ISSN");
                    header.Append(",Reporting Period Total,Reporting Period HTML,Reporting Period PDF");
                    for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                    {
                        header.Append(",");
                        header.Append(currMonth.ToString("MMM-yyyy"));
                    }
                    tw.WriteLine(header);
                    ParseJR1v4(sushiReport, tw);
                    tw.Close();
                    break;
                case "JR1GOA":
                    tw.WriteLine("Journal Report 1 GOA (R4),Number of Successful Gold Open Access Full-text Article Requests by Month and Journal");
                    tw.WriteLine(fields[0]);
                    tw.WriteLine("Period covered by report");
                    tw.WriteLine(startDateStr + " to " + endDateStr);
                    tw.WriteLine("Date run:");
                    tw.WriteLine(DateTime.Now.ToString("yyyy-M-d"));

                    // construct header
                    header = new StringBuilder(",Publisher,Platform,Journal DOI,Proprietary Identifier,Print ISSN,Online ISSN");
                    header.Append(",Reporting Period Total,Reporting Period HTML,Reporting Period PDF");
                    for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                    {
                        header.Append(",");
                        header.Append(currMonth.ToString("MMM-yyyy"));
                    }
                    tw.WriteLine(header);
                    ParseJR1v4(sushiReport, tw);
                    tw.Close();
                    break;
                case "JR2":
                    tw.WriteLine("Journal Report 2 (R4),Access Denied to Full-Text Articles by Month Journal and Category");
                    tw.WriteLine(fields[0]);
                    tw.WriteLine("Period covered by report");
                    tw.WriteLine(startDateStr + " to " + endDateStr);
                    tw.WriteLine("Date run:");
                    tw.WriteLine(DateTime.Now.ToString("yyyy-M-d"));

                    // construct header
                    header = new StringBuilder(",Publisher,Platform,Journal DOI,Proprietary Identifier,Print ISSN,Online ISSN,Access Denied Category,Reporting Period Total");
                    for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                    {
                        header.Append(",");
                        header.Append(currMonth.ToString("MMM-yyyy"));
                    }
                    tw.WriteLine(header);
                    ParseJR2v4(sushiReport, tw);
                    tw.Close();
                    break;
                case "JR5":
                    tw.WriteLine("Journal Report 5 (R4),Number of Successful Full-Text Article Requests by Year-of-Publication (YOP) and Journal");
                    tw.WriteLine(fields[0]);
                    tw.WriteLine("Period covered by report");
                    tw.WriteLine(startDateStr + " to " + endDateStr);
                    tw.WriteLine("Date run:");
                    tw.WriteLine(DateTime.Now.ToString("yyyy-M-d"));

                    // JR5 column headings are produced inside the function
                    //
                    ParseJR5v4(sushiReport, tw);
                    tw.Close();
                    break;
                case "DB1":
                    tw.WriteLine("Database Report 1 (R4), Total Searches Result Clicks and Record Views by Month and Database");
                    tw.WriteLine(fields[0]);
                    tw.WriteLine("Period covered by report");
                    tw.WriteLine(startDateStr + " to " + endDateStr);
                    tw.WriteLine("Date run:");
                    tw.WriteLine(DateTime.Now.ToString("yyyy-M-d"));

                    // construct header
                    header = new StringBuilder(",Publisher,Platform,User Activity,Reporting Period Total");
                    for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                    {
                        header.Append(",");
                        header.Append(currMonth.ToString("MMM-yyyy"));
                    }
                    tw.WriteLine(header);
                    ParseDB1v4(sushiReport, tw);
                    tw.Close();
                    break;
                case "DB2":
                    tw.WriteLine("Database Report 2 (R4), Access Denied by Month Database and Category");
                    tw.WriteLine(fields[0]);
                    tw.WriteLine("Period covered by report");
                    tw.WriteLine(startDateStr + " to " + endDateStr);
                    tw.WriteLine("Date run:");
                    tw.WriteLine(DateTime.Now.ToString("yyyy-M-d"));

                    // construct header
                    header = new StringBuilder(",Publisher,Platform,Access Denied Category");
                    for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                    {
                        header.Append(",");
                        header.Append(currMonth.ToString("MMM-yyyy"));
                    }

                    header.Append(",Reporting Period Total");
                    tw.WriteLine(header);
                    ParseDB2v4(sushiReport, tw);
                    tw.Close();
                    break;
                case "PR1":
                    tw.WriteLine("Platform Report 1 (R4): Total Searches Result Clicks and Record Views by Month and Platform");
                    tw.WriteLine(fields[0]);
                    tw.WriteLine("Period covered by report");
                    tw.WriteLine(startDateStr + " to " + endDateStr);
                    tw.WriteLine("Date run:");
                    tw.WriteLine(DateTime.Now.ToString("yyyy-M-d"));

                    // construct header
                    header = new StringBuilder("Platform,Publisher,User Activity,Reporting Period Total");
                    for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                    {
                        header.Append(",");
                        header.Append(currMonth.ToString("MMM-yyyy"));
                    }
                    tw.WriteLine(header);
                    ParsePR1v4(sushiReport, tw);
                    tw.Close();
                    break;
                case "BR1":
                    tw.WriteLine("Book Report 1 (R4): Number of Successful Title Requests by Month and Title");
                    tw.WriteLine("Customer");
                    tw.WriteLine(fields[0]);
                    tw.WriteLine("Period covered by report");
                    tw.WriteLine(startDateStr + " to " + endDateStr);
                    tw.WriteLine("Date run:");
                    tw.WriteLine(DateTime.Now.ToString("yyyy-M-d"));

                    // construct header
                    header = new StringBuilder(",Publisher,Platform,DOI,Proprietary Identifier,Print ISBN,Online ISSN,Reporting Period Total");
                    for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                    {
                        header.Append(",");
                        header.Append(currMonth.ToString("MMM-yyyy"));
                    }
                    tw.WriteLine(header);
                    ParseBR12v4(sushiReport, tw);
                    tw.Close();
                    break;
                case "BR2":
                    tw.WriteLine("Book Report 2 (R4),Number of Successful Section Requests by Month and Title");
                    tw.WriteLine("Customer,Section Type:");
                    tw.WriteLine(string.Format("{0},ft_total", fields[0]));
                    tw.WriteLine("Period covered by report");
                    tw.WriteLine(startDateStr + " to " + endDateStr);
                    tw.WriteLine("Date run:");
                    tw.WriteLine(DateTime.Now.ToString("yyyy-M-d"));

                    // construct header
                    header = new StringBuilder(",Publisher,Platform,DOI,Proprietary Identifier,Print ISBN,Online ISSN,Reporting Period Total");
                    for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                    {
                        header.Append(",");
                        header.Append(currMonth.ToString("MMM-yy"));
                    }
                    tw.WriteLine(header);
                    ParseBR12v4(sushiReport, tw);
                    tw.Close();
                    break;
                case "BR3":
                    tw.WriteLine("Book Report 3 (R4),Access Denied to Content Items by Month, Title and Category");
                    tw.WriteLine("Customer");
                    tw.WriteLine(fields[0]);
                    tw.WriteLine("Period covered by report");
                    tw.WriteLine(startDateStr + " to " + endDateStr);
                    tw.WriteLine("Date run:");
                    tw.WriteLine(DateTime.Now.ToString("yyyy-M-d"));

                    // construct header
                    header = new StringBuilder(",Publisher,Platform,Book DOI,Proprietary Identifier,ISBN,ISSN,Access Denied Category,Reporting Period Total");
                    for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                    {
                        header.Append(",");
                        header.Append(currMonth.ToString("MMM-yyyy"));
                    }
                    tw.WriteLine(header);
                    ParseBR3v4(sushiReport, tw);
                    tw.Close();
                    break;
                case "BR4":
                    tw.WriteLine("Book Report 4 (R4),Access Denied to Content Items by Month, Platform and Category");
                    tw.WriteLine("Customer");
                    tw.WriteLine(fields[0]);
                    tw.WriteLine("Period covered by report");
                    tw.WriteLine(startDateStr + " to " + endDateStr);
                    tw.WriteLine("Date run:");
                    tw.WriteLine(DateTime.Now.ToString("yyyy-M-d"));

                    // construct header
                    header = new StringBuilder(",Publisher,Platform,Access Denied Category,Reporting Period Total");
                    for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                    {
                        header.Append(",");
                        header.Append(currMonth.ToString("MMM-yyyy"));
                    }
                    tw.WriteLine(header);
                    ParseBR4v4(sushiReport, tw);
                    tw.Close();
                    break;
                case "BR5":
                    tw.WriteLine("Book Report 5 (R4),Total Searches by Month and Title");
                    tw.WriteLine("Customer");
                    tw.WriteLine(fields[0]);
                    tw.WriteLine("Period covered by report");
                    tw.WriteLine(startDateStr + " to " + endDateStr);
                    tw.WriteLine("Date run:");
                    tw.WriteLine(DateTime.Now.ToString("yyyy-M-d"));

                    // construct header
                    header = new StringBuilder(",Publisher,Platform,Book DOI,Proprietary Identifier,ISBN,ISSN,User Activity,Reporting Period Total");
                    for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                    {
                        header.Append(",");
                        header.Append(currMonth.ToString("MMM-yyyy"));
                    }
                    tw.WriteLine(header);
                    ParseBR4v4(sushiReport, tw);
                    tw.Close();
                    break;
                case "MR1":
                    tw.WriteLine("Multimedia Report 1 (R4), Number of Successful Multimedia Full Content Unit Requests by Month and Collection");
                    tw.WriteLine(fields[0]);
                    tw.WriteLine("Period covered by report");
                    tw.WriteLine(startDateStr + " to " + endDateStr);
                    tw.WriteLine("Date run:");
                    tw.WriteLine(DateTime.Now.ToString("yyyy-M-d"));

                    // construct header
                    header = new StringBuilder("Collection,Content Provider,Platform,Reporting Period Total");
                    for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                    {
                        header.Append(",");
                        header.Append(currMonth.ToString("MMM-yyyy"));
                    }
                    tw.WriteLine(header);
                    ParseMR1v4(sushiReport, tw);
                    tw.Close();
                    break;
                case "CR1":
                    tw.WriteLine("Consortium Report 1 (R4), Number of successful full-text journal article or book chapter requests by month and title");
                    tw.WriteLine(fields[0]);
                    tw.WriteLine("Period covered by report");
                    tw.WriteLine(startDateStr + " to " + endDateStr);
                    tw.WriteLine("Date run:");
                    tw.WriteLine(DateTime.Now.ToString("yyyy-M-d"));

                    // construct header
                    header = new StringBuilder("Customer,Publisher,Platform,Title,DOI,Proprietary Identifier,Print ISSN,Online ISSN");
                    header.Append(",Reporting Period Total,Reporting Period HTML,Reporting Period PDF");
                    for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                    {
                        header.Append(",");
                        header.Append(currMonth.ToString("MMM-yyyy"));
                    }
                    tw.WriteLine(header);
                    ParseCR1v4(sushiReport, tw);
                    tw.Close();
                    break;
                case "CR2":
                    tw.WriteLine("Consortium Report 2 (R4), Total searches by month and database");
                    tw.WriteLine(fields[0]);
                    tw.WriteLine("Period covered by report");
                    tw.WriteLine(startDateStr + " to " + endDateStr);
                    tw.WriteLine("Date run:");
                    tw.WriteLine(DateTime.Now.ToString("yyyy-M-d"));

                    // construct header
                    header = new StringBuilder("Customer,Publisher,Platform,User Activity,Reporting Period Total");
                    for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                    {
                        header.Append(",");
                        header.Append(currMonth.ToString("MMM-yyyy"));
                    }
                    tw.WriteLine(header);
                    ParseCR2v4(sushiReport, tw);
                    tw.Close();
                    break;
                case "CR3":
                    tw.WriteLine("Consortium Report 3 (R4), Number of Successful Multimedia Full Content Unit Requests by Month and Collection");
                    tw.WriteLine(fields[0]);
                    tw.WriteLine("Period covered by report");
                    tw.WriteLine(startDateStr + " to " + endDateStr);
                    tw.WriteLine("Date run:");
                    tw.WriteLine(DateTime.Now.ToString("yyyy-M-d"));

                    // construct header
                    header = new StringBuilder("Customer,Collection,Content Provider,Platform,Reporting Period Total");
                    for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                    {
                        header.Append(",");
                        header.Append(currMonth.ToString("MMM-yyyy"));
                    }
                    tw.WriteLine(header);
                    ParseCR3v4(sushiReport, tw);
                    tw.Close();
                    break;
                default:
                    ErrorFile.WriteLine(string.Format("{0}: Report Type {1} currently not supported.", ErrorDate, reportType));
                    break;
            }
        }

        private static void ParseJR1v4(SushiReport sushiReport, TextWriter tw)
        {
            // 
            // Report parser for JR1 and JR1GOA
            // --------------------------------
            CounterReport report = sushiReport.CounterReports[0];

            foreach (CounterCustomerReport customer in report.CustomerReports)
            {
                foreach (JournalReportItem reportItem in customer.ReportItems)
                {
                    StringBuilder line =
                    new StringBuilder(WrapComma(reportItem.ItemName) + "," + WrapComma(reportItem.ItemPublisher) + "," +
                                      WrapComma(reportItem.ItemPlatform) + "," + reportItem.DOI + "," + reportItem.Proprietary + "," +
                                      reportItem.Print_ISSN + "," + reportItem.Online_ISSN);

                    Int32 RP_Total = 0;
                    Int32 RP_Total_HTML = 0;
                    Int32 RP_Total_PDF = 0;
                    StringBuilder monthly_counts = new StringBuilder();
                    for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                    {
                        DateTime start = new DateTime(currMonth.Year, currMonth.Month, 1);
                        DateTime end = new DateTime(currMonth.Year, currMonth.Month, DateTime.DaysInMonth(currMonth.Year, currMonth.Month));

                        bool found_ttl = false;
                        foreach (CounterPerformanceItem perfItem in reportItem.PerformanceItems)
                        {
                            CounterMetric metric;
                            if (perfItem.TryGetMetric(start, end, CounterMetricCategory.Requests, out metric))
                            {
                                foreach (var instance in metric.Instances)
                                {
                                    // ft_html
                                    if (instance.Type == CounterMetricType.ft_html)
                                    {
                                        RP_Total_HTML += Convert.ToInt32(instance.Count);
                                    }
                                    // ft_pdf
                                    if (instance.Type == CounterMetricType.ft_pdf)
                                    {
                                        RP_Total_PDF += Convert.ToInt32(instance.Count);
                                    }
                                    // ft_total
                                    if (!found_ttl && instance.Type == CounterMetricType.ft_total)
                                    {
                                        monthly_counts.Append("," + instance.Count);
                                        RP_Total += Convert.ToInt32(instance.Count);
                                        found_ttl = true;
                                    }
                                }
                            }
                        }
                        if (!found_ttl) { monthly_counts.Append(","); }
                    }

                    // Add totals and monthly counts to output line
                    line.Append("," + Convert.ToString(RP_Total));
                    line.Append("," + Convert.ToString(RP_Total_HTML));
                    line.Append("," + Convert.ToString(RP_Total_PDF));
                    line.Append(monthly_counts);

                    tw.WriteLine(line);
                }
            }
        }
        // End - ParseJR1v4

        private static void ParseJR2v4(SushiReport sushiReport, TextWriter tw)
        {
            // only do one report for now
            if (sushiReport.CounterReports.Count > 0)
            {
                CounterReport report = sushiReport.CounterReports[0];

                foreach (CounterCustomerReport customer in report.CustomerReports)
                {
                    foreach (JournalReportItem reportItem in customer.ReportItems)
                    {
                        // 2 output lines : turnaway and no_license
                        StringBuilder turnaway_line =
                            new StringBuilder(WrapComma(reportItem.ItemName) + "," + WrapComma(reportItem.ItemPublisher) + "," +
                                              WrapComma(reportItem.ItemPlatform) + "," + reportItem.DOI + "," + reportItem.Proprietary + "," +
                                              reportItem.Print_ISSN + "," + reportItem.Online_ISSN +
                                              ",Access denied: concurrent/simultaneous user license limit exceeded");
                        StringBuilder no_license_line =
                            new StringBuilder(WrapComma(reportItem.ItemName) + "," + WrapComma(reportItem.ItemPublisher) + "," +
                                              WrapComma(reportItem.ItemPlatform) + "," + reportItem.DOI + "," + reportItem.Proprietary + "," +
                                              reportItem.Print_ISSN + "," + reportItem.Online_ISSN +
                                              ",Access denied: content item not licensed");

                        Int32 RP_Total_TA = 0;
                        Int32 RP_Total_NL = 0;

                        StringBuilder monthly_counts_TA = new StringBuilder();
                        StringBuilder monthly_counts_NL = new StringBuilder();
                        for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                        {
                            DateTime start = new DateTime(currMonth.Year, currMonth.Month, 1);
                            DateTime end = new DateTime(currMonth.Year, currMonth.Month, DateTime.DaysInMonth(currMonth.Year, currMonth.Month));
                            bool foundNLCount = false;
                            bool foundTACount = false;
                            foreach (CounterPerformanceItem perfItem in reportItem.PerformanceItems)
                            {
                                CounterMetric metric;
                                if (perfItem.TryGetMetric(start, end, CounterMetricCategory.Access_Denied, out metric))
                                {
                                    foreach (var instance in metric.Instances)
                                    {
                                        // get turnaway
                                        if (!foundTACount && instance.Type == CounterMetricType.turnaway)
                                        {
                                            monthly_counts_TA.Append("," + instance.Count);
                                            RP_Total_TA += Convert.ToInt32(instance.Count);
                                            foundTACount = true;
                                        }
                                        // get no_license
                                        if (!foundNLCount && instance.Type == CounterMetricType.no_license)
                                        {
                                            monthly_counts_NL.Append("," + instance.Count);
                                            RP_Total_NL += Convert.ToInt32(instance.Count);
                                            foundNLCount = true;
                                        }
                                    }
                                }
                            }
                            if (!foundTACount) { monthly_counts_TA.Append(","); }
                            if (!foundNLCount) { monthly_counts_NL.Append(","); }
                        }

                        // Tack on totals and monthly counts
                        turnaway_line.Append("," + Convert.ToString(RP_Total_TA));
                        turnaway_line.Append(monthly_counts_TA);
                        no_license_line.Append("," + Convert.ToString(RP_Total_NL));
                        no_license_line.Append(monthly_counts_NL);

                        // Print them
                        tw.WriteLine(turnaway_line);
                        tw.WriteLine(no_license_line);
                    }
                }
            }
        }
        // End - ParseJR2v4

        private static void ParseJR5v4(SushiReport sushiReport, TextWriter tw)
        {
            CounterReport report = sushiReport.CounterReports[0];

            // Build a sorted (DESC) list of Year-of-Publication (YOP) values from the report
            // to use for column headings before processing the report rows
            //
            var Years = new List<string>();
            foreach (CounterCustomerReport customer in report.CustomerReports)
            { 
                foreach (JournalReportItem reportItem in customer.ReportItems)
                {
                    foreach (CounterPerformanceItem perfItem in reportItem.PerformanceItems)
                    {
                        if ((perfItem.YOP != null) & !Years.Contains(perfItem.YOP))
                        {
                            Years.Add(perfItem.YOP);
                        }
                    }
                }
            }

            // Sort Years list descending
            Years.Sort(delegate (string x, string y)
            {
                return y.CompareTo(x);
            });

            // Build and print out column headers
            // 
            var header = new StringBuilder(",Publisher,Platform,Journal DOI,Proprietary Identifier,Print ISSN,Online ISSN");
            foreach (string YOP in Years)
            {
                header.Append(",YOP " + YOP);
            }
            tw.WriteLine(header);

            // Process thhe report Rows
            //
            DateTime start = new DateTime(StartDate.Year, StartDate.Month, 1);
            DateTime end = new DateTime(EndDate.Year, EndDate.Month, DateTime.DaysInMonth(EndDate.Year, EndDate.Month));

            foreach (CounterCustomerReport customer in report.CustomerReports)
            {
                foreach (JournalReportItem reportItem in customer.ReportItems)
                {
                    MonthData = new Dictionary<string, string>();
                    StringBuilder line =
                    new StringBuilder(WrapComma(reportItem.ItemName) + "," + WrapComma(reportItem.ItemPublisher) + "," +
                                  WrapComma(reportItem.ItemPlatform) + "," + reportItem.DOI + "," + reportItem.Proprietary + "," +
                                  reportItem.Print_ISSN + "," + reportItem.Online_ISSN);

                    foreach (CounterPerformanceItem perfItem in reportItem.PerformanceItems)
                    {
                        if ( perfItem.YOP != "" )
                        {
                            CounterMetric metric;
                            if (perfItem.TryGetMetric(start, end, CounterMetricCategory.Requests, out metric))
                            {
                                foreach (var instance in metric.Instances)
                                {
                                    // ft_total only
                                    if (instance.Type == CounterMetricType.ft_total)
                                    {
                                        MonthData.Add(perfItem.YOP, Convert.ToString(instance.Count));
                                    }
                                }
                            }
                        }
                    }

                    // Walk the MonthData by YOP to build output line, then print it
                    //
                    foreach (string YOP in Years)
                    {
                        if (MonthData.ContainsKey(YOP))
                        {
                            line.Append("," + MonthData[YOP]);
                        }
                        else
                        {
                            line.Append(",");
                        }
                    }
                    tw.WriteLine(line);
                }
            }
        }
        // End - ParseJR5v4

        private static void ParseDB1v4(SushiReport sushiReport, TextWriter tw)
        {
            // only do one report for now
            if (sushiReport.CounterReports.Count > 0)
            {
                CounterReport report = sushiReport.CounterReports[0];
                foreach (CounterCustomerReport customer in report.CustomerReports)
                {
                    foreach (CounterReportItem reportItem in customer.ReportItems)
                    {
                        StringBuilder database =
                            new StringBuilder(WrapComma(reportItem.ItemName) + "," + WrapComma(reportItem.ItemPublisher) + "," +
                                          WrapComma(reportItem.ItemPlatform));

                        // 4 output lines per database
                        StringBuilder line_search_reg = new StringBuilder(database + "," + "Regular Searches");
                        StringBuilder line_search_fed = new StringBuilder(database + "," + "Searches-federated and automated");
                        StringBuilder line_result_clicks = new StringBuilder(database + "," + "Result Clicks");
                        StringBuilder line_record_views = new StringBuilder(database + "," + "Record Views");

                        // strings for holding monthly counts
                        StringBuilder searches_reg = new StringBuilder();
                        StringBuilder searches_fed = new StringBuilder();
                        StringBuilder result_clicks = new StringBuilder();
                        StringBuilder record_views = new StringBuilder();

                        Int32 RP_Total_RegSearch = 0;
                        Int32 RP_Total_FedSearch = 0;
                        Int32 RP_Total_ResClick = 0;
                        Int32 RP_Total_RecView = 0;
                        for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                        {
                            DateTime start = new DateTime(currMonth.Year, currMonth.Month, 1);
                            DateTime end = new DateTime(currMonth.Year, currMonth.Month, DateTime.DaysInMonth(currMonth.Year, currMonth.Month));
                            bool found_reg_count = false;
                            bool found_fed_count = false;
                            bool found_rc_count = false;
                            bool found_rv_count = false;

                            foreach (CounterPerformanceItem perfItem in reportItem.PerformanceItems)
                            {
                                CounterMetric metric;

                                if (perfItem.TryGetMetric(start, end, CounterMetricCategory.Searches, out metric))
                                {
                                    foreach (var instance in metric.Instances)
                                    {
                                        // search_reg
                                        if (instance.Type == CounterMetricType.search_reg)
                                        {
                                            searches_reg.Append("," + instance.Count);
                                            RP_Total_RegSearch += Convert.ToInt32(instance.Count);
                                            found_reg_count = true;
                                        }
                                        // search_fed
                                        if (instance.Type == CounterMetricType.search_fed)
                                        {
                                            searches_fed.Append("," + instance.Count);
                                            RP_Total_FedSearch += Convert.ToInt32(instance.Count);
                                            found_fed_count = true;
                                        }
                                    }
                                }
                                if (perfItem.TryGetMetric(start, end, CounterMetricCategory.Requests, out metric))
                                {
                                    foreach (var instance in metric.Instances)
                                    {
                                        // result_click
                                        if (instance.Type == CounterMetricType.result_click)
                                        {
                                            result_clicks.Append("," + instance.Count);
                                            RP_Total_ResClick += Convert.ToInt32(instance.Count);
                                            found_rc_count = true;
                                        }
                                        // record_view
                                        if (instance.Type == CounterMetricType.record_view)
                                        {
                                            record_views.Append("," + instance.Count);
                                            RP_Total_RecView += Convert.ToInt32(instance.Count);
                                            found_rv_count = true;
                                        }
                                    }
                                }
                            }
                            if (!found_reg_count) { searches_reg.Append(","); }
                            if (!found_fed_count) { searches_fed.Append(","); }
                            if (!found_rc_count) { result_clicks.Append(","); }
                            if (!found_rv_count) { record_views.Append(","); }
                        }

                        // Tack on totals and monthly counts
                        line_search_reg.Append("," + Convert.ToString(RP_Total_RegSearch));
                        line_search_reg.Append(searches_reg);
                        line_search_fed.Append("," + Convert.ToString(RP_Total_FedSearch));
                        line_search_fed.Append(searches_fed);
                        line_result_clicks.Append("," + Convert.ToString(RP_Total_ResClick));
                        line_result_clicks.Append(result_clicks);
                        line_record_views.Append("," + Convert.ToString(RP_Total_RecView));
                        line_record_views.Append(record_views);

                        // Print them
                        tw.WriteLine(line_search_reg);
                        tw.WriteLine(line_search_fed);
                        tw.WriteLine(line_result_clicks);
                        tw.WriteLine(line_record_views);

                    }
                }
            }
        }
        // End - ParseDB1v4

        private static void ParseDB2v4(SushiReport sushiReport, TextWriter tw)
        {
            // only do one report for now
            if (sushiReport.CounterReports.Count > 0)
            {
                CounterReport report = sushiReport.CounterReports[0];
                foreach (CounterCustomerReport customer in report.CustomerReports)
                {
                    foreach (CounterReportItem reportItem in customer.ReportItems)
                    {
                        // 2 output lines : turnaway and no_license
                        StringBuilder turnaway_line =
                            new StringBuilder(WrapComma(reportItem.ItemName) + "," + WrapComma(reportItem.ItemPublisher) + "," +
                                              WrapComma(reportItem.ItemPlatform) +
                                              ",Access denied: concurrent/simultaneous user license limit exceeded");
                        StringBuilder no_license_line =
                            new StringBuilder(WrapComma(reportItem.ItemName) + "," + WrapComma(reportItem.ItemPublisher) + "," +
                                              WrapComma(reportItem.ItemPlatform) +
                                              ",Access denied: content item not licensed");

                        Int32 RP_Total_TA = 0;
                        Int32 RP_Total_NL = 0;
                        StringBuilder monthly_counts_TA = new StringBuilder();
                        StringBuilder monthly_counts_NL = new StringBuilder();

                        for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                        {
                            DateTime start = new DateTime(currMonth.Year, currMonth.Month, 1);
                            DateTime end = new DateTime(currMonth.Year, currMonth.Month, DateTime.DaysInMonth(currMonth.Year, currMonth.Month));
                            bool foundNLCount = false;
                            bool foundTACount = false;
                            foreach (CounterPerformanceItem perfItem in reportItem.PerformanceItems)
                            {
                                CounterMetric metric;

                                if (perfItem.TryGetMetric(start, end, CounterMetricCategory.Access_Denied, out metric))
                                {
                                    foreach (var instance in metric.Instances)
                                    {
                                        // get turnaway
                                        if (!foundTACount && instance.Type == CounterMetricType.turnaway)
                                        {
                                            monthly_counts_TA.Append("," + instance.Count);
                                            RP_Total_TA += Convert.ToInt32(instance.Count);
                                            foundTACount = true;
                                        }

                                        // get no_license
                                        if (!foundNLCount && instance.Type == CounterMetricType.no_license)
                                        {
                                            monthly_counts_NL.Append("," + instance.Count);
                                            RP_Total_NL += Convert.ToInt32(instance.Count);
                                            foundNLCount = true;
                                        }
                                    }
                                }
                            }
                            if (!foundTACount) { monthly_counts_TA.Append(","); }
                            if (!foundNLCount) { monthly_counts_NL.Append(","); }
                        }

                        // Tack on totals and monthly counts
                        turnaway_line.Append("," + Convert.ToString(RP_Total_TA));
                        turnaway_line.Append(monthly_counts_TA);
                        no_license_line.Append("," + Convert.ToString(RP_Total_NL));
                        no_license_line.Append(monthly_counts_NL);

                        // Print them
                        tw.WriteLine(turnaway_line);
                        tw.WriteLine(no_license_line);
                    }
                }
            }
        }
        // End - ParseDB2v4

        private static void ParsePR1v4(SushiReport sushiReport, TextWriter tw)
        {
            // only do one report for now
            if (sushiReport.CounterReports.Count > 0)
            {
                CounterReport report = sushiReport.CounterReports[0];
                foreach (CounterCustomerReport customer in report.CustomerReports)
                {
                    foreach (CounterReportItem reportItem in customer.ReportItems)
                    {
                        StringBuilder database =
                        new StringBuilder(WrapComma(reportItem.ItemPlatform) + "," + WrapComma(reportItem.ItemPublisher));

                        // 4 output lines per database
                        StringBuilder line_search_reg = new StringBuilder(database + "," + "Regular Searches");
                        StringBuilder line_search_fed = new StringBuilder(database + "," + "Searches-federated and automated");
                        StringBuilder line_result_clicks = new StringBuilder(database + "," + "Result Clicks");
                        StringBuilder line_record_views = new StringBuilder(database + "," + "Record Views");

                        // strings for holding monthly counts
                        StringBuilder searches_reg = new StringBuilder();
                        StringBuilder searches_fed = new StringBuilder();
                        StringBuilder result_clicks = new StringBuilder();
                        StringBuilder record_views = new StringBuilder();

                        Int32 RP_Total_RegSearch = 0;
                        Int32 RP_Total_FedSearch = 0;
                        Int32 RP_Total_ResClick = 0;
                        Int32 RP_Total_RecView = 0;

                        for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                        {
                            DateTime start = new DateTime(currMonth.Year, currMonth.Month, 1);
                            DateTime end = new DateTime(currMonth.Year, currMonth.Month, DateTime.DaysInMonth(currMonth.Year, currMonth.Month));

                            bool found_reg_count = false;
                            bool found_fed_count = false;
                            bool found_rc_count = false;
                            bool found_rv_count = false;
                            foreach (CounterPerformanceItem perfItem in reportItem.PerformanceItems)
                            {
                                CounterMetric metric;

                                if (perfItem.TryGetMetric(start, end, CounterMetricCategory.Searches, out metric))
                                {
                                    foreach (var instance in metric.Instances)
                                    {
                                        // search_reg
                                        if (instance.Type == CounterMetricType.search_reg)
                                        {
                                            searches_reg.Append("," + instance.Count);
                                            RP_Total_RegSearch += Convert.ToInt32(instance.Count);
                                            found_reg_count = true;
                                        }
                                        // search_fed
                                        if (instance.Type == CounterMetricType.search_fed)
                                        {
                                            searches_fed.Append("," + instance.Count);
                                            RP_Total_FedSearch += Convert.ToInt32(instance.Count);
                                            found_fed_count = true;
                                        }
                                    }
                                }
                                if (perfItem.TryGetMetric(start, end, CounterMetricCategory.Requests, out metric))
                                {
                                    foreach (var instance in metric.Instances)
                                    {
                                        // result_click
                                        if (instance.Type == CounterMetricType.result_click)
                                        {
                                            result_clicks.Append("," + instance.Count);
                                            RP_Total_ResClick += Convert.ToInt32(instance.Count);
                                            found_rc_count = true;
                                        }
                                        // record_view
                                        if (instance.Type == CounterMetricType.record_view)
                                        {
                                            record_views.Append("," + instance.Count);
                                            RP_Total_RecView += Convert.ToInt32(instance.Count);
                                            found_rv_count = true;
                                        }
                                    }
                                }
                            }
                            if (!found_reg_count) { searches_reg.Append(","); }
                            if (!found_fed_count) { searches_fed.Append(","); }
                            if (!found_rc_count) { result_clicks.Append(","); }
                            if (!found_rv_count) { record_views.Append(","); }
                        }

                        // Tack on totals and monthly counts
                        line_search_reg.Append("," + Convert.ToString(RP_Total_RegSearch));
                        line_search_reg.Append(searches_reg);
                        line_search_fed.Append("," + Convert.ToString(RP_Total_FedSearch));
                        line_search_fed.Append(searches_fed);
                        line_result_clicks.Append("," + Convert.ToString(RP_Total_ResClick));
                        line_result_clicks.Append(result_clicks);
                        line_record_views.Append("," + Convert.ToString(RP_Total_RecView));
                        line_record_views.Append(record_views);

                        // Print them
                        tw.WriteLine(line_search_reg);
                        tw.WriteLine(line_search_fed);
                        tw.WriteLine(line_result_clicks);
                        tw.WriteLine(line_record_views);

                    }
                }
            }
        }
        // End - ParsePR1v4

        private static void ParseBR12v4(SushiReport sushiReport, TextWriter tw)
        {
            // This function handles BR1 and BR2
            //
            // only do one report for now
            if (sushiReport.CounterReports.Count > 0)
            {
                CounterReport report = sushiReport.CounterReports[0];
                foreach (CounterCustomerReport customer in report.CustomerReports)
                {
                    foreach (BookReportItem reportItem in customer.ReportItems)
                    {
                        StringBuilder line =
                        new StringBuilder(WrapComma(reportItem.ItemName) + "," + WrapComma(reportItem.ItemPublisher) + "," +
                                          WrapComma(reportItem.ItemPlatform) + "," + reportItem.DOI + "," + "," + reportItem.Proprietary +
                                          reportItem.Print_ISBN + "," + reportItem.Online_ISSN);
                        StringBuilder title_counts = new StringBuilder();

                        Int32 RP_Total = 0;

                        for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                        {
                            DateTime start = new DateTime(currMonth.Year, currMonth.Month, 1);
                            DateTime end = new DateTime(currMonth.Year, currMonth.Month, DateTime.DaysInMonth(currMonth.Year, currMonth.Month));

                            bool foundCount = false;
                            foreach (CounterPerformanceItem perfItem in reportItem.PerformanceItems)
                            {
                                CounterMetric metric;

                                if (perfItem.TryGetMetric(start, end, CounterMetricCategory.Requests, out metric))
                                {
                                    foreach (var instance in metric.Instances)
                                    {
                                        // get ft_total only
                                        if (!foundCount && instance.Type == CounterMetricType.ft_total)
                                        {
                                            RP_Total += Convert.ToInt32(instance.Count);
                                            title_counts.Append("," + instance.Count);
                                            foundCount = true;
                                        }
                                    }

                                }
                            }
                            if (!foundCount) { title_counts.Append(","); }
                        }

                        // Tack on totals and monthly counts and print it
                        line.Append("," + Convert.ToString(RP_Total));
                        line.Append(title_counts);
                        tw.WriteLine(line);
                    }
                }
            }
        }
        // End - ParseBR12v4

        private static void ParseBR3v4(SushiReport sushiReport, TextWriter tw)
        {
            // only do one report for now
            if (sushiReport.CounterReports.Count > 0)
            {
                CounterReport report = sushiReport.CounterReports[0];

                foreach (CounterCustomerReport customer in report.CustomerReports)
                {
                    foreach (BookReportItem reportItem in customer.ReportItems)
                    {
                        // 2 output lines : turnaway and no_license
                        StringBuilder turnaway_line =
                            new StringBuilder(WrapComma(reportItem.ItemName) + "," + WrapComma(reportItem.ItemPublisher) + "," +
                                              WrapComma(reportItem.ItemPlatform) + "," + reportItem.DOI + "," + reportItem.Proprietary + "," +
                                              reportItem.Print_ISBN + "," + reportItem.Online_ISSN +
                                              ",Access denied: concurrent/simultaneous user license limit exceeded");
                        StringBuilder no_license_line =
                            new StringBuilder(WrapComma(reportItem.ItemName) + "," + WrapComma(reportItem.ItemPublisher) + "," +
                                              WrapComma(reportItem.ItemPlatform) + "," + reportItem.DOI + "," + reportItem.Proprietary + "," +
                                              reportItem.Print_ISBN + "," + reportItem.Online_ISSN +
                                              ",Access denied: content item not licensed");

                        Int32 RP_Total_TA = 0;
                        Int32 RP_Total_NL = 0;

                        StringBuilder monthly_counts_TA = new StringBuilder();
                        StringBuilder monthly_counts_NL = new StringBuilder();
                        for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                        {
                            DateTime start = new DateTime(currMonth.Year, currMonth.Month, 1);
                            DateTime end = new DateTime(currMonth.Year, currMonth.Month, DateTime.DaysInMonth(currMonth.Year, currMonth.Month));
                            bool foundNLCount = false;
                            bool foundTACount = false;
                            foreach (CounterPerformanceItem perfItem in reportItem.PerformanceItems)
                            {
                                CounterMetric metric;
                                if (perfItem.TryGetMetric(start, end, CounterMetricCategory.Access_Denied, out metric))
                                {
                                    foreach (var instance in metric.Instances)
                                    {
                                        // get turnaway
                                        if (!foundTACount && instance.Type == CounterMetricType.turnaway)
                                        {
                                            monthly_counts_TA.Append("," + instance.Count);
                                            RP_Total_TA += Convert.ToInt32(instance.Count);
                                            foundTACount = true;
                                        }
                                        // get no_license
                                        if (!foundNLCount && instance.Type == CounterMetricType.no_license)
                                        {
                                            monthly_counts_NL.Append("," + instance.Count);
                                            RP_Total_NL += Convert.ToInt32(instance.Count);
                                            foundNLCount = true;
                                        }
                                    }
                                }
                            }
                            if (!foundTACount) { monthly_counts_TA.Append(","); }
                            if (!foundNLCount) { monthly_counts_NL.Append(","); }
                        }

                        // Tack on totals and monthly counts
                        turnaway_line.Append("," + Convert.ToString(RP_Total_TA));
                        turnaway_line.Append(monthly_counts_TA);
                        no_license_line.Append("," + Convert.ToString(RP_Total_NL));
                        no_license_line.Append(monthly_counts_NL);

                        // Print them
                        tw.WriteLine(turnaway_line);
                        tw.WriteLine(no_license_line);
                    }
                }
            }
        }
        // End - ParseBR3v4

        private static void ParseBR4v4(SushiReport sushiReport, TextWriter tw)
        {
            // only do one report for now
            if (sushiReport.CounterReports.Count > 0)
            {
                CounterReport report = sushiReport.CounterReports[0];

                foreach (CounterCustomerReport customer in report.CustomerReports)
                {
                    foreach (CounterReportItem reportItem in customer.ReportItems)
                    {
                        // 2 output lines : turnaway and no_license
                        StringBuilder turnaway_line =
                            new StringBuilder(WrapComma(reportItem.ItemName) + "," + WrapComma(reportItem.ItemPublisher) + "," +
                                              WrapComma(reportItem.ItemPlatform) + "," +
                                              ",Access denied: concurrent/simultaneous user license limit exceeded");
                        StringBuilder no_license_line =
                            new StringBuilder(WrapComma(reportItem.ItemName) + "," + WrapComma(reportItem.ItemPublisher) + "," +
                                              WrapComma(reportItem.ItemPlatform) + "," + 
                                              ",Access denied: content item not licensed");

                        Int32 RP_Total_TA = 0;
                        Int32 RP_Total_NL = 0;

                        StringBuilder monthly_counts_TA = new StringBuilder();
                        StringBuilder monthly_counts_NL = new StringBuilder();
                        for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                        {
                            DateTime start = new DateTime(currMonth.Year, currMonth.Month, 1);
                            DateTime end = new DateTime(currMonth.Year, currMonth.Month, DateTime.DaysInMonth(currMonth.Year, currMonth.Month));
                            bool foundNLCount = false;
                            bool foundTACount = false;
                            foreach (CounterPerformanceItem perfItem in reportItem.PerformanceItems)
                            {
                                CounterMetric metric;
                                if (perfItem.TryGetMetric(start, end, CounterMetricCategory.Access_Denied, out metric))
                                {
                                    foreach (var instance in metric.Instances)
                                    {
                                        // get turnaway
                                        if (!foundTACount && instance.Type == CounterMetricType.turnaway)
                                        {
                                            monthly_counts_TA.Append("," + instance.Count);
                                            RP_Total_TA += Convert.ToInt32(instance.Count);
                                            foundTACount = true;
                                        }
                                        // get no_license
                                        if (!foundNLCount && instance.Type == CounterMetricType.no_license)
                                        {
                                            monthly_counts_NL.Append("," + instance.Count);
                                            RP_Total_NL += Convert.ToInt32(instance.Count);
                                            foundNLCount = true;
                                        }
                                    }
                                }
                            }
                            if (!foundTACount) { monthly_counts_TA.Append(","); }
                            if (!foundNLCount) { monthly_counts_NL.Append(","); }
                        }

                        // Tack on totals and monthly counts
                        turnaway_line.Append("," + Convert.ToString(RP_Total_TA));
                        turnaway_line.Append(monthly_counts_TA);
                        no_license_line.Append("," + Convert.ToString(RP_Total_NL));
                        no_license_line.Append(monthly_counts_NL);

                        // Print them
                        tw.WriteLine(turnaway_line);
                        tw.WriteLine(no_license_line);
                    }
                }
            }
        }
        // End - ParseBR4v4

        private static void ParseBR5v4(SushiReport sushiReport, TextWriter tw)
        {
            // only do one report for now
            if (sushiReport.CounterReports.Count > 0)
            {
                CounterReport report = sushiReport.CounterReports[0];

                foreach (CounterCustomerReport customer in report.CustomerReports)
                {
                    foreach (BookReportItem reportItem in customer.ReportItems)
                    {
                        StringBuilder database =
                            new StringBuilder(WrapComma(reportItem.ItemName) + "," + WrapComma(reportItem.ItemPublisher) + "," +
                                              WrapComma(reportItem.ItemPlatform) + "," + reportItem.DOI + "," + reportItem.Proprietary + "," +
                                              reportItem.Print_ISBN + "," + reportItem.Online_ISSN);

                        // 2 output lines per title
                        StringBuilder line_search_reg = new StringBuilder(database + "," + "Regular Searches");
                        StringBuilder line_search_fed = new StringBuilder(database + "," + "Searches: federated and automated");

                        // strings for holding monthly counts
                        StringBuilder searches_reg = new StringBuilder();
                        StringBuilder searches_fed = new StringBuilder();

                        Int32 RP_Total_RegSearch = 0;
                        Int32 RP_Total_FedSearch = 0;

                        for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                        {
                            DateTime start = new DateTime(currMonth.Year, currMonth.Month, 1);
                            DateTime end = new DateTime(currMonth.Year, currMonth.Month, DateTime.DaysInMonth(currMonth.Year, currMonth.Month));
                            bool found_reg_count = false;
                            bool found_fed_count = false;
                            foreach (CounterPerformanceItem perfItem in reportItem.PerformanceItems)
                            {
                                CounterMetric metric;
                                if (perfItem.TryGetMetric(start, end, CounterMetricCategory.Searches, out metric))
                                {
                                    foreach (var instance in metric.Instances)
                                    {
                                        // search_reg
                                        if (instance.Type == CounterMetricType.search_reg)
                                        {
                                            searches_reg.Append("," + instance.Count);
                                            RP_Total_RegSearch += Convert.ToInt32(instance.Count);
                                            found_reg_count = true;
                                        }
                                        // search_fed
                                        if (instance.Type == CounterMetricType.search_fed)
                                        {
                                            searches_fed.Append("," + instance.Count);
                                            RP_Total_FedSearch += Convert.ToInt32(instance.Count);
                                            found_fed_count = true;
                                        }
                                    }
                                }
                            }
                            if (!found_reg_count) { searches_reg.Append(","); }
                            if (!found_fed_count) { searches_fed.Append(","); }
                        }

                        // Tack on totals and monthly counts
                        line_search_reg.Append("," + Convert.ToString(RP_Total_RegSearch));
                        line_search_reg.Append(searches_reg);
                        line_search_fed.Append("," + Convert.ToString(RP_Total_FedSearch));
                        line_search_fed.Append(searches_fed);

                        // Print them
                        tw.WriteLine(line_search_reg);
                        tw.WriteLine(line_search_fed);
                    }
                }
            }
        }
        // End - ParseBR5v4

        private static void ParseMR1v4(SushiReport sushiReport, TextWriter tw)
        {
            // only do one report for now
            if (sushiReport.CounterReports.Count > 0)
            {
                CounterReport report = sushiReport.CounterReports[0];
                foreach (CounterCustomerReport customer in report.CustomerReports)
                {
                    foreach (BookReportItem reportItem in customer.ReportItems)
                    {
                        StringBuilder line =
                        new StringBuilder(WrapComma(reportItem.ItemName) + "," + WrapComma(reportItem.ItemPublisher) + "," +
                                          WrapComma(reportItem.ItemPlatform));
                        StringBuilder title_counts = new StringBuilder();

                        Int32 RP_Total = 0;
                        for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                        {
                            DateTime start = new DateTime(currMonth.Year, currMonth.Month, 1);
                            DateTime end = new DateTime(currMonth.Year, currMonth.Month, DateTime.DaysInMonth(currMonth.Year, currMonth.Month));

                            bool foundCount = false;
                            foreach (CounterPerformanceItem perfItem in reportItem.PerformanceItems)
                            {
                                CounterMetric metric;

                                if (perfItem.TryGetMetric(start, end, CounterMetricCategory.Requests, out metric))
                                {
                                    foreach (var instance in metric.Instances)
                                    {
                                        // get ft_total only
                                        if (!foundCount && instance.Type == CounterMetricType.ft_total)
                                        {
                                            RP_Total += Convert.ToInt32(instance.Count);
                                            title_counts.Append("," + instance.Count);
                                            foundCount = true;
                                        }
                                    }
                                }
                            }
                            if (!foundCount) { title_counts.Append(","); }
                        }

                        // Tack on totals and monthly counts and print it
                        line.Append("," + Convert.ToString(RP_Total));
                        line.Append(title_counts);
                        tw.WriteLine(line);
                    }
                }
            }
        }
        // End - ParseMR1v4

        private static void ParseCR1v4(SushiReport sushiReport, TextWriter tw)
        {
            CounterReport report = sushiReport.CounterReports[0];
            foreach (CounterCustomerReport customer in report.CustomerReports)
            {
                foreach (ConsortiumReportItem reportItem in customer.ReportItems)
                {
                    StringBuilder line =
                    new StringBuilder(WrapComma(customer.Customer_ID) + "," + WrapComma(reportItem.ItemPublisher) + "," +
                                      WrapComma(reportItem.ItemPlatform) + "," + WrapComma(reportItem.ItemName) + "," +
                                      reportItem.DOI + "," + reportItem.Proprietary + "," +
                                      reportItem.Print_ISSN + "," + reportItem.Online_ISSN );

                    Int32 RP_Total = 0;
                    Int32 RP_Total_HTML = 0;
                    Int32 RP_Total_PDF = 0;
                    StringBuilder monthly_counts = new StringBuilder();

                    for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                    {
                        DateTime start = new DateTime(currMonth.Year, currMonth.Month, 1);
                        DateTime end = new DateTime(currMonth.Year, currMonth.Month, DateTime.DaysInMonth(currMonth.Year, currMonth.Month));

                        bool foundCount = false;
                        foreach (CounterPerformanceItem perfItem in reportItem.PerformanceItems)
                        {
                            CounterMetric metric;
                            if (perfItem.TryGetMetric(start, end, CounterMetricCategory.Requests, out metric))
                            {
                                foreach (var instance in metric.Instances)
                                {
                                    // ft_html
                                    if (instance.Type == CounterMetricType.ft_html)
                                    {
                                        RP_Total_HTML += Convert.ToInt32(instance.Count);
                                    }
                                    // ft_pdf
                                    if (instance.Type == CounterMetricType.ft_pdf)
                                    {
                                        RP_Total_PDF += Convert.ToInt32(instance.Count);
                                    }
                                    // ft_total
                                    if (!foundCount && instance.Type == CounterMetricType.ft_total)
                                    {
                                        monthly_counts.Append("," + instance.Count);
                                        RP_Total += Convert.ToInt32(instance.Count);
                                        foundCount = true;
                                    }
                                }
                            }
                        }
                        if (!foundCount) { monthly_counts.Append(","); }
                    }

                    // Add totals and monthly counts to output line
                    line.Append("," + Convert.ToString(RP_Total));
                    line.Append("," + Convert.ToString(RP_Total_HTML));
                    line.Append("," + Convert.ToString(RP_Total_PDF));
                    line.Append(monthly_counts);

                    tw.WriteLine(line);

                }
            }
        }
        // End - ParseCR1v4

        private static void ParseCR2v4(SushiReport sushiReport, TextWriter tw)
        {
            // only do one report for now
            if (sushiReport.CounterReports.Count > 0)
            {
                CounterReport report = sushiReport.CounterReports[0];
                foreach (CounterCustomerReport customer in report.CustomerReports)
                {
                    foreach (CounterReportItem reportItem in customer.ReportItems)
                    {
                        StringBuilder database =
                            new StringBuilder(WrapComma(customer.Customer_Name) + "," + WrapComma(reportItem.ItemName) + "," +
                                              WrapComma(reportItem.ItemPublisher) + "," + WrapComma(reportItem.ItemPlatform));

                        // 4 output lines per database
                        StringBuilder line_search_reg = new StringBuilder(database + "," + "Regular Searches");
                        StringBuilder line_search_fed = new StringBuilder(database + "," + "Searches-federated and automated");
                        StringBuilder line_result_clicks = new StringBuilder(database + "," + "Result Clicks");
                        StringBuilder line_record_views = new StringBuilder(database + "," + "Record Views");

                        // strings for holding monthly counts
                        StringBuilder searches_reg = new StringBuilder();
                        StringBuilder searches_fed = new StringBuilder();
                        StringBuilder result_clicks = new StringBuilder();
                        StringBuilder record_views = new StringBuilder();

                        Int32 RP_Total_RegSearch = 0;
                        Int32 RP_Total_FedSearch = 0;
                        Int32 RP_Total_ResClick = 0;
                        Int32 RP_Total_RecView = 0;

                        for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                        {
                            DateTime start = new DateTime(currMonth.Year, currMonth.Month, 1);
                            DateTime end = new DateTime(currMonth.Year, currMonth.Month, DateTime.DaysInMonth(currMonth.Year, currMonth.Month));

                            bool found_reg_count = false;
                            bool found_fed_count = false;
                            bool found_rc_count = false;
                            bool found_rv_count = false;
                            foreach (CounterPerformanceItem perfItem in reportItem.PerformanceItems)
                            {
                                CounterMetric metric;
                                if (perfItem.TryGetMetric(start, end, CounterMetricCategory.Searches, out metric))
                                {
                                    foreach (var instance in metric.Instances)
                                    {
                                        // search_reg
                                        if (instance.Type == CounterMetricType.search_reg)
                                        {
                                            searches_reg.Append("," + instance.Count);
                                            RP_Total_RegSearch += Convert.ToInt32(instance.Count);
                                            found_reg_count = true;
                                        }
                                        // search_fed
                                        if (instance.Type == CounterMetricType.search_fed)
                                        {
                                            searches_fed.Append("," + instance.Count);
                                            RP_Total_FedSearch += Convert.ToInt32(instance.Count);
                                            found_fed_count = true;
                                        }
                                    }
                                }
                                if (perfItem.TryGetMetric(start, end, CounterMetricCategory.Requests, out metric))
                                {
                                    foreach (var instance in metric.Instances)
                                    {
                                        // result_click
                                        if (instance.Type == CounterMetricType.result_click)
                                        {
                                            result_clicks.Append("," + instance.Count);
                                            RP_Total_ResClick += Convert.ToInt32(instance.Count);
                                            found_rc_count = true;
                                        }
                                        // record_view
                                        if (instance.Type == CounterMetricType.record_view)
                                        {
                                            record_views.Append("," + instance.Count);
                                            RP_Total_RecView += Convert.ToInt32(instance.Count);
                                            found_rv_count = true;
                                        }
                                    }
                                }
                            }
                            if (!found_reg_count) { searches_reg.Append(","); }
                            if (!found_fed_count) { searches_fed.Append(","); }
                            if (!found_rc_count) { result_clicks.Append(","); }
                            if (!found_rv_count) { record_views.Append(","); }
                        }

                        // Tack on totals and monthly counts
                        line_search_reg.Append("," + Convert.ToString(RP_Total_RegSearch));
                        line_search_reg.Append(searches_reg);
                        line_search_fed.Append("," + Convert.ToString(RP_Total_FedSearch));
                        line_search_fed.Append(searches_fed);
                        line_result_clicks.Append("," + Convert.ToString(RP_Total_ResClick));
                        line_result_clicks.Append(result_clicks);
                        line_record_views.Append("," + Convert.ToString(RP_Total_RecView));
                        line_record_views.Append(record_views);

                        // Print them
                        tw.WriteLine(line_search_reg);
                        tw.WriteLine(line_search_fed);
                        tw.WriteLine(line_result_clicks);
                        tw.WriteLine(line_record_views);

                    }
                }
            }
        }
        // End - ParseCR2v4

        private static void ParseCR3v4(SushiReport sushiReport, TextWriter tw)
        {
            // only do one report for now
            if (sushiReport.CounterReports.Count > 0)
            {
                CounterReport report = sushiReport.CounterReports[0];
                foreach (CounterCustomerReport customer in report.CustomerReports)
                {
                    foreach (BookReportItem reportItem in customer.ReportItems)
                    {
                        StringBuilder line = new StringBuilder(WrapComma(customer.Customer_Name) + "," + WrapComma(reportItem.ItemName) + "," +
                                                               WrapComma(reportItem.ItemPublisher) + "," + WrapComma(reportItem.ItemPlatform));
                        StringBuilder title_counts = new StringBuilder();

                        Int32 RP_Total = 0;
                        for (DateTime currMonth = StartDate; currMonth <= EndDate; currMonth = currMonth.AddMonths(1))
                        {
                            DateTime start = new DateTime(currMonth.Year, currMonth.Month, 1);
                            DateTime end = new DateTime(currMonth.Year, currMonth.Month, DateTime.DaysInMonth(currMonth.Year, currMonth.Month));

                            bool foundCount = false;
                            foreach (CounterPerformanceItem perfItem in reportItem.PerformanceItems)
                            {
                                CounterMetric metric;

                                if (perfItem.TryGetMetric(start, end, CounterMetricCategory.Requests, out metric))
                                {
                                    foreach (var instance in metric.Instances)
                                    {
                                        // get ft_total only
                                        if (!foundCount && instance.Type == CounterMetricType.ft_total)
                                        {
                                            RP_Total += Convert.ToInt32(instance.Count);
                                            title_counts.Append("," + instance.Count);
                                            foundCount = true;
                                        }
                                    }
                                }
                            }
                            if (!foundCount) { title_counts.Append(","); }
                        }

                        // Tack on totals and monthly counts and print it
                        line.Append("," + Convert.ToString(RP_Total));
                        line.Append(title_counts);
                        tw.WriteLine(line);
                    }
                }
            }
        }
        // End - ParseCR3v4


        // wrap string in quotes if it contains commas
        private static string WrapComma(string input)
        {
            if (input.Contains(","))
            {
                input = "\"" + input + "\"";
            }
            return input;
        }
    }
}

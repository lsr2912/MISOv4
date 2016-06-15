/*
*/

using System;
using System.Collections.Generic;
using System.Text;

namespace SushiLibrary
{
    public class CounterCustomerReport
    {
        public string Customer_ID { get; set; }
        public string Customer_Name { get; set; }
        public string Customer_Consortium_Code { get; set; }
        public string Customer_Consortium_WellKnownName { get; set; }

        public List<CounterReportItem> ReportItems { get; set; }
    }
}

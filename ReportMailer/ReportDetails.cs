using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportMailer
{
    public class ReportDetails
    {
        public string ReportName { get; set; }
        public string CardCode { get; set; }
        public string CustomerEmail { get; set; }
        public DateTime Todate { get; set; }
        public string ReportFilename { get; set; }
    }
}

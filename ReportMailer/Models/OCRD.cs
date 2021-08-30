using Dapper.Contrib.Extensions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReportMailer.Models
{
    [Table("OCRD")]
    class OCRD
    {
        [Key]
        public string CardCode { get; set; }
        public string CardType { get; set; }
        public string CardName { get; set; }
        public string U_send { get; set; }
        public string E_Mail { get; set; }

    }
}

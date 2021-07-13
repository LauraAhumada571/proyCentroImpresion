using System;
using System.Collections.Generic;
using System.Text;

namespace proyCentroImpresion.Models
{
    public class DataExtractCollection
    {
        public long id_invoice { get; set; }
        public ulong id_contract { get; set; }
        public DateTime invoice_date { get; set; }
        public DateTime expiry_date { get; set; }
        public int invoice_value { get; set; }
        public string type_person { get; set; }
        public string name { get; set; }
        public long id_customer { get; set; }
        public string address { get; set; }
        public string city { get; set; }
        public long subtotal { get; set; }
        public long iva { get; set; }
        public long total { get; set; }
        public string random_text { get; set; }
        public long previous_invoice { get; set; }
        public long penalty_interest { get; set; }
        public long collection_costs { get; set; }
        public long total_pp { get; set; }
        public long previous_balance { get; set; }
        public long actual_penalty_interest { get; set; }
        public long actual_collection_costs { get; set; }
        public long month_invoice { get; set; }
        public long total_payable { get; set; }
        public string barcode { get; set; }
        public Boolean generate_extract { get; set; }
        public string output_type { get; set; }
        public string email { get; set; }
    }
}

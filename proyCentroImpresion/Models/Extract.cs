using System;


namespace proyCentroImpresion.Models
{
    public class Extract
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
    }
}

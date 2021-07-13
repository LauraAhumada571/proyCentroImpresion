
namespace proyCentroImpresion.Models
{
    public class Statementpayment
    {
        public long id_invoice { get; set; }
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
    }

}

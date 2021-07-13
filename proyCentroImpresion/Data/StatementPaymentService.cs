using proyCentroImpresion.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;

namespace proyCentroImpresion.Data
{
    class StatementPaymentService : IStatementPayment
    {
        //cargar la información del archivo plano denominado DatosPagoExtracto
        public List<Statementpayment> GetStatementPayment()
        {
            var statement_payment= new List<Statementpayment>();

            try
            {

                StreamReader sreader = new StreamReader("E:\\DatosPagoExtracto.txt", Encoding.GetEncoding("iso-8859-1"));
                var line = sreader.ReadLine();

                while (line != null)
                {
                    string[] data = line.Trim().Split(';');
                 
                    statement_payment.Add(new Statementpayment()
                    {
                        id_invoice = long.Parse(data[0]),
                        subtotal = long.Parse(data[1]),
                        iva = long.Parse(data[2]),
                        total = long.Parse(data[3]),
                        random_text = data[4],
                        previous_invoice = long.Parse(data[5]),
                        penalty_interest = long.Parse(data[6]),
                        collection_costs = long.Parse(data[7]),
                        total_pp = long.Parse(data[8]),
                        previous_balance = long.Parse(data[9]),
                        actual_penalty_interest = long.Parse(data[10]),
                        actual_collection_costs = long.Parse(data[11]),
                        month_invoice = long.Parse(data[12]),
                        total_payable = long.Parse(data[13]),
                        barcode = data[14]
                    });


                    line = sreader.ReadLine();
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            

            return statement_payment;
        }
    }
}

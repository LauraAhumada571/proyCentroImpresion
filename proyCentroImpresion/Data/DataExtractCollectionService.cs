using proyCentroImpresion.Models;
using System;
using System.Collections.Generic;
using System.Windows;

namespace proyCentroImpresion.Data
{
    public class DataExtractCollectionService : IDataExtractCollection
    {
        //unir la información en un modelo de datos
        public List<DataExtractCollection> GetDataExtractCollection(List<Extract> extract, List<Statementpayment> statementpayment, List<WithheldExtract> withheldextract, List<OutputType> outputtype)
        {
            var dataextractcollection = new List<DataExtractCollection>();


            try
            {
                int i = 0;
                Boolean flag = true;
                string[] flag2 = new string[2];

                for (i = 0; i < extract.Count; i++)
                {
                    for (int j = 0; j < statementpayment.Count; j++)
                    {
                        if (extract[i].id_invoice == statementpayment[j].id_invoice)
                        {
                            flag = searchWithheldExtract(withheldextract, extract[i].id_invoice);
                            flag2 = defineoutputtype(outputtype, extract[i].id_invoice);
                            
                            dataextractcollection.Add(new DataExtractCollection()
                            {
                                id_invoice = extract[i].id_invoice,
                                id_contract = extract[i].id_contract,
                                invoice_date = extract[i].invoice_date,
                                expiry_date = extract[i].expiry_date,
                                invoice_value = extract[i].invoice_value,
                                type_person = extract[i].type_person,
                                name = extract[i].name,
                                id_customer = extract[i].id_customer,
                                address = extract[i].address,
                                city = extract[i].city,
                                subtotal = statementpayment[j].subtotal,
                                iva = statementpayment[j].iva,
                                total = statementpayment[j].total,
                                random_text = statementpayment[j].random_text,
                                previous_invoice = statementpayment[j].previous_invoice,
                                penalty_interest = statementpayment[j].penalty_interest,
                                collection_costs = statementpayment[j].collection_costs,
                                total_pp = statementpayment[j].total_pp,
                                previous_balance = statementpayment[j].previous_balance,
                                actual_penalty_interest = statementpayment[j].actual_penalty_interest,
                                actual_collection_costs = statementpayment[j].actual_collection_costs,
                                month_invoice = statementpayment[j].month_invoice,
                                total_payable = statementpayment[j].total_payable,
                                barcode = statementpayment[j].barcode,
                                generate_extract = flag,
                                output_type = flag2[0],
                                email = flag2[1]
                            });
                            i++;
                        }
                    }
                }

            }
            catch (Exception e)
            {
                MessageBox.Show("Error al cargar el archivo " + e);
            }

            return dataextractcollection;
        }

        //definir si se debe o no generar el extracto 
        private Boolean searchWithheldExtract(List<WithheldExtract> withheldextract, long id_invoice)
        {
            Boolean state = true;

            for (int i = 0; i < withheldextract.Count; i++)
            {
                if (withheldextract[i].id_invoice == id_invoice)
                {
                    state = false;
                }
            }
            return state;
        }

        //definir el tipo de entrega de los archivos
        private string[] defineoutputtype(List<OutputType> outputtype, long id_invoice)
        {
            string[] array = new string[2];

            for ( int i = 0; i < outputtype.Count; i++)
            {
                if (outputtype[i].id_invoice == id_invoice)
                {
                    array[0] = "send by email";
                    array[1] = outputtype[i].email;
                }
            }
            return array;
        }
    }
}

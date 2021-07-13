using proyCentroImpresion.Models;
using System.IO;
using System.Collections.Generic;
using System;
using System.Text;
using System.Windows;

namespace proyCentroImpresion.Data
{
    public class ExtractService : IExctract

    {
        //extraer la infromación del archivo plano denominado DatosExtracto
        public List<Extract> GetExtract()
        {
            var extract = new List<Extract>();
          
            try
            {
               
                StreamReader sreader = new StreamReader("E:\\DatosExtracto.txt", Encoding.GetEncoding("iso-8859-1"));
                var line = sreader.ReadLine();

                while ( line != null)
                {
                    string[] data = line.Trim().Split(';');
             
                    extract.Add(new Extract()
                    {
                        id_invoice = long.Parse(data[0]),
                        id_contract = Convert.ToUInt64(data[1]),
                        invoice_date = DateTime.ParseExact(data[2], "dd/MM/yyyy", null),
                        expiry_date = DateTime.ParseExact(data[3], "dd/MM/yyyy", null),
                        invoice_value = Int32.Parse(data[4]),
                        type_person = data[5],
                        name = data[6],
                        id_customer = long.Parse(data[7]),
                        address = data[8],
                        city = data[9]
                    }) ; 
                    
                    
                    line = sreader.ReadLine();
                }
            }
            catch(Exception e) 
            {
                MessageBox.Show("Error al cargar el archivo " + e);
            }

            return extract;

        }
    }
}

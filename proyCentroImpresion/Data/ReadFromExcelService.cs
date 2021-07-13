using proyCentroImpresion.Models;
using System;
using System.Collections.ObjectModel;
using System.Data.OleDb;
using System.Threading.Tasks;


namespace proyCentroImpresion.Data
{
    public class ReadFromExcelService
    {
        OleDbConnection Conn;
        OleDbCommand Cmd;

        //configurar ruta para cargar los datos del archivo excel denominado EXTRACTOS A RETENER
        public void GetExcelDataServiceWithheldExtract()
        {
            string ExcelFilePath = @"E:\\EXTRACTOS A RETENER.xls";
            string excelConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelFilePath + ";Extended Properties=Excel 12.0;Persist Security Info=True";
            Conn = new OleDbConnection(excelConnectionString);
        }

        //configurar ruta para cargar los datos del archivo excel denominado EXTRACTOS DE ENVIO MAIL
        public void GetExcelDataServiceOutputType()
        {
            string ExcelFilePath = @"E:\\EXTRACTOS DE ENVIO MAIL.xls";
            string excelConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + ExcelFilePath + ";Extended Properties=Excel 12.0;Persist Security Info=True";
            Conn = new OleDbConnection(excelConnectionString);
        }

        //obtener datos del archivo EXTRACTOS A RETENER
        public async Task<ObservableCollection<WithheldExtract>> GetWithheldExtractAsync()
        {
            GetExcelDataServiceWithheldExtract();
            ObservableCollection<WithheldExtract> withheldExtracts = new ObservableCollection<WithheldExtract>();

            await Conn.OpenAsync();
            Cmd = new OleDbCommand();
            Cmd.Connection = Conn;
            Cmd.CommandText = "Select * from [Hoja1$]";
            var reader = await Cmd.ExecuteReaderAsync();
            int i = 0;

            while (reader.Read())
            {
                withheldExtracts.Add(new WithheldExtract()
                {
                    id = i,
                    id_invoice = Convert.ToInt64(reader["Numero_de_Cuenta"])
                });

                i++;
            }

            reader.Close();
            Conn.Close();
            return withheldExtracts;
        }

        //obtener datos del archivo EXTRACTOS DE ENVIO MAIL
        public async Task<ObservableCollection<OutputType>> GetOutpuTtypeAsync()
        {
            GetExcelDataServiceOutputType();
            ObservableCollection<OutputType> outputype = new ObservableCollection<OutputType>();

            await Conn.OpenAsync();
            Cmd = new OleDbCommand();
            Cmd.Connection = Conn;
            Cmd.CommandText = "Select * from [Hoja1$]";
            var reader = await Cmd.ExecuteReaderAsync();
            int i = 0;


            while (reader.Read())
            {
                outputype.Add(new OutputType()
                {
                    id_invoice = Convert.ToInt64(reader["Número de Cuenta"]),
                    outputtype = "45",
                    email = reader["Correo Electronico"].ToString()
                });

                i++;
            }

            reader.Close();
            Conn.Close();
            return outputype;
        }

    }
}

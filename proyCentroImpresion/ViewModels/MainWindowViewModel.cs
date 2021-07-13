using proyCentroImpresion.Data;
using proyCentroImpresion.Models;
using proyCentroImpresion.Views;
using System.Collections.Generic;
using System.Windows.Input;



//Librería ITextSharp
using iTextSharp.text;
using iTextSharp.text.html.simpleparser;
using iTextSharp.text.pdf;

//Librería iTextSharp XML Worker
using iTextSharp.tool.xml;
using System.IO;
using System.Text;
using System.Windows;
using System;

namespace proyCentroImpresion.ViewModels
{
    public class MainWindowViewModel : ViewModelBase
    {
        #region Properties
        private List<Extract> extracts;
        public List<Extract> Extracts
        {
            get
            {
                return extracts;
            }
            set
            {
                if (extracts == value)
                {
                    return;
                }
                extracts = value;
                OnPropertyChanged("Extracts");
            }

        }

        private List<DataExtractCollection> dataextractcollections;
        public List<DataExtractCollection> Dataextractcollections
        {
            get
            {
                return dataextractcollections;
            }
            set
            {
                if (dataextractcollections == value)
                {
                    return;
                }
                dataextractcollections = value;
                OnPropertyChanged("Dataextractcollections");
            }

        }

        private List<Statementpayment> statementpayment;
        public List<Statementpayment> Statementpayment
        {
            get
            {
                return statementpayment;
            }
            set
            {
                if (statementpayment == value)
                {
                    return;
                }
                statementpayment = value;
                OnPropertyChanged("Statementpayment");
            }
        }

        private List<WithheldExtract> withheldextract;
        public List<WithheldExtract> Withheldextract
        {
            get
            {
                return withheldextract;
            }
            set
            {
                if (withheldextract == value)
                {
                    return;
                }
                withheldextract = value;
                OnPropertyChanged("Withheldextract");
            }
        }

        private List<OutputType> outputtype;
        public List<OutputType> Outputtype
        {
            get
            {
                return outputtype;
            }
            set
            {
                if (outputtype == value)
                {
                    return;
                }
                outputtype = value;
                OnPropertyChanged("Outputtype");
            }
        }

        #endregion

        #region Commands

        //se usa para mapear
        private ICommand statementpaymentCommand;

        //esta es el ICommand de la vista
        public ICommand StatementpaymentCommand
        {
            get
            {
                if (statementpaymentCommand == null)
                {
                    statementpaymentCommand = new RelayCommand(param => this.CommandExecute(), null);
                }
                return statementpaymentCommand;
            }
        }

        private ICommand generateextractsCommand;
        public ICommand GenerateextractsCommand
        {
            get
            {
                if (generateextractsCommand == null)
                {
                    generateextractsCommand = new RelayCommand(param => this.GenerateExtracts(), null);
                }
                return generateextractsCommand;
            }
        }

        private ICommand sendEmailCommand;
        public ICommand SendEmailCommand
        {
            get
            {
                if (sendEmailCommand == null)
                {
                    sendEmailCommand = new RelayCommand(param => this.sendEmail(), null);
                }
                return sendEmailCommand;
            }
        }

        #endregion

        public MainWindowViewModel()
        {
        }

        //incovando el servicio para cargar los extractos
        private List<Extract> ExtractCommandExecute()
        {
            ExtractService extractService = new ExtractService();

            var result = extractService.GetExtract();

            Extracts = new List<Extract>(result);

            return Extracts;
        }

        //invocando el servicio para cargar los pagos de los extractos
        private List<Statementpayment> StatementpaymentCommandExecute()
        {
            var statementpaymentService = new StatementPaymentService();

            var result = statementpaymentService.GetStatementPayment();

            Statementpayment = new List<Statementpayment>(result);

            return Statementpayment;

        }

        //enviando datos para unir en un modelo
        private void JoinData(List<Extract> listextracs, List<Statementpayment> liststatementpayments, List<WithheldExtract> listwithheldextract, List<OutputType> outputtype)
        {
            var dataextractcollectionService = new DataExtractCollectionService();

            var result = dataextractcollectionService.GetDataExtractCollection(listextracs, liststatementpayments, listwithheldextract, outputtype);

            try
            {
                Dataextractcollections = new List<DataExtractCollection>(result);
            }
            catch (Exception e)
            {
                MessageBox.Show("Error al subir los archivos" + e);
            }
            finally
            {
                MessageBox.Show("Archivos cargados con exito");
            }
        }
    

        private void CommandExecute()
        {
            List<Extract> extract = ExtractCommandExecute();

            List<Statementpayment> statementpayment = StatementpaymentCommandExecute();

            List<WithheldExtract> withheldextract = ReadWithheld();

            List<OutputType> outputtype = ReadOutputType();

            JoinData(extract, statementpayment, withheldextract, outputtype);
        }

        //invocando el servicio para cargar los extractos a retener
        private List<WithheldExtract> ReadWithheld()
        { 
            var withheldextract = new ReadFromExcelService();

            var result = withheldextract.GetWithheldExtractAsync();

            Withheldextract = new List<WithheldExtract>(result.Result);

            return Withheldextract;
        }

        //invocando el servicio para cargar el típo de salida de los extractos (enviar por correo o imprimir)
        private List<OutputType> ReadOutputType()
        {

            var outputtype = new ReadFromExcelService();

            var result = outputtype.GetOutpuTtypeAsync();

            Outputtype = new List<OutputType>(result.Result);

            return Outputtype;
        }

        //generando los extractos en PDF
        private void GenerateExtracts()
        {
            var data = Dataextractcollections;


            //establecer tipo de letra para los titulos
            BaseFont _title = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, true);
            iTextSharp.text.Font title = new iTextSharp.text.Font(_title, 12f, iTextSharp.text.Font.BOLD, new BaseColor(0, 0, 0));

            //establecer tipo de letra para los parrafos
            BaseFont _paragraph = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, true);
            iTextSharp.text.Font paragraph = new iTextSharp.text.Font(_paragraph, 10f, iTextSharp.text.Font.NORMAL, new BaseColor(0, 0, 0));

            BaseColor color = new BaseColor(217, 217, 217);

            var spacer = new Paragraph("")
            {
                SpacingBefore = 35f,
                SpacingAfter = 35f,
            };

            var headerTable = new PdfPTable(new[] { 50f, 50f })
            {
                //HorizontalAlignment = Rigth,
                WidthPercentage = 85,
                DefaultCell = { MinimumHeight = 22f },
            };

            var columnWidths = new[] { 50f, 50f };

            var table = new PdfPTable(columnWidths)
            {
                WidthPercentage = 85f,
                DefaultCell = { MinimumHeight = 30f }
            };

            var client = new PdfPTable(new[] { 50f, 50f })
            {
                WidthPercentage = 85,
                DefaultCell = { MinimumHeight = 22f },
            };

            var text = new PdfPTable(new[] { 85f })
            {
                WidthPercentage = 85,
                DefaultCell = { MinimumHeight = 22f },
            };

            try
            {
                for (int i = 0; i < data.Count; i++)
                {
                    //Creando la carpeta Imprimir
                    string printPath = @"E:\\imprimir";
                    if (!Directory.Exists(printPath))
                    {
                        Directory.CreateDirectory(printPath);
                    }

                    //Creando la carpeta EnviarEmail
                    string sendPath = @"E:\\enviarEmail";
                    if (!Directory.Exists(sendPath))
                    {
                        Directory.CreateDirectory(sendPath);
                    }

                    //Lugar de Destino para imprimir PDF
                    string OutputPrintPath = "E:\\imprimir\\" + data[i].id_invoice + ".pdf";

                    //Lugar de Destino para enviar PDF por correo
                    string OutputSendPath = "E:\\enviarEmail\\" + data[i].id_invoice + ".pdf";

                    //Se crea una instancia la clase a ocupar para generar el PDF y al mismo tiempo se da el formato del tipo de hoja y sus respectivas dimensiones.
                    iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(iTextSharp.text.PageSize.A4, 10f, 10f, 10f, 0f);
                    if (data[i].generate_extract == true)
                    {
                        if (dataextractcollections[i].output_type == "send by email")
                        {
                            //Se obtiene la carpeta donde se va a guardar el archivo PDF si es enviado el extracto por correo electrónico
                            PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new FileStream(OutputSendPath, FileMode.Create));
                        }
                        else
                        {
                            //Se obtiene la carpeta donde se va a guardar el archivo PDF si es impreso el extracto
                            PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new FileStream(OutputPrintPath, FileMode.Create));
                        }


                        headerTable.DeleteBodyRows();
                        client.DeleteBodyRows();
                        table.DeleteBodyRows();
                        text.DeleteBodyRows();

                        //Se abre el documento PDF que se creo previamente
                        pdfDoc.Open();


                        long sum = data[i].previous_invoice + data[i].penalty_interest + data[i].collection_costs + data[i].actual_penalty_interest + data[i].actual_collection_costs;

                        headerTable.DeleteBodyRows();

                        headerTable.AddCell(new PdfPCell(new Phrase("Logo", title)) { Border = 0, Rowspan = 4 });
                        headerTable.AddCell(new PdfPCell(new Phrase("CARVAJAL SOLUCIONES DE COMUNICACIÓN S.A.S.", paragraph)) { Border = 0, HorizontalAlignment = Element.ALIGN_RIGHT, VerticalAlignment = Element.ALIGN_MIDDLE });
                        headerTable.AddCell(new PdfPCell(new Phrase("800096812", paragraph)) { Border = 0, HorizontalAlignment = Element.ALIGN_RIGHT });
                        headerTable.AddCell(new PdfPCell(new Phrase("Calle 29 Norte No. 6A-40 Cali", paragraph)) { Border = 0, HorizontalAlignment = Element.ALIGN_RIGHT });
                        headerTable.AddCell(new PdfPCell(new Phrase("Extracto No." + data[i].id_invoice.ToString(), paragraph)) { Border = 0, HorizontalAlignment = Element.ALIGN_RIGHT });

                        pdfDoc.Add(headerTable);
                        pdfDoc.Add(spacer);


                        client.AddCell(new PdfPCell(new Phrase("Señor(a)", paragraph)) { Colspan = 2, Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        client.AddCell(new PdfPCell(new Phrase(data[i].name.ToString(), paragraph)) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        client.AddCell(new PdfPCell(new Phrase("Fecha de factura: " + data[i].invoice_date.Day + "/" + data[i].invoice_date.Month + "/" + data[i].invoice_date.Year, paragraph)) { Border = 0, HorizontalAlignment = Element.ALIGN_RIGHT });
                        client.AddCell(new PdfPCell(new Phrase(data[i].address.ToString(), paragraph)) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        client.AddCell(new PdfPCell(new Phrase("Fecha de vencimiento: " + data[i].expiry_date.Day + "/" + data[i].expiry_date.Month + "/" + data[i].expiry_date.Year, paragraph)) { Border = 0, HorizontalAlignment = Element.ALIGN_RIGHT });
                        client.AddCell(new PdfPCell(new Phrase(data[i].city.ToString(), paragraph)) { Border = 0, HorizontalAlignment = Element.ALIGN_LEFT });
                        client.AddCell(new PdfPCell(new Phrase("Número de contrato: " + data[i].id_contract.ToString(), paragraph)) { Border = 0, HorizontalAlignment = Element.ALIGN_RIGHT });

                        pdfDoc.Add(client);
                        pdfDoc.Add(spacer);

                        table.AddCell(new PdfPCell(new Phrase("RESUMEN", title)) { BackgroundColor = color, Colspan = 2, Padding = 6, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("Saldo anterior: " + data[i].previous_balance.ToString(), paragraph)) { Border = 1, Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("Total abonos: " + data[i].total_pp.ToString(), paragraph)) { Border = 1, Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("Saldo Actual: " + data[i].total_payable.ToString(), paragraph)) { Border = 0, Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("Total cargos: " + sum, paragraph)) { Border = 0, Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("CONCEPTO", title)) { BackgroundColor = color, Padding = 6, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("VALOR", title)) { BackgroundColor = color, Padding = 6, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_CENTER });
                        table.AddCell(new PdfPCell(new Phrase("Facturas Anteriores", paragraph)) { Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(data[i].previous_invoice.ToString(), paragraph)) { Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase("Intereses de mora (FA)", paragraph)) { Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(data[i].penalty_interest.ToString(), paragraph)) { Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase("Gastos de cobranza (FA)", paragraph)) { Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(data[i].collection_costs.ToString(), paragraph)) { Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase("Intereses de mora", paragraph)) { Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(data[i].actual_penalty_interest.ToString(), paragraph)) { Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase("Gastos de cobranza", paragraph)) { Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(data[i].actual_collection_costs.ToString(), paragraph)) { Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase("Subtotal factura del mes", paragraph)) { Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(data[i].subtotal.ToString(), paragraph)) { Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase("IVA factura del mes", paragraph)) { Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(data[i].iva.ToString(), paragraph)) { Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase("Total factura del mes", paragraph)) { Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });
                        table.AddCell(new PdfPCell(new Phrase(data[i].month_invoice.ToString(), paragraph)) { Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });

                        pdfDoc.Add(table);
                        pdfDoc.Add(spacer);

                        text.AddCell(new PdfPCell(new Phrase("Agradecemos su confianza en nosotros", paragraph)) { Border = 0, Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });
                        text.AddCell(new PdfPCell(new Phrase(data[i].random_text.ToString(), paragraph)) { Border = 0, Padding = 8, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_LEFT });
                        text.AddCell(new PdfPCell(new Phrase(data[i].barcode.ToString(), paragraph)) { Border = 0, PaddingTop = 45, VerticalAlignment = Element.ALIGN_MIDDLE, HorizontalAlignment = Element.ALIGN_RIGHT });
                        pdfDoc.Add(text);

                        //Se cierra el archivo PDF que se estaba llenando de datos.
                        pdfDoc.Close();
                    }
                    else
                    {
                        
                    }
                   
                }

                createExtractPrinttxt();
                createExtrcatEmailtxt();
                createWithheldExtracttxt();

                MessageBox.Show("Se crearon todos los documentos PDF y archivos planos, por favor revise en Disco E: ");
            }
            catch ( NullReferenceException e)
            {
                MessageBox.Show("Aún no hay información, por favor cargue los archivos " + e);
            }
        }

        //invocando el método para el envio de correo 
        private void sendEmail()
        {
            var sendemail = new sendExtractByEmail();
            var data = Dataextractcollections;
            
            sendemail.send(data);
        }

        //generar archivo plano con los extractos a imprimir
        private void createExtractPrinttxt()
        {

            var data = Dataextractcollections;

            using (StreamWriter outputfile = new StreamWriter("E:\\extractosparaimpresion.txt"))
            {
                
                for (int i = 0; i < data.Count; i++)
                {
                    if (data[i].output_type != "send by email" && data[i].generate_extract == true)
                    {
                        outputfile.WriteLine(i + "AAA|" + data[i].id_invoice + "|" + data[i].id_contract + "|" + data[i].invoice_date + "|" + data[i].expiry_date + "|" +
                        data[i].invoice_value + "|" + data[i].type_person + "|" + data[i].name + "|" + data[i].id_customer + "|" + data[i].address + "|"
                        + data[i].city + "|" + +data[i].subtotal + "|" + data[i].iva + "|" + data[i].total + "|" + data[i].random_text + "|"
                        + data[i].previous_invoice + "|" + data[i].penalty_interest + "|" + data[i].collection_costs + "|" + data[i].total_pp + "|"
                        + data[i].previous_balance + "|" + data[i].actual_penalty_interest + "|" + data[i].actual_collection_costs + "|" + data[i].month_invoice + "|"
                        + data[i].total_payable + "|" + data[i].barcode);
                    }
                       
                }
            }
        }

        //generar archivo plano con los extractos a enviar por correo electrónico
        private void createExtrcatEmailtxt()
        {

            var data = Dataextractcollections;

            using (StreamWriter outputfile = new StreamWriter("E:\\extractosparaenviarcorreos.txt"))
            {

                for (int i = 0; i < data.Count; i++)
                {
                    if (data[i].output_type == "send by email" && data[i].generate_extract == true)
                    {
                        outputfile.WriteLine(i + "AAA|" + data[i].id_invoice + "|" + data[i].id_contract + "|" + data[i].invoice_date + "|" + data[i].expiry_date + "|" +
                        data[i].invoice_value + "|" + data[i].type_person + "|" + data[i].name + "|" + data[i].id_customer + "|" + data[i].address + "|"
                        + data[i].city + "|" + +data[i].subtotal + "|" + data[i].iva + "|" + data[i].total + "|" + data[i].random_text + "|"
                        + data[i].previous_invoice + "|" + data[i].penalty_interest + "|" + data[i].collection_costs + "|" + data[i].total_pp + "|"
                        + data[i].previous_balance + "|" + data[i].actual_penalty_interest + "|" + data[i].actual_collection_costs + "|" + data[i].month_invoice + "|"
                        + data[i].total_payable + "|" + data[i].barcode);
                    }

                }
            }
        }

        //generar archivo plano con los extractos retenidos
        private void createWithheldExtracttxt()
        {

            var data = Dataextractcollections;

            using (StreamWriter outputfile = new StreamWriter("E:\\extractosretenidos.txt"))
            {

                for (int i = 0; i < data.Count; i++)
                {
                    if (data[i].generate_extract == false)
                    {
                        outputfile.WriteLine( i + "AAA|" +data[i].id_invoice + "|" + data[i].id_contract + "|" + data[i].invoice_date + "|" + data[i].expiry_date + "|" +
                        data[i].invoice_value + "|" + data[i].type_person + "|" + data[i].name + "|" + data[i].id_customer + "|" + data[i].address + "|"
                        + data[i].city + "|" + +data[i].subtotal + "|" + data[i].iva + "|" + data[i].total + "|" + data[i].random_text + "|"
                        + data[i].previous_invoice + "|" + data[i].penalty_interest + "|" + data[i].collection_costs + "|" + data[i].total_pp + "|"
                        + data[i].previous_balance + "|" + data[i].actual_penalty_interest + "|" + data[i].actual_collection_costs + "|" + data[i].month_invoice + "|"
                        + data[i].total_payable + "|" + data[i].barcode);
                    }

                }
            }
        }


    }

}
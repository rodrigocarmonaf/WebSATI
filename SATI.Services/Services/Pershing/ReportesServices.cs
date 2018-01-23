using SATI.Common.Common;
using System.Data;
using System.Configuration;
using System;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Text;
using System.Collections.Generic;
using System.Web;
using SATI.Common.Entities.Pershing;
using System.Linq;
using System.IO.Compression;


namespace SATI.Services.Services.Pershing
{
    public class ReportesServices
    {
        private PershingCommon _pershingCommon = new PershingCommon();

        #region LibroOperaciones
        public string GenerarPDFLibroOperaciones(string desde, string hasta)
        {
            DataTable dtLibroOperaciones = _pershingCommon.ObtenerLibroOperacionesPershing(desde, hasta);

            if (!Directory.Exists(Directory.GetCurrentDirectory() + ConfigurationManager.AppSettings["folder:pdf"].ToString() + "LibroOperaciones"))
            {
                Directory.CreateDirectory(Environment.CurrentDirectory + ConfigurationManager.AppSettings["folder:pdf"].ToString() + "LibroOperaciones");
            }

            string nameFile = "LibroOperaciones//LibroOperaciones_" + DateTime.Now.ToString().Replace("-", "").Replace(":", "").Replace("\\", "").Replace(" / ", "").Replace(" ", "") + ".pdf";
            string path = Environment.CurrentDirectory + ConfigurationManager.AppSettings["folder:pdf"].ToString() + nameFile;

            Document documento = new Document(PageSize.A0.Rotate());
            PdfWriter write = PdfWriter.GetInstance(documento,
                new FileStream(path, FileMode.Create));

            documento.AddTitle("Archivo Generado Automaticamente Por SATI");
            documento.AddCreator("Systema SATI");
            documento.Open();



            Font _TitleFont = new Font(Font.FontFamily.HELVETICA, 15, Font.BOLD, BaseColor.BLACK);
            Font _BodyFont = new Font(Font.FontFamily.HELVETICA, 12, Font.NORMAL, BaseColor.BLACK);
            Font _TitleTableFont = new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK);
            Font _TitleTblDetalle = new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, BaseColor.BLACK);

            documento.Add(new Paragraph(""));
            documento.Add(Chunk.NEWLINE);


            PdfPTable TblTitulo = new PdfPTable(1);
            TblTitulo.WidthPercentage = 90;

            PdfPCell clTitulo = new PdfPCell(new Phrase("\nLibro de Órdenes para ser ejecutadas en el Mercado Internacional, bajo la Circular N° 1.046, de la SVS.", _TitleFont));
            clTitulo.BorderWidth = 0;
            clTitulo.BorderColorBottom = BaseColor.BLACK;
            clTitulo.BorderWidthBottom = 1;
            clTitulo.HorizontalAlignment = Element.ALIGN_CENTER;
            clTitulo.BackgroundColor = new BaseColor(211, 211, 211);
            clTitulo.FixedHeight = 40;
            TblTitulo.AddCell(clTitulo);

            PdfPTable TblDetalle = new PdfPTable(3);
            TblDetalle.WidthPercentage = 90;
            TblDetalle.SetWidths(new float[] { 70f, 7.0f, 10f });
            PdfPCell clDetalle = new PdfPCell(new Phrase(""));
            clDetalle.Border = 0;
            TblDetalle.AddCell(clDetalle);
            clDetalle = new PdfPCell(new Phrase("Rango de Fecha : \n\nFecha Generación : ", _TitleTblDetalle));
            clDetalle.HorizontalAlignment = Element.ALIGN_CENTER;
            clDetalle.BackgroundColor = new BaseColor(211, 211, 211);
            clDetalle.Border = 0;
            TblDetalle.AddCell(clDetalle);
            clDetalle = new PdfPCell(new Phrase($"{DateTime.Parse(desde).ToShortDateString()} Al {DateTime.Parse(hasta).ToShortDateString()} \n\n{DateTime.Now.ToShortDateString()} ",_TitleTblDetalle));
            clDetalle.Border = 0;
            TblDetalle.AddCell(clDetalle);

            string[] TitulosCabezera = new string[]
            {
                "Fecha Operacion","Fecha Comision","Num Comision","Tipo Operacion",
                "Tipo Instrumento","Título","Intermediario Extranjero","Cantidad o Corte","Moneda",
                "Precio","Interes Devengado","Monto Bruto","Monto Comision","Monto Gastos",
                "Sec Fee","Comisión Tanner","Gastos Tanner","Monto Neto","Nombre Agente","Rut Cliente","Nombre Cliente"
                ,"Estado"
            };

            PdfPTable TblLibroOperacion = new PdfPTable(22)
            {
                WidthPercentage = 90
            };

            TblLibroOperacion.SetWidths(GetHeaderWidthsLibroOperaciones(_TitleTableFont, TitulosCabezera));

            for (int i = 0; i <= TitulosCabezera.Length - 1; i++)
            {
                PdfPCell Header = new PdfPCell(new Phrase("\n" + TitulosCabezera[i], _TitleTableFont));
                Header.Border = Rectangle.NO_BORDER;
                Header.FixedHeight = 30;
                Header.BackgroundColor = new BaseColor(211, 211, 211);

                TblLibroOperacion.AddCell(Header);
            }

            if (dtLibroOperaciones.Rows.Count > 0)
            {
                foreach (DataRow rows in dtLibroOperaciones.Rows)
                {
                    foreach (DataColumn colum in dtLibroOperaciones.Columns)
                    {
                        PdfPCell Body = new PdfPCell(new Phrase(rows[colum.ColumnName].ToString(), _BodyFont));

                        TblLibroOperacion.AddCell(Body);

                    }
                }
            }

            documento.Add(TblTitulo);
            documento.Add(TblDetalle);
            documento.Add(Chunk.NEWLINE);
            documento.Add(TblLibroOperacion);

            documento.Close();
            write.Close();

            AddPageNumberAndLogo(path);


            return string.Format("/SATI_PDF/{0}", nameFile);
        }

        public float[] GetHeaderWidthsLibroOperaciones(Font font, params string[] headers)
        {
            var result = new float[22];

            result[0] = 0.9f; //fecha operacion
            result[1] = 0.9f; //fecha Comision
            result[3] = 1.0f; //tipo operacion
            result[4] = 1.0f; //tipo instrumento
            result[5] = 0.7f; //titulo
            result[6] = 1.2f; //intermediario extranjero
            result[7] = 1.2f; //cantidad o corte
            result[8] = 0.5f;  //moneda
            result[9] = 1.3f;  //precio
            result[10] = 1.3f; //interes Devengado
            result[11] = 1.3f; //monto Bruto
            result[12] = 1.3f; //monto comision
            result[13] = 1.3f; //monto gastados
            result[14] = 0.5f; //sec fee
            result[15] = 1.3f; //comision tanner
            result[16] = 1.3f; //gastos tanner
            result[17] = 1.3f; //monto neto
            result[18] = 2.9f; //nombre agente
            result[19] = 1.0f; //rut cliente
            result[20] = 2.9f; //nombre cliente
            result[21] = 0.9f; //estado
            return result;
        }
        #endregion

        #region Registro Recepcion
        public string GenerarPDFRegistroRecepciones(string desde, string hasta)
        {
            DataTable dtLibroOperaciones = _pershingCommon.ObtenerRegistroRecepcionesPershing(desde, hasta);

            if (!Directory.Exists(Directory.GetCurrentDirectory() + ConfigurationManager.AppSettings["folder:pdf"].ToString() + "RegistroRecepcion"))
            {
                Directory.CreateDirectory(Environment.CurrentDirectory + ConfigurationManager.AppSettings["folder:pdf"].ToString() + "RegistroRecepcion");
            }

            string nameFile = "RegistroRecepcion//RegistroRecepcion_" + DateTime.Now.ToString().Replace("-", "").Replace(":", "").Replace("\\", "").Replace(" / ", "").Replace(" ", "") + ".pdf";
            string path = Environment.CurrentDirectory + ConfigurationManager.AppSettings["folder:pdf"].ToString() + nameFile;

            Document documento = new Document(PageSize.A0.Rotate());
            PdfWriter write = PdfWriter.GetInstance(documento,
                new FileStream(path, FileMode.Create));

            documento.AddTitle("Archivo Generado Automaticamente Por SATI");
            documento.AddCreator("Systema SATI");
            documento.Open();



            Font _TitleFont = new Font(Font.FontFamily.HELVETICA, 15, Font.BOLD, BaseColor.BLACK);
            Font _BodyFont = new Font(Font.FontFamily.HELVETICA, 12, Font.NORMAL, BaseColor.BLACK);
            Font _TitleTableFont = new Font(Font.FontFamily.HELVETICA, 12, Font.BOLD, BaseColor.BLACK);
            Font _TitleTblDetalle = new Font(Font.FontFamily.HELVETICA, 14, Font.BOLD, BaseColor.BLACK);

            documento.Add(new Paragraph(""));
            documento.Add(Chunk.NEWLINE);


            PdfPTable TblTitulo = new PdfPTable(1);
            TblTitulo.WidthPercentage = 90;

            PdfPCell clTitulo = new PdfPCell(new Phrase("\nRegistro de Comisiones Mercado Internacional, bajo la Cirular N° 1.046, de la SVS\n", _TitleFont));
            clTitulo.BorderWidth = 0;
            clTitulo.BorderColorBottom = BaseColor.BLACK;
            clTitulo.BorderWidthBottom = 1;
            clTitulo.HorizontalAlignment = Element.ALIGN_CENTER;
            clTitulo.BackgroundColor = new BaseColor(211, 211, 211);
            clTitulo.FixedHeight = 40;
            TblTitulo.AddCell(clTitulo);

            PdfPTable TblDetalle = new PdfPTable(3);
            TblDetalle.WidthPercentage = 90;
            TblDetalle.SetWidths(new float[] { 70f, 7.0f, 10f });
            PdfPCell clDetalle = new PdfPCell(new Phrase(""));
            clDetalle.Border = 0;
            TblDetalle.AddCell(clDetalle);
            clDetalle = new PdfPCell(new Phrase("Rango de Fecha : \n\nFecha Generación : ", _TitleTblDetalle));
            clDetalle.HorizontalAlignment = Element.ALIGN_CENTER;
            clDetalle.BackgroundColor = new BaseColor(211, 211, 211);
            clDetalle.Border = 0;
            TblDetalle.AddCell(clDetalle);
            clDetalle = new PdfPCell(new Phrase($"{DateTime.Parse(desde).ToShortDateString()} Al {DateTime.Parse(hasta).ToShortDateString()} \n\n{DateTime.Now.ToShortDateString()} ", _TitleTblDetalle));
            clDetalle.Border = 0;
            TblDetalle.AddCell(clDetalle);

            string[] TitulosCabezera = new string[]
            {
                "Rut Cliente","Nombre Cliente","Agente","Num. Comisión","Fecha Recepcion","Tipo Comision","Tipo Instrumento",
                "Titulo","Intermediario","Mercado","Cantidad o Corte","Moneda","Precio Limite","Fecha Limite","Estado"
            };

            PdfPTable TblLibroOperacion = new PdfPTable(15)
            {
                WidthPercentage = 90
            };

            TblLibroOperacion.SetWidths(GetHeaderWidthsRegistroRecepcion(_TitleTableFont, TitulosCabezera));

            for (int i = 0; i <= TitulosCabezera.Length - 1; i++)
            {
                PdfPCell Header = new PdfPCell(new Phrase("\n" + TitulosCabezera[i], _TitleTableFont));
                Header.Border = Rectangle.NO_BORDER;
                Header.FixedHeight = 30;
                Header.BackgroundColor = new BaseColor(211, 211, 211);

                TblLibroOperacion.AddCell(Header);
            }

            if (dtLibroOperaciones.Rows.Count > 0)
            {
                foreach (DataRow rows in dtLibroOperaciones.Rows)
                {
                    foreach (DataColumn colum in dtLibroOperaciones.Columns)
                    {
                        PdfPCell Body = new PdfPCell(new Phrase(rows[colum.ColumnName].ToString(), _BodyFont));

                        TblLibroOperacion.AddCell(Body);

                    }
                }
            }

            documento.Add(TblTitulo);
            documento.Add(TblDetalle);
            documento.Add(Chunk.NEWLINE);
            documento.Add(TblLibroOperacion);

            documento.Close();
            write.Close();

            AddPageNumberAndLogo(path);


            return string.Format("/SATI_PDF/{0}", nameFile);
        }

        public float[] GetHeaderWidthsRegistroRecepcion(Font font, params string[] headers)
        {
            var result = new float[15];

            result[0] = 0.5f; //Rut Cliente
            result[1] = 1.9f; //Nombre Cliente
            result[3] = 0.7f; //Numero Comision
            result[4] = 0.7f; //Fecha Recepcion
            result[5] = 0.7f; //titulo
            result[6] = 1.2f; //intermediario extranjero
            result[7] = 1.2f; //cantidad o corte
            result[8] = 0.5f;  //moneda
            result[9] = 1.3f;  //precio
            result[10] = 1.3f; //interes Devengado
            result[11] = 1.3f; //monto Bruto
            result[12] = 1.3f; //monto comision
            result[13] = 1.3f; //monto gastados
            result[14] = 0.5f; //sec fee

            return result;
        }
        #endregion

        #region Header y Paginacion

        protected void AddPageNumberAndLogo(string url)
        {
            byte[] bytes = File.ReadAllBytes(url);
            AcroFields pdfFormFields;
            string imagenTanner = "LogoTCB.PNG";
            string imagenPath = Path.Combine(@"img/", imagenTanner);

            Font blackFont = FontFactory.GetFont("Arial", 20.5f, Font.NORMAL, BaseColor.BLACK);
            using (MemoryStream stream = new MemoryStream())
            {
                PdfReader reader = new PdfReader(bytes);
                using (PdfStamper stamper = new PdfStamper(reader, stream))
                {
                    pdfFormFields = stamper.AcroFields;
                    int pages = reader.NumberOfPages;
                    for (int i = 1; i <= pages; i++)
                    {
                        ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_CENTER, new Phrase("Pagina " + i.ToString() + " de " + pages, blackFont), 230f, 15f, 0);

                        Image instanceImg = Image.GetInstance(imagenPath);
                        instanceImg.ScalePercent(12);
                        PdfContentByte overContent = stamper.GetOverContent(i);
                        instanceImg.SetAbsolutePosition(3160f, 2295f);
                        instanceImg.Alignment = Image.ALIGN_RIGHT;
                        overContent.AddImage(instanceImg);

                    }
                }
                bytes = stream.ToArray();
            }
            File.WriteAllBytes(url, bytes);
        }


        #endregion

        #region ContratosZip
        public string GenerarContratos(string fecha)
        {
            List<Contrato> ListadoOperaciones = _pershingCommon.ListadoOperacionesPendientes(DateTime.Parse(fecha));        
            List<Contrato> ListadoContratos = _pershingCommon.ListadoContratosPendientes(DateTime.Parse(fecha));
            List<ContratosCliente> ListadoFinalContratosOperaciones = new List<ContratosCliente>();
            
            string[] ListadoRutaContratos = new string[0];

            foreach (Contrato contrato in ListadoContratos)
            {
                List<Contrato> ListadoOperacionesDeContrato = new List<Contrato>();

                foreach (Contrato operaciones in ListadoOperaciones)
                {
                    if(operaciones.NumeroCuenta == contrato.NumeroCuenta)
                    {
                        ListadoOperacionesDeContrato.Add(operaciones);
                        _pershingCommon.ActualizarMarcaContrato(operaciones.Folio);
                    }
                }

                if(ListadoOperacionesDeContrato.Count > 0)
                {
                    ContratosCliente contratoCliente = new ContratosCliente();
                    contratoCliente.FechaProceso = DateTime.Parse(fecha);
                    contratoCliente.NombreCliente = contrato.NombreCliente;
                    contratoCliente.NumCuenta = contrato.NumeroCuenta;
                    contratoCliente.RutCliente = contrato.RutCliente;
                    contratoCliente.ListadoContratos = ListadoOperacionesDeContrato;
                    ListadoFinalContratosOperaciones.Add(contratoCliente);
                }                      
            }

            foreach(ContratosCliente gencontratos in ListadoFinalContratosOperaciones)
            {
               Array.Resize(ref ListadoRutaContratos, ListadoRutaContratos.Length + 1);
               ListadoRutaContratos[ListadoRutaContratos.Length - 1] = CreaContrato(gencontratos);
            }


           
            return GeneraZipContratos(ListadoRutaContratos);
        }

        private string GeneraZipContratos(string[] PathArchivos)
        {
            try
            {
                if (!Directory.Exists(Directory.GetCurrentDirectory() + ConfigurationManager.AppSettings["folder:pdf"].ToString() + "ZIP"))
                {
                    Directory.CreateDirectory(Environment.CurrentDirectory + ConfigurationManager.AppSettings["folder:pdf"].ToString() + "ZIP");
                }

                Random aleatorio = new Random();
                int numAleatorio = aleatorio.Next(99999);
                string directoryTemp = Environment.CurrentDirectory + "\\SATI_PDF\\ZIP\\";
                string directoryPDF = Environment.CurrentDirectory + "\\SATI_PDF\\";
                string nombreZip = DateTime.Now.ToString().Replace("-", "").Replace(":", "").Replace(" ", "").Trim() + ".zip";

                Directory.CreateDirectory(directoryTemp + numAleatorio);

                for (int i = 0; i <= PathArchivos.Length - 1; i++)
                {
                    File.Copy(directoryPDF + PathArchivos[i], directoryTemp + numAleatorio + "\\" + PathArchivos[i].Replace("Contratos//", ""));
                }

                ZipFile.CreateFromDirectory(directoryTemp + numAleatorio, directoryTemp + nombreZip);
                Directory.Delete(directoryTemp + numAleatorio, true);

                return string.Format("/SATI_PDF/ZIP/{0}", nombreZip);

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            return "";
        }

        public string CreaContrato(ContratosCliente contratosCliente)
        {
            try
            {


                if (!Directory.Exists(Directory.GetCurrentDirectory() + ConfigurationManager.AppSettings["folder:pdf"].ToString() + "Contratos"))
                {
                    Directory.CreateDirectory(Environment.CurrentDirectory + ConfigurationManager.AppSettings["folder:pdf"].ToString() + "Contratos");
                }

                string nameFile = "Contratos//ORDEN 1046_PERSHING_" + contratosCliente.NombreCliente + " " + DateTime.Now.ToString("ddMMyy_HHmmssfff") + ".pdf";
                string path = Environment.CurrentDirectory + ConfigurationManager.AppSettings["folder:pdf"].ToString() + nameFile;

                Document documento = new Document(PageSize.LETTER, 30, 30, 30, 30);
                PdfWriter writer = PdfWriter.GetInstance(documento, new FileStream(path, FileMode.Create));

                documento.Open();

                string imagenTanner = "LogoTCB.PNG";
                string imagenPath = Path.Combine(@"img/", imagenTanner);
                Image imgHeader = Image.GetInstance(imagenPath);
                imgHeader.ScalePercent(6);

                Paragraph sep = new Paragraph(" ");
                Font fntTitle = FontFactory.GetFont("Calibri", "utf-8", true, 10, Font.BOLD, BaseColor.BLACK);
                Paragraph para = new Paragraph("ORDEN DE COMPRAVENTA DE VALORES EXTRANJEROS EN MERCADOS EXTERNOS", fntTitle);
                para.Leading = 3;
                para.Alignment = 1;

                Font fntText = FontFactory.GetFont("Calibri", "utf-8", true, 7.5f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);
                Font fntTextN = FontFactory.GetFont("Calibri", "utf-8", true, 7.5f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);

                //Concatenar Nombre al mes
                Paragraph pharse = new Paragraph("Santiago, " + contratosCliente.FechaProceso.Day + " " + contratosCliente.FechaProceso.Month + " de " + contratosCliente.FechaProceso.Year, fntText);
                Paragraph chunk_ = new Paragraph("Señores", fntText);
                Paragraph chunk_cli = new Paragraph("Tanner Corredores de Bolsa S.A.", fntText);
                Paragraph chunk_a = new Paragraph("Presente", fntText);

                Paragraph para_a = new Paragraph("De mi consideración, ", fntText);
                Paragraph para_b = new Paragraph("Por medio de la presente, confirmo y ratifico que las operaciones más abajo detalladas, fueron todas y cada una de ellas instruidas por mí persona a Tanner Corredores de bolsa S.A. en adelante tambien Tanner para ser ejecutadas a través de Pershing LLC., en mi cuenta N° " + contratosCliente.NumCuenta, fntText);
                para_b.Alignment = Element.ALIGN_JUSTIFIED;

                Font fntTextA = FontFactory.GetFont("Calibri", "utf-8", true, 6.5f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);
                Font fntTxtFoot = FontFactory.GetFont("Calibri", "utf-8", true, 4f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);

                PdfPTable tablaOp = new PdfPTable(11);
                tablaOp.WidthPercentage = 100;
                tablaOp.SetWidths(new Single[] { 5.0f, 3.0f, 7.0F, 3.0f, 5.0f, 8.0f, 7.0f, 6.0f, 6.0f, 8.0f, 5.0f });

                PdfPCell cellA = new PdfPCell(new Phrase("Fecha Op.", fntTextN));
                cellA.HorizontalAlignment = 1;
                cellA.FixedHeight = 20;
                PdfPCell cellB = new PdfPCell(new Phrase("Folio", fntTextN));
                cellB.HorizontalAlignment = 1;
                PdfPCell cellBc = new PdfPCell(new Phrase("Emisor", fntTextN));
                cellBc.HorizontalAlignment = 1;
                PdfPCell cellC = new PdfPCell(new Phrase("MDA", fntTextN));
                cellC.HorizontalAlignment = 1;
                PdfPCell cellD = new PdfPCell(new Phrase("Tipo/Inst", fntTextN));
                cellD.HorizontalAlignment = 1;
                PdfPCell cellE = new PdfPCell(new Phrase("ISIN", fntTextN));
                cellE.HorizontalAlignment = 1;
                PdfPCell cellF = new PdfPCell(new Phrase("Nominales", fntTextN));
                cellF.HorizontalAlignment = 1;
                PdfPCell cellG = new PdfPCell(new Phrase("Precio", fntTextN));
                cellG.HorizontalAlignment = 1;
                PdfPCell cellH = new PdfPCell(new Phrase("Tipo Ope.", fntTextN));
                cellH.HorizontalAlignment = 1;
                PdfPCell cellI = new PdfPCell(new Phrase("Remuneración", fntTextN));
                cellI.HorizontalAlignment = 1;
                PdfPCell cellJ = new PdfPCell(new Phrase("Mercado", fntTextN));
                cellJ.HorizontalAlignment = 1;

                tablaOp.AddCell(cellA);
                tablaOp.AddCell(cellB);
                tablaOp.AddCell(cellBc);
                tablaOp.AddCell(cellC);
                tablaOp.AddCell(cellD);
                tablaOp.AddCell(cellE);
                tablaOp.AddCell(cellF);
                tablaOp.AddCell(cellG);
                tablaOp.AddCell(cellH);
                tablaOp.AddCell(cellI);
                tablaOp.AddCell(cellJ);

                foreach (Contrato contratos in contratosCliente.ListadoContratos)
                {
                    //Fecha Op.
                    PdfPCell cellCA = new PdfPCell(new Phrase(contratosCliente.FechaProceso.ToShortDateString(), fntTextA));
                    cellCA.FixedHeight = 25;
                    cellCA.HorizontalAlignment = Element.ALIGN_CENTER;
                    tablaOp.AddCell(cellCA);

                    //Folio
                    PdfPCell cellCB = new PdfPCell(new Phrase(contratos.Folio.ToString(), fntTextA));
                    cellCB.HorizontalAlignment = Element.ALIGN_CENTER;
                    tablaOp.AddCell(cellCB);

                    //Emisor
                    PdfPCell cellCC = new PdfPCell(new Phrase(contratos.Simbol, fntTextA));
                    cellCC.HorizontalAlignment = Element.ALIGN_CENTER;
                    tablaOp.AddCell(cellCC);

                    //Moneda
                    PdfPCell cellCD = new PdfPCell(new Phrase(contratos.Currency, fntTextA));
                    cellCD.HorizontalAlignment = Element.ALIGN_CENTER;
                    tablaOp.AddCell(cellCD);

                    //Tipo/Inst
                    PdfPCell cellCE;
                    if (contratos.Currency.Length <= 4)
                    {
                        cellCE = new PdfPCell(new Phrase("Acción", fntTextA));
                    }
                    else
                    {
                        cellCE = new PdfPCell(new Phrase("Bono", fntTextA));
                    }

                    cellCE.HorizontalAlignment = Element.ALIGN_CENTER;
                    tablaOp.AddCell(cellCE);

                    //Isin
                    PdfPCell cellCF = new PdfPCell(new Phrase(contratos.Isin, fntTextA));
                    cellCF.HorizontalAlignment = Element.ALIGN_CENTER;
                    tablaOp.AddCell(cellCF);

                    //Nominales
                    PdfPCell cellCG = new PdfPCell(new Phrase(contratos.Nominales.ToString(), fntTextA));
                    cellCG.HorizontalAlignment = Element.ALIGN_CENTER;
                    tablaOp.AddCell(cellCG);

                    //Precio
                    PdfPCell cellCH = new PdfPCell(new Phrase(contratos.Precio.ToString(), fntTextA));
                    cellCH.HorizontalAlignment = Element.ALIGN_CENTER;
                    tablaOp.AddCell(cellCH);

                    //Tipo Ope.
                    PdfPCell cellCI = new PdfPCell(new Phrase(contratos.TipoOperacion, fntTextA));
                    cellCI.HorizontalAlignment = Element.ALIGN_CENTER;
                    tablaOp.AddCell(cellCI);

                    //Remuneración
                    PdfPCell cellCJ = new PdfPCell(new Phrase(contratos.Comision.ToString(), fntTextA));
                    cellCJ.HorizontalAlignment = Element.ALIGN_CENTER;
                    tablaOp.AddCell(cellCJ);

                    //Mercado
                    PdfPCell cellCK = new PdfPCell(new Phrase("EEUU", fntTextA));
                    cellCK.HorizontalAlignment = Element.ALIGN_CENTER;
                    tablaOp.AddCell(cellCK);
                }

                documento.Add(imgHeader);
                documento.Add(sep);
                documento.Add(sep);
                documento.Add(para);
                documento.Add(sep);
                documento.Add(pharse);
                documento.Add(sep);
                documento.Add(chunk_);
                documento.Add(chunk_cli);
                documento.Add(chunk_a);
                documento.Add(sep);
                documento.Add(para_a);
                documento.Add(sep);
                documento.Add(para_b);
                documento.Add(sep);
                documento.Add(sep);
                documento.Add(tablaOp);
                documento.Add(sep);
                documento.Add(sep);
                documento.Add(sep);
                documento.Add(sep);

                Paragraph countpage = new Paragraph();
                countpage.Alignment = Element.ALIGN_RIGHT;
                countpage.Add(new Chunk("Página 1 de 3", fntTxtFoot));

                PdfPTable footTbl = new PdfPTable(1);
                footTbl.TotalWidth = 300;
                footTbl.HorizontalAlignment = Element.ALIGN_CENTER;

                PdfPCell cellFoot = new PdfPCell(countpage);
                cellFoot.Border = 0;
                cellFoot.Padding = 5;

                footTbl.AddCell(cellFoot);
                footTbl.WriteSelectedRows(0, -1, 515, 30, writer.DirectContent);

                documento.NewPage();

                documento.Add(imgHeader);

                Font fntTextB = FontFactory.GetFont("Calibri", "utf-8", true, 7.5f, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                Font fntTextR = FontFactory.GetFont("Calibri", "utf-8", true, 7.5f, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);

                Paragraph para_c = new Paragraph();
                para_c.Alignment = Element.ALIGN_JUSTIFIED;
                para_c.Add(new Chunk("a) Entrega, recepcion de valores y liquidación: ", fntTextB));
                para_c.Add(new Chunk("Los procedimientos para efectuar la entrega, recepción y liquidación de la(s) orden(es) constan en el contrato de comisión para la compra y venta de valores extranjeros en mercados externos celebrado entre el Cliente y Tanner Corredores de Bolsa S.A., el cual está vigente.", fntTextR));

                Paragraph para_d = new Paragraph();
                para_d.Alignment = Element.ALIGN_JUSTIFIED;
                para_d.Add(new Chunk("b) Custodia: ", fntTextB));
                para_d.Add(new Chunk("Solicito que los instrumentos adquiridos a través de esta(s) orden(es), sean ingresador a mi nombre en mi cuenta que mantengo en Pershing; quien ejercerá el servicio de custodia del mismo.", fntTextR));

                Paragraph para_e = new Paragraph();
                para_e.Alignment = Element.ALIGN_JUSTIFIED;
                para_e.Add(new Chunk("c) Egreso de Custodia: ", fntTextB));
                para_e.Add(new Chunk("Solicito que (el) los intrumento(s) vendido(s) a través de estas comisiones, sean egresados desde mi cuenta que mantengo en Pershing.", fntTextR));

                Paragraph para_f = new Paragraph();
                para_f.Alignment = Element.ALIGN_JUSTIFIED;
                para_f.Add(new Chunk("d) Autorización: ", fntTextB));
                para_f.Add(new Chunk("Faculto expresamente a Tanner para actual a nombre propio en el cumplimiento de esta(s) orden(es).", fntTextR));

                Paragraph para_g = new Paragraph();
                para_g.Alignment = Element.ALIGN_JUSTIFIED;
                para_g.Add(new Chunk("e) Responsabilidades: ", fntTextB));
                para_g.Add(new Chunk("Declaro estar en conocimiento y aceptar que: (i) Tanner Corredores de Bolsa S.A. no asume responsabilidad alguna por la solvencia de los emisores de los valores que adquiera en virtud de estas instrucciones, o por la rentabilidad de los mismos; (ii) Tanner Corredores de Bolsa S.A. no se hace responsable por las variaciones en el tipo de cambio que pudieran producirse; (iii) Tanner Corredores de Bolsa S.A. no se hace responsable de retornar las divisas si por disposiciones de derecho interno del país donde se realiza la orden, se restrinja o impida el acceso a la compra o remesa de divisas; (iv) Tanner Corredores de Bolsa S.A. no se hace responsable de las variaciones impositivas a empresas y/o inversionistas en los mercados donde se realizan las órdenes; (v) Tanner Corredores de Bolsa S.A. se declara responsable por los dineros recibidos y por la ejecución de la orden de acuerdo a las instrucciones del Cliente; (vi) El Cliente será el único responsable de cumplir con las normas de cambios internacionales y demás disposiciones establecidas por el Banco Central de Chile, no asumiendo Tanner Corredores de Bolsa S.A., o el intermediario extranjero, obligación alguna a este respecto; (vii) El Cliente será el responsable de informar las operaciones, inversiones y resultados de éstas al Servicio de Impuestos Internos, conforme a las disposiciones que al efecto estén vigentes; (viii) Tanner no cobrará comisión directa por la ejecución de esta(s) orden(es), sin perjuicio que puede compartir los ingresos que obtengan los intermediarios o emisores extranjeros con ocasión del cumplimiento de la misma, ya sea producto de comisiones, de diferencias de precio de los valores que el Cliente ordene adquirir o enajenar o por otros conceptos; (ix) Pershing efectuó los cobros indicados por concepto de remuneración y de derechos, los que se cargaron en el monto neto de la(s) orden(es). En virtud de lo anterior, otorgo a Tanner un completo, total e irrevocable finiquito con relación a la(s) mencionada(s) operación(es)efectuada(s) sobre mi cuenta en Pershing LLC, declarando que, en su calidad de comisionista, Tanner Corredores de Bolsa S.A. ha dado cabal y oportuno cumplimiento a todas las obligaciones derivadas en la ejecución de dicha(s) orden(es).", fntTextR));

                Paragraph para_h = new Paragraph();
                para_h.Alignment = Element.ALIGN_JUSTIFIED;
                para_h.Add(new Chunk("f) ", fntTextB));
                para_h.Add(new Chunk(" El Cliente declara conocer y aceptar los riesgos asociados las operaciones de compra y venta que realiza a través de Tanner, por lo que acepta los eventuales riesgos que se puedan producir respecto al resultado de las inversiones producto de esta orden, entre los que se encuentran que el emisor de los valores entre en cesación de pagos o quiebra, o por las fluctuaciones de mercado, o por la mayor o menor rentabilidad que puedan experimentar estas inversiones.", fntTextR));

                Paragraph para_i = new Paragraph();
                para_i.Alignment = Element.ALIGN_JUSTIFIED;
                para_i.Add(new Chunk("g) ", fntTextB));
                para_i.Add(new Chunk("  Esta orden se realiza en virtud de la Circular No. 1.046 de la Superintendencia de Valores y Seguros, en donde se autoriza a los Corredores de Bolsa, prestar a sus clientes el servicio de asesorías, estudios, y comisión de compra y venta de valores extranjeros en mercados de valores externos, a través de intermediarios de valores autorizados. En dicha se especifica el procedimiento y los requisitos que deben ser observados por los clientes para instruir las comisiones de compra y venta, y los requisitos que deben ser cumplidos por los corredores nacionales, para ejecutar las comisiones conferidas. ", fntTextR));

                Paragraph para_j = new Paragraph("Sin otro particular, saluda atentamente a Usted", fntTextR);
                para_j.Alignment = Element.ALIGN_JUSTIFIED;

                documento.Add(sep);
                documento.Add(para_c);
                documento.Add(para_d);
                documento.Add(para_e);
                documento.Add(para_f);
                documento.Add(para_g);
                documento.Add(para_h);
                documento.Add(sep);

                countpage = new Paragraph();
                countpage.Alignment = Element.ALIGN_RIGHT;
                countpage.Add(new Chunk("Página 2 de 3", fntTxtFoot));

                footTbl = new PdfPTable(1);
                footTbl.TotalWidth = 300;
                footTbl.HorizontalAlignment = Element.ALIGN_CENTER;

                cellFoot = new PdfPCell(countpage);
                cellFoot.Border = 0;
                cellFoot.Padding = 10;

                footTbl.AddCell(cellFoot);
                footTbl.WriteSelectedRows(0, -1, 515, 30, writer.DirectContent);

                documento.NewPage();
                documento.Add(imgHeader);
                documento.Add(para_i);
                documento.Add(sep);
                documento.Add(para_j);
                documento.Add(sep);
                documento.Add(sep);
                documento.Add(sep);
                documento.Add(sep);

                PdfPTable tablaFirma = new PdfPTable(2);
                tablaFirma.WidthPercentage = 90;

                PdfPCell cellFirmaImg = new PdfPCell();
                PdfPCell cellFirmaDummy = new PdfPCell();
                cellFirmaDummy.Border = 0;
                cellFirmaImg.Border = 0;
                cellFirmaImg.HorizontalAlignment = PdfPCell.LEFT_BORDER;


                string imagenFirma = "firma_rm.PNG";
                string imagenPathFirma = Path.Combine(@"img/", imagenFirma);
                Image imgFirmaR = Image.GetInstance(imagenPathFirma);

                imgFirmaR.ScalePercent(50);
                imgFirmaR.Alignment = Image.ALIGN_LEFT;

                cellFirmaImg.AddElement(imgFirmaR);
                tablaFirma.AddCell(cellFirmaImg);
                tablaFirma.AddCell(cellFirmaDummy);

                PdfPCell cellFirmaT = new PdfPCell();
                cellFirmaT.Border = 0;

                cellFirmaT.AddElement(new Paragraph("Tanner Corredores de Bolsa S.A.", fntTextB));
                cellFirmaT.AddElement(new Paragraph("RUT N° 80.962.600-8", fntText));
                cellFirmaT.AddElement(new Paragraph("Domicilio: Apoquindo 3650 Of. 902", fntText));
                cellFirmaT.AddElement(new Paragraph("p.p. Renato Madrid", fntText));
                cellFirmaT.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;

                tablaFirma.AddCell(cellFirmaT);

                PdfPCell cellFirmaC = new PdfPCell();
                cellFirmaC.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                cellFirmaC.Border = 0;

                cellFirmaC.AddElement(new Paragraph("Cliente: " + contratosCliente.NombreCliente, fntText));
                cellFirmaC.AddElement(new Paragraph("Rut: " + contratosCliente.RutCliente, fntText));

                DataTable dtDomicilio = new DataTable();
                string Domicilio, Ciudad;
                //dtDomicilio = clnte.cl_cn_direccion_cliente(txtRutCliente.Text.Substring(0, Len(txtRutCliente.Text) - 2))

                if (dtDomicilio.Rows.Count > 0)
                {
                    dtDomicilio.Rows[0][0] = dtDomicilio.Rows[0][0].ToString().Trim().Replace("¥", "Ñ");
                    dtDomicilio.Rows[0][1] = dtDomicilio.Rows[0][1].ToString().Trim().Replace("San Fabi n", "San Fabián");
                    dtDomicilio.Rows[0][1] = dtDomicilio.Rows[0][1].ToString().Trim().Replace("San Nicol s", "San Nicolás");
                    dtDomicilio.Rows[0][1] = dtDomicilio.Rows[0][1].ToString().Trim().Replace("R nquil", "Ránquil");
                    dtDomicilio.Rows[0][1] = dtDomicilio.Rows[0][1].ToString().Trim().Replace("Chill n Viejo", "Chillán Viejo");
                    dtDomicilio.Rows[0][1] = dtDomicilio.Rows[0][1].ToString().Trim().Replace("Chill n", "Chillán");
                    dtDomicilio.Rows[0][1] = dtDomicilio.Rows[0][1].ToString().Trim().Replace("Santa B rbara", "Santa Bárbara");
                    dtDomicilio.Rows[0][1] = dtDomicilio.Rows[0][1].ToString().Trim().Replace("Juan Fern ndez", "Juan Fernández");
                    dtDomicilio.Rows[0][1] = dtDomicilio.Rows[0][1].ToString().Trim().Replace("Santa B rbara", "Santa Bárbara");

                    Domicilio = dtDomicilio.Rows[0][0].ToString().Trim() + ", " + dtDomicilio.Rows[0][1].ToString().Trim();

                    //Chaflán para solucionar el problema de las "a" con acento.
                    dtDomicilio.Rows[0][2] = dtDomicilio.Rows[0][2].ToString().Trim().Replace("San Fabi n", "San Fabián");
                    dtDomicilio.Rows[0][2] = dtDomicilio.Rows[0][2].ToString().Trim().Replace("San Nicol s", "San Nicolás");
                    dtDomicilio.Rows[0][2] = dtDomicilio.Rows[0][2].ToString().Trim().Replace("R nquil", "Ránquil");
                    dtDomicilio.Rows[0][2] = dtDomicilio.Rows[0][2].ToString().Trim().Replace("Chill n Viejo", "Chillán Viejo");
                    dtDomicilio.Rows[0][2] = dtDomicilio.Rows[0][2].ToString().Trim().Replace("Chill n", "Chillán");
                    dtDomicilio.Rows[0][2] = dtDomicilio.Rows[0][2].ToString().Trim().Replace("Santa B rbara", "Santa Bárbara");
                    dtDomicilio.Rows[0][2] = dtDomicilio.Rows[0][2].ToString().Trim().Replace("Juan Fern ndez", "Juan Fernández");
                    dtDomicilio.Rows[0][2] = dtDomicilio.Rows[0][2].ToString().Trim().Replace("Santa B rbara", "Santa Bárbara");

                    Ciudad = dtDomicilio.Rows[0][2].ToString().Trim();
                }
                else
                {
                    Domicilio = "";
                    Ciudad = "";
                }

                cellFirmaC.AddElement(new Paragraph("Domicilio:." + Domicilio, fntText));
                cellFirmaC.AddElement(new Paragraph(Ciudad, fntText));

                DataTable RepLegal = new DataTable(); /*clnte.cl_cn_rep_legal_cliente(txtRutCliente.Text.Substring(0, Len(txtRutCliente.Text) - 2))*/

                foreach (DataRow dtRep in RepLegal.Rows)
                {
                    cellFirmaC.AddElement(new Paragraph("Apoderado: " + dtRep[1].ToString().Trim(), fntText));
                    cellFirmaC.AddElement(new Paragraph("Rut:" + dtRep[0].ToString().Trim(), fntText));
                }

                tablaFirma.AddCell(cellFirmaC);
                documento.Add(tablaFirma);

                documento.Add(sep);
                documento.Add(sep);
                documento.Add(sep);
                documento.Add(sep);
                documento.Add(sep);
                documento.Add(sep);

                countpage = new Paragraph();
                countpage.Alignment = Element.ALIGN_RIGHT;
                countpage.Add(new Chunk("Página 3 de 3", fntTxtFoot));

                footTbl = new PdfPTable(1);
                footTbl.TotalWidth = 300;
                footTbl.HorizontalAlignment = Element.ALIGN_CENTER;

                cellFoot = new PdfPCell(countpage);
                cellFoot.Border = 0;
                cellFoot.Padding = 10;

                footTbl.AddCell(cellFoot);
                footTbl.WriteSelectedRows(0, -1, 515, 30, writer.DirectContent);

                documento.Close();

                return nameFile;
            }
            catch(Exception e)
            {
                Console.Write(e.Message);
            }

            return null;
        }
        #endregion

        #region Utilidades Contratos
        private List<Contrato> ObtenerContratosSeleccionados(List<Contrato> listadoContratos, string[] contratosId)
        {
            List<Contrato> ListadoContratosSeleccionados = new List<Contrato>();
            foreach (Contrato contrato in listadoContratos)
            {
                if (contratosId.Where(p => p == contrato.Id.ToString()).ToList().Count > 0)
                {
                    ListadoContratosSeleccionados.Add(contrato);
                }
            }

            return ListadoContratosSeleccionados;
        }
        #endregion
    }
}
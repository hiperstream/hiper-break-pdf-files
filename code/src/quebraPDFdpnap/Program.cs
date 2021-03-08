using System;
using System.Collections.Generic;
using System.IO;
using System.Data;

using Microsoft.VisualBasic;
using clsDPNAPTemplate;
using iText.Kernel.Pdf;
using System.Linq.Expressions;
using System.Runtime.CompilerServices;
using System.Linq;

namespace quebraPDFdpnap
    {
    public class Program
        {
        static string arqECM = string.Empty;
        static string dirOUT = string.Empty;
        static string dirBKP = string.Empty;
        static string TXTbkp = string.Empty;
        static string TXTgat = string.Empty;

        //static string dirLOG = string.Empty;
        static string arqPDF = string.Empty;
        public static clsDPNAPTP DPNAP { get; set; }
        static void Main(string[] args)
            {
            Run();
            }

        private static void Run()
            {

            DPNAP = new clsDPNAPTP("Use apenas no DPNAP (v0.1)...");
            arqECM = RetornarEntrada("ENTRADA");
            arqPDF = RetornarEntrada("ENTRADAPDF");
            dirOUT = DPNAP.PastaSaida;
            dirBKP = DPNAP.PastaRecuros; // para onde vamos mover o PDF

            //DPNAP.GravarnoLog(arqECM);
            //DPNAP.GravarnoLog(arqPDF);

            try
                {
                DataTable ecmData = GetDataTableFromCSVFile(arqECM);
                // Console.WriteLine("Rows count: " + ecmData.Rows.Count);
                //*
                DPNAP.GravarnoLog("Quebrando PDF");
                DPNAP.GravarnoLog("Arquivo ECM...: " + arqECM);
                DPNAP.GravarnoLog("Arquivo PDF...: " + arqPDF);
                DPNAP.GravarnoLog("Pasta de saída: " + dirOUT);
                //*/
                // abre o pdfIn

                if (File.Exists(arqPDF))
                    {
                    DPNAP.GravarnoLog("\n\nAbrindo " + arqPDF + "\n");
                    }
                else
                    {
                    DPNAP.GravarnoLog("\n\nProblemas abrindo " + arqPDF + "\n");
                    }

                PdfDocument pdfSRC = new PdfDocument(new PdfReader(arqPDF));

                DPNAP.GravarnoLog("O arquivo " + arqPDF + " possui " + pdfSRC.GetNumberOfPages() + " páginas.");

                TXTbkp = Path.Combine(dirOUT, "BKP", Path.GetFileNameWithoutExtension(arqECM) + ".TXT");
                TXTgat = Path.Combine(dirOUT, Path.GetFileNameWithoutExtension(arqECM) + ".TXT");

                dirOUT = Path.Combine(dirOUT, Path.GetFileNameWithoutExtension(arqECM));


                if (!Directory.Exists(dirOUT))
                    Directory.CreateDirectory(dirOUT);

                foreach (DataRow element in ecmData.Rows)
                    {
                    var arqECMsai = Path.Combine(dirOUT, Path.GetFileName(element.Field<string>("arqECM")) + "_noPass.pdf"); // nome do pdf a ser gravado
                    var arqECMfim = Path.Combine(dirOUT, Path.GetFileName(element.Field<string>("arqECM")) + ".pdf"); // nome do pdf a ser gravado
                    var pagInicial = int.Parse(element.Field<string>("pagInicial"));
                    var pagFatura = int.Parse(element.Field<string>("pagFatura"));
                    var pdfSenhaS = element.Field<string>("pdfSenha");
                    //var pdfSenha = System.Text.Encoding.UTF8.GetBytes(element.Field<string>("pdfSenha"));

                    PdfDocument pdfOUT = new PdfDocument(new PdfWriter(arqECMsai));
                    pdfOUT.InitializeOutlines();

                    //PdfDocument pdfPASS
                    IList<int> pages = new List<int>();

                    for (int i = 0; i < pagFatura; i++)
                        {
                        pages.Add(pagInicial + i);
                        }
                    _ = pdfSRC.CopyPagesTo(pages, pdfOUT);

                    pdfOUT.SetCloseWriter(true);
                    pdfOUT.Close();

                    if (pdfSenhaS.Length > 0 && pdfSenhaS != "NOPASSWORD" )
                        {
                        //DPNAP.GravarnoLog($"Tem senha:");
                        //DPNAP.GravarnoLog($"  Lendo _noPass: {arqECMsai}");
                        //DPNAP.GravarnoLog($"  Criando senha: {arqECMfim}");
                        PdfReader reader = new PdfReader(arqECMsai);
                        WriterProperties props = new WriterProperties().SetStandardEncryption(
                            System.Text.Encoding.UTF8.GetBytes(pdfSenhaS), 
                            System.Text.Encoding.UTF8.GetBytes(pdfSenhaS), 
                            EncryptionConstants.ALLOW_PRINTING,
                            encryptionAlgorithm: EncryptionConstants.ENCRYPTION_AES_256 | EncryptionConstants.DO_NOT_ENCRYPT_METADATA);
                        PdfWriter writer = new PdfWriter(new PdfWriter(arqECMfim), props);
                        var pdfDoc = new PdfDocument(reader, writer);
                        pdfDoc.Close();
                        reader.Close();
                        writer.Close();
                        File.Delete(arqECMsai);
                        }
                    else
                        {
                        //DPNAP.GravarnoLog($"Sem senha:");
                        //DPNAP.GravarnoLog($"  Lendo _noPass: {arqECMsai}");
                        //DPNAP.GravarnoLog($"  Criando senha: {arqECMfim}");
                        File.Move(arqECMsai, arqECMfim);
                        }
                    
                    }
                pdfSRC.Close();
                DPNAP.codSaida = 0;
                DPNAP.GravarnoLog("Movendo:");
                DPNAP.GravarnoLog("DE:   " + TXTbkp);
                DPNAP.GravarnoLog("PARA: " + TXTgat);

                if (File.Exists(TXTbkp))
                    {
                    File.Move(TXTbkp, TXTgat);
                    }
                else
                    {
                    DPNAP.GravarnoLog("Não encontrei: " + TXTbkp);
                    }


                DPNAP.GravarnoLog("Terminado com sucesso!!!");
                Environment.Exit(0);
                }
            catch (Exception ex)
                {
                DPNAP.GravarnoLog(ex.Message);
                Environment.Exit(3);
                }
            }

        private static DataTable GetDataTableFromCSVFile(string csv_file_path)
            {
            DataTable csvData = new DataTable();

            try
                {
                Microsoft.VisualBasic.FileIO.TextFieldParser tfp = new Microsoft.VisualBasic.FileIO.TextFieldParser(csv_file_path);
                var csvReader = tfp;
                csvReader.SetDelimiters(new string[] { ";" });
                csvReader.HasFieldsEnclosedInQuotes = false;
                string[] colFields = csvReader.ReadFields();
                foreach (string column in colFields)
                    {
                    DataColumn datecolumn = new DataColumn(column);
                    datecolumn.AllowDBNull = true;
                    csvData.Columns.Add(datecolumn);
                    }
                while (!csvReader.EndOfData)
                    {
                    string[] fieldData = csvReader.ReadFields();
                    // making empty value as null
                    for (int i = 0; i < fieldData.Length; i++)
                        {
                        if (fieldData[i] == "")
                            {
                            fieldData[i] = null;
                            }
                        }
                    csvData.Rows.Add(fieldData);
                    }
                }
            catch (Exception ex)
                {
                Console.WriteLine(ex.Message);
                }
            return csvData;
            }
        private static string RetornarEntrada(string NomeEntrada)
            {
            string entrada = null;
            entrada = DPNAP.Entradas.ToList().Find(e => e.NomeEntrada.ToLower() == NomeEntrada.ToLower()).ArquivoEntrada.ToString();
            if (entrada == null) DPNAP.GravarnoLog("Nome da entrada do dpnap não encontrado");
            return entrada;
            }
        }

    }

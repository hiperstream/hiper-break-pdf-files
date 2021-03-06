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
        public static clsDPNAPTP DPNAP { get; set; }
        static void Main()
            {
            Run();
            }

        private static void Run()
            {
            string TXTbkp;
            string TXTgat;
            bool SENHA;
            string arqPDF = string.Empty;

            DPNAP = new clsDPNAPTP("Use apenas no DPNAP (v0.1)...");
            var arqECM = RetornarEntrada("ENTRADA");

            if (DPNAP.Entradas.Count() > 1 && DPNAP.PastaRecuros.Trim() == "")
            {
                arqPDF = RetornarEntrada("ENTRADAPDF");
            }
            else if (DPNAP.Entradas.Count() == 1 && DPNAP.PastaRecuros.Trim() != "")
            {
                arqPDF = DPNAP.PastaRecuros + Path.GetFileNameWithoutExtension(arqECM) + "." + DPNAP.Engine;
            }
            else
            {
                DPNAP.GravarnoLog("Se existe mais de uma entrada o caminho de recursos deve estar vazio!!");
                Environment.Exit(3);
            }

            string dirOUT = DPNAP.PastaSaida;
            string[] splitJob = DPNAP.Job.Split(':');
            SENHA = (splitJob[0] == "SENHA" && !(Path.GetFileNameWithoutExtension(arqECM).Contains("$OCAL")));

            try
                {
                DataTable ecmData = GetDataTableFromCSVFile(arqECM);
                DPNAP.GravarnoLog("Quebrando PDF");
                DPNAP.GravarnoLog("Arquivo ECM...: " + arqECM);
                DPNAP.GravarnoLog("Arquivo PDF...: " + arqPDF);
                DPNAP.GravarnoLog("Pasta de saída: " + dirOUT);
                if (SENHA)
                    {
                    DPNAP.GravarnoLog(">>> Aplicando senha nos PDF individuais");
                    }
                else
                    {
                    DPNAP.GravarnoLog(">>> Arquivos PDF individuais sem senha");
                    }
                    
                if (File.Exists(arqPDF))
                    {
                    DPNAP.GravarnoLog("\n\nAbrindo " + arqPDF + "\n");
                    }
                else
                    {
                    DPNAP.GravarnoLog("\n\nNão encontrei " + arqPDF + "\n");
                    throw new FileNotFoundException("This file was not found.");
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
                    var arqECMsai = Path.Combine(dirOUT, Path.GetFileName(element.Field<string>("arqECM")) + "_noPass.pdf"); 
                    var arqECMfim = Path.Combine(dirOUT, Path.GetFileName(element.Field<string>("arqECM")) + ".pdf");
                    var pagInicial = int.Parse(element.Field<string>("pagInicial"));
                    var pagFatura = int.Parse(element.Field<string>("pagFatura"));
                    string pdfSenhaS = null;
                    if (SENHA) 
                        {
                        pdfSenhaS = element.Field<string>(splitJob[1]);
                        }

                    PdfDocument pdfOUT = new PdfDocument(new PdfWriter(arqECMsai));
                    pdfOUT.InitializeOutlines();

                    IList<int> pages = new List<int>();

                    for (int i = 0; i < pagFatura; i++)
                        {
                        pages.Add(pagInicial + i);
                        }
                    _ = pdfSRC.CopyPagesTo(pages, pdfOUT);

                    pdfOUT.SetCloseWriter(true);
                    pdfOUT.Close();

                    if (SENHA)
                        {
                        PdfComSenha(arqECMsai, arqECMfim, pdfSenhaS);
                        }
                    else
                        {
                        File.Move(arqECMsai, arqECMfim);
                        }
                    }
                pdfSRC.Close();
                DPNAP.codSaida = 0;
                DPNAP.GravarnoLog("Movendo:");
                DPNAP.GravarnoLog("DE.....: " + TXTbkp);
                DPNAP.GravarnoLog("PARA...: " + TXTgat);

                if (File.Exists(TXTbkp))
                {
                    File.Move(TXTbkp, TXTgat);
                }
                else
                {
                    DPNAP.GravarnoLog("Não encontrei: " + TXTbkp);
                }

                if (DPNAP.PastaRecuros != null)
                {
                    string bkpPDF = Path.Combine(Path.GetDirectoryName(arqPDF), "BKP", Path.GetFileName(arqPDF));
                    File.Move(arqPDF, bkpPDF);
                }
                else
                {
                    DPNAP.GravarnoLog("Não foi necessario mover o PDF!: " + Path.GetFileName(arqPDF));
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
                    DataColumn datacolumn = new DataColumn(column)
                        {
                        AllowDBNull = true
                        };
                    csvData.Columns.Add(column: datacolumn);
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
            if (entrada == null) DPNAP.GravarnoLog("Nome da entrada do dpnap não encontrado!");
            return entrada;
            }

        public static void PdfComSenha(string arqECMsai, string arqECMfim, string pdfSenhaS)
            {
            PdfReader reader = new PdfReader(arqECMsai);
            WriterProperties props = new WriterProperties().SetStandardEncryption(
                System.Text.Encoding.UTF8.GetBytes(pdfSenhaS),
                System.Text.Encoding.UTF8.GetBytes(pdfSenhaS),
                EncryptionConstants.ALLOW_PRINTING,
                encryptionAlgorithm: EncryptionConstants.ENCRYPTION_AES_256 | EncryptionConstants.DO_NOT_ENCRYPT_METADATA);
            PdfWriter writer = new PdfWriter(new PdfWriter(arqECMfim), props);
            var pdfDoc = new PdfDocument(reader, writer);
            pdfDoc.Close();
            writer.Close();
            reader.Close();
            File.Delete(arqECMsai);
            }
        }
    }

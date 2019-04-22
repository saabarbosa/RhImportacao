using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using Excel;

namespace RhImportacao
{
    class Program
    {
        static void Main(string[] args)
        {
            //string op = "-i"; //args[0]
            //string file = @"C:\temp\TemplateFuncionario.xls"; //args[1]
            ////string file = @"C:\temp\TemplateHeadCount.xls"; //args[1]
            //string user = "admin";  //args[2]
            //string pass = "GB@rb0sa";  //args[3]


            string op = "";
            string file = "";
            string user = "";
            string pass = "";
            try
            {
                op = args[0];
                file = @args[1];
                user = args[2];
                pass = args[3];
            }
            catch (Exception)
            {
                Console.WriteLine("RhImportacao -i|+i \"PathFullFile\"  Username  Password\n");
                Console.WriteLine("-i ou +i       = limpar e incrementa ou apenas incrementa");
                Console.WriteLine("\"PathFullFile\" = caminho completo da planilha excel, ex: c:\\temp\\arquivo.xls");
                Console.WriteLine("Username       = usuario administrador da rotina de importacao.");
                Console.WriteLine("Password       = senha do administrador da rotina de importacao.");
                return;
            }

            try
            {
                if (File.Exists(file))
                {

                    string nomeArquivo = Path.GetFileNameWithoutExtension(file);
                    string extensao = Path.GetExtension(file);

                    FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read);
                    IExcelDataReader excelReader;

                    if (extensao.ToLower().Equals(".xlsx"))
                        excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    else
                        excelReader = ExcelReaderFactory.CreateBinaryReader(stream);


                    DataSet ds = excelReader.AsDataSet();
                    if (ds != null && ds.Tables != null && ds.Tables.Count > 0)
                    {
                        DataTable dte = ds.Tables[0];

                        if ( (user.Equals("admin")) && (pass.Equals("GB@rb0sa")))
                        {
                            string result = "";
                            DateTime inicio = DateTime.Now;
                            string mensagem = "aguarde, processando (essa operacao pode levar alguns minutos)...";
                            switch (nomeArquivo.ToLower())
                            {
                                case "templateheadcount":
                                    Console.WriteLine(mensagem);
                                    result = popularDadosHeadCount(op, dte);
                                    break;
                                case "templatefuncionario":
                                    Console.WriteLine(mensagem);
                                    result = popularDadosFuncionario(op, dte);
                                    break;
                                case "templatebloquearrp":
                                    Console.WriteLine(mensagem);
                                    result = popularDadosBloquearRP(op, dte);
                                    break;
                                default:
                                    result = "Nome de arquivo invalido.";
                                    break;

                            }
                            DateTime fim = DateTime.Now;
                            TimeSpan ts = fim - inicio;
                            Console.WriteLine((!String.IsNullOrEmpty(result))? "Dado(s) atualizado(s) com sucesso. " + result + " reg(s) em " + Math.Round(ts.TotalSeconds).ToString() + " segundo(s)." : "Erro ao tentar importar a planilha. Contate o adm do sistema.{"+ result +"}");
                        }
                        else
                        {
                            Console.WriteLine("Usuario e/ou senha invalido.");
                        }


                    }
                    stream.Close();

                }
                else
                {
                    Console.WriteLine("Caminho de arquivo inexistente.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                
            }
            //Console.ReadKey();

        }

        static private string popularDadosHeadCount(string op, DataTable dte)
        {

            StringBuilder sb = new StringBuilder();
            int? retorno = 0;
            if (dte.Rows[0][0].ToString().ToUpper().Equals("BANDEIRA") && dte.Rows[0][1].ToString().ToUpper().Equals("ESTABELECIMENTO")
                && dte.Rows[0][2].ToString().ToUpper().Equals("CODCARGO") && dte.Rows[0][3].ToString().ToUpper().Equals("QE")
                && dte.Rows[0][4].ToString().ToUpper().Equals("QA") && dte.Rows[0][5].ToString().ToUpper().Equals("VL")
                && dte.Rows[0][6].ToString().ToUpper().Equals("VD") && dte.Rows[0][7].ToString().ToUpper().Equals("VS"))
            {
                if (op.ToLower().Equals("-i"))
                    Database.Delete("aux_head_count");

                int posLinha = 1;
                //Percorre o excel para inserir os dados
                for (; posLinha < dte.Rows.Count; posLinha++)
                {
                    DataRow row = dte.Rows[posLinha];

                    if (!string.IsNullOrEmpty(row[1].ToString()))
                    {
                        string ban = row[0].ToString();
                        string estabelecimento = row[1].ToString();
                        string codcargo = row[2].ToString();
                        int qe = int.Parse(row[3].ToString());
                        int qa = int.Parse(row[4].ToString());
                        int vl = int.Parse(row[5].ToString());
                        int vd = int.Parse(row[6].ToString());
                        int vs = int.Parse(row[7].ToString());

                        // Insere a informação na tabela aux_funcionario
                        sb.Append("insert into aux_head_count ");
                        sb.Append("(BAN, ESTABELECIMENTO, CODCARGO, QE, QA, VL, VD, VS, dtHeadCount) ");
                        sb.Append("values('").Append(ban).Append("', '").Append(estabelecimento).Append("', '").Append(codcargo).Append("', ").Append(qe).Append(", ");
                        sb.Append(qa).Append(", ").Append(vl).Append(", ").Append(vd).Append(", ").Append(vs).Append(", ").Append("getDate()").Append(");");

                        //retorno = Database.Inserir(ban, estabelecimento, codcargo, qe, qa, vl, vd, vs);
                    }

                }
                retorno = Database.Inserir(sb);
                return "{" + retorno.ToString() + "} " + (posLinha-1).ToString();
            }
            else
            {
                return "O arquivo excel selecionado não está no modelo permitido. Favor selecionar outro arquivo! (Revisar Template de Importação)";
            }
        }

        static private string popularDadosFuncionario(string op, DataTable dte)
        {
            int? retorno = 0;
            StringBuilder sb = new StringBuilder();
            if (dte.Rows[0][0].ToString().ToUpper().Equals("BANDEIRA") && dte.Rows[0][1].ToString().ToUpper().Equals("ESTABELECIMENTO")
                && dte.Rows[0][2].ToString().ToUpper().Equals("MATRICULA") && dte.Rows[0][3].ToString().ToUpper().Equals("NOME")
                && dte.Rows[0][4].ToString().ToUpper().Equals("CODCARGO") && dte.Rows[0][5].ToString().ToUpper().Equals("CARGO")
                && dte.Rows[0][6].ToString().ToUpper().Equals("UNIDADE") && dte.Rows[0][7].ToString().ToUpper().Equals("UNIDADEDESCRICAO")
                && dte.Rows[0][8].ToString().ToUpper().Equals("CENTROCUSTO") && dte.Rows[0][9].ToString().ToUpper().Equals("CENTROCUSTODESCRICAO")
                && dte.Rows[0][10].ToString().ToUpper().Equals("JORNADA") && dte.Rows[0][11].ToString().ToUpper().Equals("STATUS")
                && dte.Rows[0][12].ToString().ToUpper().Equals("QUEBRA") && dte.Rows[0][13].ToString().ToUpper().Equals("CONT")
                && dte.Rows[0][14].ToString().ToUpper().Equals("EMPRESA"))
            {

                if (op.ToLower().Equals("-i"))
                    Database.Delete("aux_funcionario");

                int posLinha = 1;
                //Percorre o excel para inserir os dados
                for ( ; posLinha < dte.Rows.Count; posLinha++)
                {
                    DataRow row = dte.Rows[posLinha];

                    if (!string.IsNullOrEmpty(row[1].ToString()))
                    {
                        string ban = row[0].ToString();
                        string estabelecimento = row[1].ToString();
                        string matricula = row[2].ToString();
                        string nome = row[3].ToString();
                        string codcargo = row[4].ToString();
                        string cargo = row[5].ToString();
                        string unidade = row[6].ToString();
                        string unidadeDescricao = row[7].ToString();
                        string centroCusto = row[8].ToString();
                        string centroCustoDescricao = row[9].ToString();
                        string jornada = row[10].ToString();
                        string status = row[11].ToString();
                        string quebra = row[12].ToString();
                        if (quebra.Equals(""))
                            quebra = "NULL";
                        int cont = (!row[13].ToString().Equals("")) ? int.Parse(row[13].ToString()) : 0;

                        string empresa = row[14].ToString();
                        // Insere a informação na tabela aux_funcionario

                        sb.Append("insert into aux_funcionario ");
                        sb.Append("(BAN, ESTABELECIMENTO, MATRICULA, NOME, CODCARGO, CARGO, UNIDADE, UNIDADEDESCRICAO, CENTROCUSTO, CENTROCUSTODESCRICAO, JORNADA, STATUS, QUEBRA, CONT, EMPRESA) ");
                        sb.Append("values('").Append(ban).Append("', '").Append(estabelecimento).Append("', '").Append(matricula).Append("', '").Append(nome).Append("', '");
                        sb.Append(codcargo).Append("', '").Append(cargo).Append("', '").Append(unidade).Append("', '").Append(unidadeDescricao).Append("', '").Append(centroCusto).Append("', '");
                        sb.Append(centroCustoDescricao).Append("', '").Append(jornada).Append("', '").Append(status).Append("', '").Append(quebra).Append("', ").Append(cont).Append(", '").Append(empresa).Append("');");
  
                        //retorno = Database.Inserir(ban, estabelecimento, matricula, nome, codcargo, cargo, unidade, unidadeDescricao, centroCusto, centroCustoDescricao, jornada, status, quebra, cont, empresa);

                    }

                }
                retorno = Database.Inserir(sb);
                return "{" + retorno.ToString() + "} " + (posLinha - 1).ToString();
            }
            else
            {
                return "O arquivo excel selecionado não está no modelo permitido. Favor selecionar outro arquivo! (Revisar Template de Importação)";
            }
        }

        static private string popularDadosBloquearRP(string op, DataTable dte)
        {
            return "";
            /*
            if (dte.Rows[0][0].ToString().ToUpper().Equals("BANDEIRA") && dte.Rows[0][1].ToString().ToUpper().Equals("ESTABELECIMENTO")
            && dte.Rows[0][2].ToString().ToUpper().Equals("CODCARGO") && dte.Rows[0][3].ToString().ToUpper().Equals("QE")
            && dte.Rows[0][4].ToString().ToUpper().Equals("QA") && dte.Rows[0][5].ToString().ToUpper().Equals("VL")
            && dte.Rows[0][6].ToString().ToUpper().Equals("VD") && dte.Rows[0][7].ToString().ToUpper().Equals("VS"))
            {

                //Percorre o excel para inserir os dados
                for (int posLinha = 1; posLinha < dte.Rows.Count; posLinha++)
                {
                    DataRow row = dte.Rows[posLinha];

                    if (!string.IsNullOrEmpty(row[1].ToString()))
                    {
                        string ban = row[0].ToString();
                        string estabelecimento = row[1].ToString();
                        string codcargo = row[2].ToString();
                        string qe = row[3].ToString();
                        string qa = row[4].ToString();
                        string vl = row[5].ToString();
                        string vd = row[6].ToString();
                        string vs = row[7].ToString();

                        // Insere a informação na tabela aux_funcionario
                        int? retorno = Database.Inserir(ban, estabelecimento, codcargo, qe, qa, vl, vd, vs);
                    }

                }
                return "Dados inseridos com sucesso.";
            }
            else
            {
                return "O arquivo excel selecionado não está no modelo permitido. Favor selecionar outro arquivo! (Revisar Template de Importação)";
            }
            */
        }

    }


}

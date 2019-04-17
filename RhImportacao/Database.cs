
using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace RhImportacao
{
    class Database
    {

        public static int? Inserir(string ban, string estabelecimento, string codCargo, int qe, int qa, int vl, int vd, int vs)
        {

            int? nrLinhas;
            SqlConnection con = new SqlConnection();
            con.ConnectionString = Util.CONNECTION_STRING;

            try
            {
                SqlCommand comm = new SqlCommand();
                comm.Connection = con;
                comm.CommandText = @"
                insert into aux_head_count
                (BAN, ESTABELECIMENTO, CODCARGO, QE, QA, VL, VD, VS, dtHeadCount) 
                values (@ban, @estabelecimento, @codCargo, @qe, @qa, @vl, @vd, @vs, getDate())";

                con.Open();

                comm.Parameters.Add(new SqlParameter("BAN", ban));
                comm.Parameters.Add(new SqlParameter("ESTABELECIMENTO", estabelecimento));
                comm.Parameters.Add(new SqlParameter("CODCARGO", codCargo));
                comm.Parameters.Add(new SqlParameter("QE", qe));
                comm.Parameters.Add(new SqlParameter("QA", qa));
                comm.Parameters.Add(new SqlParameter("VL", vl));
                comm.Parameters.Add(new SqlParameter("VD", vd));
                comm.Parameters.Add(new SqlParameter("VS", vs));

                nrLinhas = comm.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                nrLinhas = null;
            }
            finally
            {
                con.Close();
            }
            return nrLinhas;
        }

        public static int? Inserir(string ban, string estabelecimento, string matricula, string nome, string codCargo, string cargo, string unidade, string unidadeDescricao, string centroCusto, string centroCustoDescricao, string jornada, string status, string quebra, int cont, string empresa)
        {

            int? nrLinhas;
            SqlConnection con = new SqlConnection();
            con.ConnectionString = Util.CONNECTION_STRING;

            try
            {
                SqlCommand comm = new SqlCommand();
                comm.Connection = con;
                comm.CommandText = @"
                insert into aux_funcionario
                (BAN ,ESTABELECIMENTO ,MATRICULA ,NOME ,CODCARGO ,CARGO ,UNIDADE ,UNIDADEDESCRICAO ,CENTROCUSTO ,CENTROCUSTODESCRICAO ,JORNADA ,STATUS ,QUEBRA ,CONT ,EMPRESA) 
                values (@ban,@estabelecimento, @matricula, @nome, @codCargo, @cargo, @unidade, @unidadeDescricao, @centroCusto, @centroCustoDescricao, @jornada, @status, @quebra, @cont, @empresa)";
 
                con.Open();

                comm.Parameters.Add(new SqlParameter("BAN", ban));
                comm.Parameters.Add(new SqlParameter("ESTABELECIMENTO", estabelecimento));
                comm.Parameters.Add(new SqlParameter("MATRICULA", matricula));
                comm.Parameters.Add(new SqlParameter("NOME", nome));
                comm.Parameters.Add(new SqlParameter("CODCARGO", codCargo));
                comm.Parameters.Add(new SqlParameter("CARGO", cargo));
                comm.Parameters.Add(new SqlParameter("UNIDADE", unidade));
                comm.Parameters.Add(new SqlParameter("UNIDADEDESCRICAO", unidadeDescricao));
                comm.Parameters.Add(new SqlParameter("CENTROCUSTO", centroCusto));
                comm.Parameters.Add(new SqlParameter("CENTROCUSTODESCRICAO", centroCustoDescricao));
                comm.Parameters.Add(new SqlParameter("JORNADA", jornada));
                comm.Parameters.Add(new SqlParameter("STATUS", status));
                comm.Parameters.Add(new SqlParameter("QUEBRA", quebra));
                comm.Parameters.Add(new SqlParameter("CONT", cont));
                comm.Parameters.Add(new SqlParameter("EMPRESA", empresa));

                
                nrLinhas = comm.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                nrLinhas = null;
            }
            finally
            {
                con.Close();
            }
            return nrLinhas;


        }

        public static int? Inserir(StringBuilder sqlInsert)
        {

            int? nrLinhas;
            SqlConnection con = new SqlConnection();
            con.ConnectionString = Util.CONNECTION_STRING;

            try
            {
                SqlCommand comm = new SqlCommand();
                comm.Connection = con;
                comm.CommandText = sqlInsert.ToString();
                con.Open();

                nrLinhas = comm.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                nrLinhas = null;
            }
            finally
            {
                con.Close();
            }
            return nrLinhas;


        }

        public static int? Delete(string table)
        {
            int? nrLinhas;
            SqlConnection con = new SqlConnection();
            con.ConnectionString = Util.CONNECTION_STRING;

            try
            {
                SqlCommand comm = new SqlCommand();
                comm.Connection = con;
                comm.CommandText = @"
                delete from " + table;

                con.Open();

                nrLinhas = comm.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                nrLinhas = null;
            }
            finally
            {
                con.Close();
            }
            return nrLinhas;
        }

    }
}

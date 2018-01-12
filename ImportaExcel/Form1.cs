using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ImportaExcel
{
    public partial class Form1 : Form
    {
        const string STR_CONN = @"Password=s2u0r1f0;Persist Security Info=True;User ID=sa;Initial Catalog=db_eneagrama;Data Source=VARGOLA-PC\SQLEXPRESS14";

        private Excel.Application xlApp = null;
        private Excel.Workbooks xlWorkbooks = null;
        private Excel.Workbook xlWorkbook = null;
        private Excel.Sheets xlWorksheets = null;
        private Excel._Worksheet xlWorksheet = null;
        private Excel.Range xlRange = null;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            txtPlanilha.Text = "";

            openFileDialog1.Title = "Selecionar planilha Excel";
            //openFileDialog1.InitialDirectory = @"h:\backup\wrk\eneagrama\planilhas";
            openFileDialog1.Filter = "Excel (*.xls;*.xlsx)|*.xls;*.xlsx";
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;

            DialogResult dr = openFileDialog1.ShowDialog();

            if(dr == System.Windows.Forms.DialogResult.OK)
            {
                txtPlanilha.Text = openFileDialog1.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (txtPlanilha.Text.Trim() == String.Empty)
            {
                MessageBox.Show("Informe a planilha");
                txtPlanilha.Focus();
                return;
            }

            try
            {
                xlApp = new Excel.Application();
                xlWorkbooks = xlApp.Workbooks;
                xlWorkbook = xlWorkbooks.Open(txtPlanilha.Text.Trim(),
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                xlWorksheets = xlWorkbook.Sheets;
                xlWorksheet = null;
                xlRange = null;

                for (int plan = 1; plan <= xlWorksheets.Count; plan++)
                {
                    xlWorksheet = xlWorksheets[plan];
                    xlRange = xlWorksheet.UsedRange;

                    int rowTurma = 1;

                    string ds_turma = getCell(ref xlRange, rowTurma, 1).ToUpper();
                    string id_curso = getCell(ref xlRange, rowTurma, 2);
                    string nr_ano = getCell(ref xlRange, rowTurma, 3);
                    string nr_mes = getCell(ref xlRange, rowTurma, 4);
                    string id_professor = getCell(ref xlRange, rowTurma, 5);
                    string ds_local = getCell(ref xlRange, rowTurma, 6).ToUpper();

                    string sql = @"insert into tb_turmas (ds_turma, id_curso, nr_ano, nr_mes, id_professor, ds_local)
                                   values (@ds_turma, @id_curso, @nr_ano, @nr_mes, @id_professor, @ds_local)
                                   select scope_identity()";
                    List<SqlParameter> param = new List<SqlParameter>();
                    param.Add(new SqlParameter("@ds_turma", ds_turma));
                    param.Add(new SqlParameter("@id_curso", id_curso));
                    param.Add(new SqlParameter("@nr_ano", nr_ano));
                    param.Add(new SqlParameter("@nr_mes", nr_mes));
                    param.Add(new SqlParameter("@id_professor", id_professor));
                    param.Add(new SqlParameter("@ds_local", ds_local));
                    int id_turma = comandoSQL(sql, param);
                    param.Clear();

                    int colNome = getCellToInt32(ref xlRange, rowTurma, 7);

                    int[] colEmail = new int[3];
                    colEmail[0] = getCellToInt32(ref xlRange, rowTurma, 8);
                    colEmail[1] = getCellToInt32(ref xlRange, rowTurma, 9);
                    colEmail[2] = getCellToInt32(ref xlRange, rowTurma, 10);

                    int[] colFone = new int[3];
                    colFone[0] = getCellToInt32(ref xlRange, rowTurma, 11);
                    colFone[1] = getCellToInt32(ref xlRange, rowTurma, 12);
                    colFone[2] = getCellToInt32(ref xlRange, rowTurma, 13);

                    int rowIni = 3;

                    for (int row = rowIni; row <= xlRange.Rows.Count; row++)
                    {
                        string key = "t" + id_turma + "-" + (row - 2);

                        if (colNome > 0)
                        {
                            string nome = getCell(ref xlRange, row, colNome).ToUpper();
                            sql = @"insert into tb_alunos (nm_aluno, id_key_planilha)
                                    values (@nm_aluno, @id_key_planilha)
                                    select scope_identity()";
                            param.Add(new SqlParameter("@nm_aluno", nome));
                            param.Add(new SqlParameter("@id_key_planilha", key));
                            int id_aluno = comandoSQL(sql, param);
                            param.Clear();

                            short id_email = 0;
                            for (int i = 0; i < 3; i++)
                            {
                                if (colEmail[i] > 0)
                                {
                                    string email = getCell(ref xlRange, row, colEmail[i]).ToLower();
                                    if (email.Contains("@"))
                                    {
                                        id_email += 1;
                                        sql = @"insert into tb_alunos_emails (id_aluno, id_email, ds_email)
                                                values (@id_aluno, @id_email, @ds_email)";
                                        param.Add(new SqlParameter("@id_aluno", id_aluno));
                                        param.Add(new SqlParameter("@id_email", id_email));
                                        param.Add(new SqlParameter("@ds_email", email));
                                        comandoSQL(sql, param);
                                        param.Clear();
                                    }
                                }
                            }

                            int id_telefone = 0;
                            for (int i = 0; i < 3; i++)
                            {
                                if (colFone[i] > 0)
                                {
                                    string telefone = getCell(ref xlRange, row, colFone[i]).Replace(" ", "")
                                        .Replace("(", "").Replace(")", "")
                                        .Replace("-", "").Replace(".", "").Trim();
                                    if (!telefone.Equals(String.Empty))
                                    {
                                        short nr_ddd = 0;
                                        int nr_telefone = 0;
                                        short id_tipo_telefone = 1;
                                        if (telefone.Length <= 8)
                                        {
                                            nr_ddd = 11;
                                            nr_telefone = Convert.ToInt32(telefone);
                                            id_tipo_telefone = 1;
                                        }
                                        else if (telefone.Length == 9)
                                        {
                                            nr_ddd = 11;
                                            nr_telefone = Convert.ToInt32(telefone);
                                            id_tipo_telefone = 2;
                                        }
                                        else if (telefone.Length == 10)
                                        {
                                            nr_ddd = Convert.ToInt16(telefone.Substring(0, 2));
                                            nr_telefone = Convert.ToInt32(telefone.Substring(2));
                                            id_tipo_telefone = 1;
                                        }
                                        else
                                        {
                                            nr_ddd = Convert.ToInt16(telefone.Substring(0, 2));
                                            nr_telefone = Convert.ToInt32(telefone.Substring(2));
                                            id_tipo_telefone = (short)(telefone.Substring(2, 1).Equals("9") ? 2 : 1);
                                        }
                                        id_telefone += 1;
                                        sql = @"insert into tb_alunos_telefones (id_aluno, id_telefone, nr_ddd, nr_telefone, id_tipo_telefone)
                                                values (@id_aluno, @id_telefone, @nr_ddd, @nr_telefone, @id_tipo_telefone)";
                                        param.Add(new SqlParameter("@id_aluno", id_aluno));
                                        param.Add(new SqlParameter("@id_telefone", id_telefone));
                                        param.Add(new SqlParameter("@nr_ddd", nr_ddd));
                                        param.Add(new SqlParameter("@nr_telefone", nr_telefone));
                                        param.Add(new SqlParameter("@id_tipo_telefone", id_tipo_telefone));
                                        comandoSQL(sql, param);
                                        param.Clear();
                                    }
                                }
                            }
                        }
                    }

                    releaseComObject(xlRange);
                    xlRange = null;
                    releaseComObject(xlWorksheet);
                    xlWorksheet = null;
                }

                MessageBox.Show("Processo finalizado com sucesso!");

            }
            catch (Exception ex)
            {
                MessageBox.Show("ERRO: " + ex.Message);
            }
            finally
            {
                xlWorkbook.Close();
                liberarProcessoExcel();
            }
        }

        private string getCell(ref Excel.Range xlRange, int row, int col)
        {
            var cell = xlRange.Cells[row, col];
            if (cell != null && cell.Value2 != null)
            {
                return cell.Value2.ToString().Trim();
            }
            return "";
        }

        private int getCellToInt32(ref Excel.Range xlRange, int row, int col)
        {
            string cell = getCell(ref xlRange, row, col);
            return cell.Equals(String.Empty) ? 0 : Convert.ToInt32(cell);
        }

        private void releaseComObject(object o)
        {
            if (o != null)
            {
                Marshal.ReleaseComObject(o);
                o = null;
            }
        }

        private void liberarProcessoExcel()
        {
            releaseComObject(xlRange);
            xlRange = null;
            releaseComObject(xlWorksheet);
            xlWorksheet = null;
            releaseComObject(xlWorksheets);
            xlWorksheets = null;
            releaseComObject(xlWorkbook);
            xlWorkbook = null;
            releaseComObject(xlWorkbooks);
            xlWorkbooks = null;
            if (xlApp != null)
            {
                xlApp.Quit();
                releaseComObject(xlApp);
            }
            xlApp = null;
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();
        }

        private int comandoSQL(string sql, List<SqlParameter> param)
        {
            SqlConnection conn = new SqlConnection(STR_CONN);
            conn.Open();
            SqlCommand comm = conn.CreateCommand();
            comm.CommandText = sql;
            comm.CommandType = CommandType.Text;
            foreach(var p in param)
            {
                comm.Parameters.Add(p);
            }
            int result = 0;
            if (sql.ToLower().Contains("select scope_identity()"))
            {
                result = Convert.ToInt32(comm.ExecuteScalar());
            }
            else
            {
                result = comm.ExecuteNonQuery();
            }
            comm.Dispose();
            conn.Close();
            conn.Dispose();
            return result;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            liberarProcessoExcel();
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
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
            openFileDialog1.InitialDirectory = @"h:\backup\wrk\eneagrama\planilhas";
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

            xlApp = new Excel.Application();
            xlWorkbooks = xlApp.Workbooks;
            xlWorkbook = xlWorkbooks.Open(txtPlanilha.Text.Trim(),
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            xlWorksheets = xlWorkbook.Sheets;
            xlWorksheet = null;
            xlRange = null;

            int rowIni = 2;

            for(int i = rowIni; i <= xlWorksheets.Count; i++)
            {
                xlWorksheet = xlWorksheets[i];
                xlRange = xlWorksheet.UsedRange;

                int rowTurma = 1;
                string ds_turma = getCell(xlRange, rowTurma, 1);
                string id_curso = getCell(xlRange, rowTurma, 2);
                string nr_ano = getCell(xlRange, rowTurma, 3);
                string nr_mes = getCell(xlRange, rowTurma, 4);
                string id_professor = getCell(xlRange, rowTurma, 5);
                string ds_local = getCell(xlRange, rowTurma, 6);

                string sql = @"insert into tb_turmas (ds_turma, id_curso, nr_ano, nr_mes, id_professor, ds_local)
                        values (@ds_turma, @id_curso, @nr_ano, @nr_mes, @id_professor, @ds_local)";
                List<SqlParameter> param = new List<SqlParameter>();
                param.Add(new SqlParameter("@ds_turma", ds_turma));
                param.Add(new SqlParameter("@id_curso", id_curso));
                param.Add(new SqlParameter("@nr_ano", nr_ano));
                param.Add(new SqlParameter("@nr_mes", nr_mes));
                param.Add(new SqlParameter("@id_professor", id_professor));
                param.Add(new SqlParameter("@ds_local", ds_local));
                int id_turma = comandoSQL(sql, param);

                int colNome = getCellToInt32(xlRange, rowTurma, 7);
                int colEmail1 = getCellToInt32(xlRange, rowTurma, 8);
                int colEmail2 = getCellToInt32(xlRange, rowTurma, 9);
                int colEmail3 = getCellToInt32(xlRange, rowTurma, 10);
                int colFone1 = getCellToInt32(xlRange, rowTurma, 11);
                int colFone2 = getCellToInt32(xlRange, rowTurma, 12);
                int colFone3 = getCellToInt32(xlRange, rowTurma, 13);

                int c = 0;
                for(int row = rowIni; row <= xlRange.Rows.Count; row++)
                {
                    c++;
                    string key = "t" + id_turma + "-" + c;

                    if (colNome > 0)
                    {
                        string nome = getCell(xlRange, row, colNome);

                        sql = "insert into tb_alunos (nm_aluno, id_key_planilha) values ( " +
                            "'" + nome.Trim().ToUpper() + "', " +
                            key + " )";
                        string idAluno = "7622";

                    }

                    for(col = 1; col <= xlRange.Columns.Count; col++)
                    {
                        var cell = xlRange.Cells[row, col];
                        if (cell != null && cell.Value2 != null)
                        {
                            var linha = "[" + row + ", " + col + "] = " + cell.Value2;
                            richTextBox1.AppendText("\n" + linha);
                        }
                    }
                }

                releaseComObject(xlRange);
                xlRange = null;
                releaseComObject(xlWorksheet);
                xlWorksheet = null;
            }

            xlWorkbook.Close();
            liberarProcessoExcel();
            
        }

        private string getCell(Excel.Range xlRange, int row, int col)
        {
            var cell = xlRange.Cells[row, col];
            if (cell != null && cell.Value2 != null)
            {
                return cell.ToString().Trim();
            }
            return "";
        }

        private int getCellToInt32(Excel.Range xlRange, int row, int col)
        {
            string cell = getCell(xlRange, row, col);
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
            xlApp.Quit();
            releaseComObject(xlApp);
            xlApp = null;
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
            int result = comm.ExecuteNonQuery();
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

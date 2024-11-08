using Npgsql;
using OfficeOpenXml;
using System.Data;

namespace ExcelToBD
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private List<string> ReadExcel(string filename)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var lista = new List<string>();
            using (var pckg = new ExcelPackage(filename))
            {
                var worksheet = pckg.Workbook.Worksheets[0];//primera hoja

                var colcount = worksheet.Dimension.End.Row; //La cantidad de filas
                for (int i = 2; i < colcount; i++) //Iniciamos de la fila 2, porque la 1 es el encabezado en este caso
                {
                    string row = "A" + i;//Leemos la fila A, pero tambien podríamos indicarle otra, incluso varias
                    var cellvalue = worksheet.Cells[row].Text; //Obtenemos el texto
                    if (cellvalue != null && cellvalue != "")
                    {
                        lista.Add(cellvalue);
                    }
                }
                return lista;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == DialogResult.OK)
            {
                string fileExt = Path.GetExtension(file.FileName);
                if (fileExt == ".xlsx")
                {
                    try
                    {
                        var dtexcel = ReadExcel(file.FileName);
                        InsertInBD(dtexcel);
                        MessageBox.Show("Se guardó correctamente en BD: ", "", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ocurrió la siguiente excepción: "+ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        throw;
                    }
                }
                else
                {
                    MessageBox.Show("El archivo ingresado no es un excel.","", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void InsertInBD(List<string> lista)
        {
            string cadena = "Host=localhost;port=5432;username=postgres;Password=myps;Database=mybd";
            
                using (var conection = new NpgsqlConnection(cadena))
                {
                   

                    string query = @"INSERT INTO public.cliente(
	                                ""Nombre"", ""Fechacreacion"", ""FechModificacion"", ""Estado"", ""CreadoPor"")
	                                VALUES(@cliente, current_date, current_date, true, 'admin');";
                try
                {
                    using (var comand = new NpgsqlCommand(query, conection))
                    {
                        conection.Open();
                        foreach (var l in lista)
                        {
                            comand.Parameters.Clear();
                            comand.Parameters.AddWithValue("@cliente", l);
                            comand.ExecuteNonQuery();
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw;
                }
                finally
                {
                    conection.Close();
                }
            }
           
           
        }
    }
}

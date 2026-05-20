using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.SQLite;
using System.Data;
using System.IO;
using System.Windows.Forms;
//using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using OfficeOpenXml;
using System.Diagnostics;
using OxyPlot;
using OxyPlot.WindowsForms;
using OxyPlot.Series;
using OxyPlot.Axes;
using System.Data.SQLite.Generic;

namespace FridgeData
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    ///
    public partial class MainWindow : Window
    {

        public string ip = "192.168.31.250";
        public string port = "56000";
        public string path = "data_table.db";
        public string cs = @"URI=file:" + System.AppDomain.CurrentDomain.BaseDirectory + "\\data_table.db";
        public SQLiteConnection con;
        public SQLiteCommand com;
        public SQLiteDataReader dr;
        public float temp_value = 30;
        public thinks u1 = new thinks();
        public thinks u2 = new thinks();
        public DataTable dataTable;
        DataGridView view = new DataGridView();


        public MainWindow()
        {



            InitializeComponent();
            ApiHelper.instal();
            //pathB = System.AppDomain.CurrentDomain.BaseDirectory + "DataBase\\DataBase.accdb";

           
            ip = ip_block.Text;
            port = port_block.Text;

            u1.ip_myThink = ip;
            u1.port_myThink = port;

            DataShow();
            


            u1.ip_myThink=File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory + "\\savesT\\f1.txt");
            u1.port_myThink=File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory + "\\savesT\\f2.txt");
            u2.ip_myThink=File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory + "\\savesT\\f3.txt");
            u2.port_myThink=File.ReadAllText(System.AppDomain.CurrentDomain.BaseDirectory + "\\savesT\\f4.txt");
            
            
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {


            ip = ip_block.Text;
            port = port_block.Text;


            var info = await dataProcessor.LoadInformation(ip, port);

            var info2 = await tempProcessor.LoadInformation(ip, port);


            t1.Text = $"is running {info.IS_RUNNING}";
            t2.Text = $"temp. {(float.Parse(info2.value) / 100)}";
            t3.Text = $"error {info2.error}";
            temp_value = (float.Parse(info2.value) / 100);
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            DataInsert();
            DataShow();


        }

        public void DataShow()
        {
            var con = new SQLiteConnection(cs);
            con.Open();
            string kwd = "Select * from TableDatas";
            var cmd = new SQLiteCommand(kwd, con);
            dr = cmd.ExecuteReader();

            DataTable dt = new DataTable();
            dt.Load(dr);
            dtGrid.ItemsSource = dt.DefaultView;
            dataTable = dt;
            view.DataSource = dataTable.DefaultView;

            
            con.Close();

            



        }



        public void DataInsert()
        {
            var con = new SQLiteConnection(cs);
            con.Open();
            string kwd = "INSERT INTO `TableDatas`(`ID`, `Temperatura`, `Data`, `Godzina`) VALUES(NULL, @temp, @data, @godz)";
            var cmd = new SQLiteCommand(kwd, con);
            cmd.Parameters.AddWithValue("@temp", temp_value);
            cmd.Parameters.AddWithValue("@data", DateTime.Now.ToShortDateString());
            cmd.Parameters.AddWithValue("@godz", DateTime.Now.ToString("HH:mm:ss"));

            //SQLiteDataAdapter dataAdapter = new SQLiteDataAdapter(cmd);
            cmd.ExecuteNonQuery();
            
            con.Close();

        }




        public void klik1(object sender, RoutedEventArgs e)
        {
           
            u2.port_myThink = port_block.Text;
            u2.ip_myThink = ip_block.Text;
            ip_block.Text = u1.ip_myThink;
            port_block.Text = u1.port_myThink;
            SaveFileDialog dialog = new SaveFileDialog();
            File.WriteAllText(dialog.FileName= System.AppDomain.CurrentDomain.BaseDirectory+"\\savesT\\f1.txt", u1.ip_myThink); ;
            File.WriteAllText(dialog.FileName= System.AppDomain.CurrentDomain.BaseDirectory+"\\savesT\\f2.txt", u1.port_myThink);
            

            


        }

        public void klik2(object sender, RoutedEventArgs e)
        {
            u1.port_myThink = port_block.Text;
            u1.ip_myThink = ip_block.Text;
            ip_block.Text = u2.ip_myThink;
            port_block.Text = u2.port_myThink;
            SaveFileDialog dialog = new SaveFileDialog();
            File.WriteAllText(dialog.FileName = System.AppDomain.CurrentDomain.BaseDirectory + "\\savesT\\f3.txt", u2.ip_myThink); ;
            File.WriteAllText(dialog.FileName = System.AppDomain.CurrentDomain.BaseDirectory + "\\savesT\\f4.txt", u2.port_myThink);

        }

       
        

       

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            string connectionString = "Data Source=data_table.db;Version=3;";
            using (SQLiteConnection connection = new SQLiteConnection(connectionString))
            {
                connection.Open();



                string sql = "SELECT * FROM TableDatas";
                using (SQLiteCommand command = new SQLiteCommand(sql, connection))
                {
                    using (SQLiteDataReader reader = command.ExecuteReader())
                    {
                        // Tworzenie nowego pliku Excel
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        ExcelPackage excelPackage = new ExcelPackage();
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Dane");



                        // Wypisanie danych w pionie
                        worksheet.Cells[1, 1].Value = "ID";
                        worksheet.Cells[1, 2].Value = "Temperatura";
                        worksheet.Cells[1, 3].Value = "Data";
                        worksheet.Cells[1, 4].Value = "Godzina";



                        // Wypisanie danych w poziomie
                        int row = 2;
                        while (reader.Read())
                        {
                            worksheet.Cells[row, 1].Value = Convert.ToInt32(reader["ID"]);
                            worksheet.Cells[row, 2].Value = Convert.ToDouble(reader["Temperatura"]);
                            worksheet.Cells[row, 3].Value = reader["Data"].ToString();
                            worksheet.Cells[row, 4].Value = reader["Godzina"].ToString();
                            row++;
                            
                        }



                        // Zapisywanie pliku Excel
                        string filename = "Dane.xlsx";
                        FileInfo fileInfo = new FileInfo(filename);
                        excelPackage.SaveAs(fileInfo);
                        



                    }
                }
            }
            System.Windows.Forms.MessageBox.Show("Plik Excel został utworzony.");
           

        }

        
        public void setValue()
        {
            List<float> valuesF = new List<float>();

            foreach (DataRow row in dataTable.Rows) {

                valuesF.Add ( float.Parse(row[1].ToString()));
                
            }
            
            

            
            

           

            r1.Height = valuesF[valuesF.Count - 1];
            r2.Height = valuesF[valuesF.Count - 2];
            r3.Height = valuesF[valuesF.Count - 3];
            r4.Height = valuesF[valuesF.Count - 4];
            r5.Height = valuesF[valuesF.Count - 5];
            r6.Height = valuesF[valuesF.Count - 6];
            r7.Height = valuesF[valuesF.Count - 7];
            r8.Height = valuesF[valuesF.Count - 8];
            r9.Height = valuesF[valuesF.Count - 9];
            r10.Height = valuesF[valuesF.Count - 10];






        }


        public void wykres()
        {

            SQLite.SQLiteConnection connection_ = new SQLite.SQLiteConnection(cs);
        
            string query = "Select * from TableDatas";

            List<DB> data = connection_.Query<DB>(query);
            List<DateTime> dates = new List<DateTime>();
            List<double> temperatures = new List<double>();

            //tworzenie wykresu
            var plotModel = new PlotModel {Title = "Temperatura"};
            plotModel.Axes.Add(new LinearAxis
            {

                Title = "Temperatura",
                Minimum = -10,
                Maximum = 30,
                MajorStep = 5,
                MinorStep = 1,


            }) ; 





        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            setValue();
        }
    }
}

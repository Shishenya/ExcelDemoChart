using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ChartTest.Enums;
using ExcelDataReader;

namespace ChartTest
{
    public partial class Form1 : Form
    {

        private string _fileName = string.Empty;
        private DataTableCollection _tableExcelCollection = null;
        private CurrentState _currentState = CurrentState.noFileOpen;

        public Form1()
        {
            InitializeComponent();
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult result = openFileDialog1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    _fileName = openFileDialog1.FileName;
                    // Text = _fileName;

                    OpenExcelFile(_fileName);
                    _currentState = CurrentState.yesFileOpen;
                    UpdateLabelState();

                }
                else
                {
                    throw new Exception("Вы не выбрали файл!");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка открытия файла", MessageBoxButtons.OK);
                _currentState = CurrentState.noFileOpen;
            }
        }

        private void OpenExcelFile(string pathFile)
        {
            FileStream stream = File.Open(pathFile, FileMode.Open, FileAccess.Read);

            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);

            DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });

            _tableExcelCollection = db.Tables;
            excelComboBox.Items.Clear();

            // выводим список листов из файла в комбо бокс
            foreach (DataTable dataTable in _tableExcelCollection)
            {
                excelComboBox.Items.Add(dataTable.TableName);
            }

            excelComboBox.SelectedIndex = 0;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            UpdateLabelState();
        }

        private void excelComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable table = _tableExcelCollection[Convert.ToString(excelComboBox.SelectedItem)];
            dataGridView1.DataSource = table;

            SetChartExcelPoints();
        }

        /// <summary>
        /// Обновляет состояние текущего статуса
        /// </summary>
        private void UpdateLabelState()
        {
            switch (_currentState)
            {
                case CurrentState.noFileOpen:
                    labelCurrentState.Text = "Для начала работы выберите файл в меню";
                    break;
                case CurrentState.yesFileOpen:
                    labelCurrentState.Text = $"Вы работаете с файлом: {_fileName}";
                    break;
                default:
                    labelCurrentState.Text = "";
                    break;
            }
        }

        /// <summary>
        /// Строит график прибыли
        /// </summary>
        private void SetChartExcelPoints()
        {
            chartExcel.Series[0].Points.Clear();
            chartExcel.Titles[0].Text = Convert.ToString(excelComboBox.SelectedItem);
            for (int i = 1; i < dataGridView1.Rows.Count; i++)
            {
                if (!Int32.TryParse(dataGridView1[0, i].Value.ToString(), out int x))
                {
                    continue;
                }
                else
                {
                    float y = float.Parse(dataGridView1[1, i].Value.ToString());

                    chartExcel.Series[0].Points.AddXY(x, y);
                }
            }
        }

        /// <summary>
        /// Кнопка выхода
        /// </summary>
        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Выйти из программы?", "Выход", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }
    }
}

using Microsoft.VisualBasic.ApplicationServices;
using System.ComponentModel;
using WinFormsApp11.Models;

namespace WinFormsApp11
{
    public partial class Form1 : Form
    {
        ExcelApplication excel;
        List<UserModel> users;
        public Form1()
        {
            InitializeComponent();
            excel = new ExcelApplication();
            users = new List<UserModel>();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dilog = new OpenFileDialog();
            dilog.Filter = "xlsx files|*.xlsx";

            if(dilog.ShowDialog() == DialogResult.OK)
            {
                excel.InitWorkbook(dilog.FileName);
                var result = excel.GetUsers();
                users = result.ToList();
                ShowData();
            }
        }

        void ShowData()
        {
            var bindingList = new BindingList<UserModel>(users);
            var source = new BindingSource(bindingList, null);
            dataGridView1.DataSource = source;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            excel.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var arr = users.ToArray();
            excel.SetUsers(arr);
        }
    }
}
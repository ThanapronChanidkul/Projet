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
using System.Xml.Linq;
using OfficeOpenXml;

namespace ProjectCode1
{
    public partial class ShopCar : Form
    {
        public string name;
        public int amount;
        public string color;
        public ShopCar()
        {
            InitializeComponent();
            combocolor1.Items.Add("black");
            combocolor1.Items.Add("white");
            combocolor2.Items.Add("white");
            combocolor3.Items.Add("white");
            combocolor4.Items.Add("black");
            combocolor4.Items.Add("red");
            combocolor5.Items.Add("white");
            combocolor5.Items.Add("black");
            combocolor5.Items.Add("grey");
            combocolor6.Items.Add("grey");
            combocolor6.Items.Add("black");
        }

        private void button1_Click_2(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (Int32.Parse(textBox1.Text) >= 1)
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                name = pictureBox9.Text;
                amount = Int32.Parse(this.textBox1.Text);
                color = combocolor1.Text;
                Car car = new Car(name, amount, color);
                var message = $"{car.getName()} จำนวน {car.getAmount} สี\n {car.getMeat}";
                using (var package = new ExcelPackage(new FileInfo(@"C:\Users\User\Documents\Visual Studio 2022\Vs\ProjectCode1\ProjectCode1\SoBer.xlsx")))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["sheet1"];
                    if (worksheet != null)
                    {
                        throw new Exception("Worksheet does not exist");
                    }
                    int lastUsedRow = worksheet.Dimension.End.Row;

                    ExcelRange cell = worksheet.Cells[lastUsedRow + 1, 1];
                    cell.Value = message;

                    package.Save();
                }
                this.DialogResult = DialogResult.OK;
                this.Hide();
                MessageBox.Show(message);
            }
            if (Int32.Parse(textBox2.Text) >= 1)
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                name = pictureBox9.Text;
                amount = Int32.Parse(this.textBox1.Text);
                color = combocolor1.Text;
                Car car = new Car(name, amount, color);
                var message = $"{car.getName()} ราคา {car.getAmount} สี\n {car.getMeat}";
                using (var package = new ExcelPackage(new FileInfo(@"C:\Users\User\Documents\Visual Studio 2022\Vs\ProjectCode1\ProjectCode1\SoBer.xlsx")))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["sheet1"];
                    if (worksheet != null)
                    {
                        throw new Exception("Worksheet does not exist");
                    }
                    int lastUsedRow = worksheet.Dimension.End.Row;

                    ExcelRange cell = worksheet.Cells[lastUsedRow + 1, 1];
                    cell.Value = message;

                    package.Save();
                }
                this.DialogResult = DialogResult.OK;
                this.Hide();
                MessageBox.Show(message);
            }
            if (Int32.Parse(textBox3.Text) >= 1)
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                name = pictureBox9.Text;
                amount = Int32.Parse(this.textBox1.Text);
                color = combocolor1.Text;
                Car car = new Car(name, amount, color);
                var message = $"{car.getName()} ราคา {car.getAmount} สี\n {car.getMeat}";
                using (var package = new ExcelPackage(new FileInfo(@"C:\Users\User\Documents\Visual Studio 2022\Vs\ProjectCode1\ProjectCode1\SoBer.xlsx")))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["sheet1"];
                    if (worksheet != null)
                    {
                        throw new Exception("Worksheet does not exist");
                    }
                    int lastUsedRow = worksheet.Dimension.End.Row;

                    ExcelRange cell = worksheet.Cells[lastUsedRow + 1, 1];
                    cell.Value = message;

                    package.Save();
                }
                this.DialogResult = DialogResult.OK;
                this.Hide();
                MessageBox.Show(message);
            }
            if (Int32.Parse(textBox4.Text) >= 1)
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                name = pictureBox9.Text;
                amount = Int32.Parse(this.textBox1.Text);
                color = combocolor1.Text;
                Car car = new Car(name, amount, color);
                var message = $"{car.getName()} ราคา {car.getAmount} สี\n {car.getMeat}";
                using (var package = new ExcelPackage(new FileInfo(@"C:\Users\User\Documents\Visual Studio 2022\Vs\ProjectCode1\ProjectCode1\SoBer.xlsx")))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["sheet1"];
                    if (worksheet != null)
                    {
                        throw new Exception("Worksheet does not exist");
                    }
                    int lastUsedRow = worksheet.Dimension.End.Row;

                    ExcelRange cell = worksheet.Cells[lastUsedRow + 1, 1];
                    cell.Value = message;

                    package.Save();
                }
                this.DialogResult = DialogResult.OK;
                this.Hide();
                MessageBox.Show(message);
            }
            if (Int32.Parse(textBox5.Text) >= 1)
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                name = pictureBox9.Text;
                amount = Int32.Parse(this.textBox1.Text);
                color = combocolor1.Text;
                Car car = new Car(name, amount, color);
                var message = $"{car.getName()} ราคา {car.getAmount} สี\n {car.getMeat}";
                using (var package = new ExcelPackage(new FileInfo(@"C:\Users\User\Documents\Visual Studio 2022\Vs\ProjectCode1\ProjectCode1\SoBer.xlsx")))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["sheet1"];
                    if (worksheet != null)
                    {
                        throw new Exception("Worksheet does not exist");
                    }
                    int lastUsedRow = worksheet.Dimension.End.Row;

                    ExcelRange cell = worksheet.Cells[lastUsedRow + 1, 1];
                    cell.Value = message;

                    package.Save();
                }
                this.DialogResult = DialogResult.OK;
                this.Hide();
                MessageBox.Show(message);
            }
            if (Int32.Parse(textBox6.Text) >= 1)
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                name = pictureBox9.Text;
                amount = Int32.Parse(this.textBox1.Text);
                color = combocolor1.Text;
                Car car = new Car(name, amount, color);
                var message = $"{car.getName()} ราคา {car.getAmount} สี\n {car.getMeat}";
                using (var package = new ExcelPackage(new FileInfo(@"C:\Users\User\Documents\Visual Studio 2022\Vs\ProjectCode1\ProjectCode1\SoBer.xlsx")))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets["sheet1"];
                    if (worksheet != null)
                    {
                        throw new Exception("Worksheet does not exist");
                    }
                    int lastUsedRow = worksheet.Dimension.End.Row;

                    ExcelRange cell = worksheet.Cells[lastUsedRow + 1, 1];
                    cell.Value = message;

                    package.Save();
                }
                this.DialogResult = DialogResult.OK;
                this.Hide();
                MessageBox.Show(message);
            }
        }
    }
}

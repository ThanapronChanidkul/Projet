using OfficeOpenXml;

namespace ProjectCode1
{
    public partial class ShopBike : Form
    {
        public string name;
        public int price;
        public string color;
        public ShopBike()
        {
            InitializeComponent();
            combocolor1.Items.Add("black");
            combocolor2.Items.Add("white");
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void combocolor7_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (Int32.Parse(textBox7.Text) >= 1)
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                name = pictureBox9.Text;
                price = 0;
                color = combocolor1.Text;
                Car car = new Car(name, price, color);
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
            if (Int32.Parse(textBox8.Text) >= 1)
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
                name = pictureBox9.Text;
                price = 0;
                color = combocolor1.Text;
                Car car = new Car(name, price, color);
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

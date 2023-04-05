
using OfficeOpenXml;

namespace ProjectCode1
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.Commercial;
            using (ExcelPackage package = new ExcelPackage(new FileInfo(@"C:\Users\User\Documents\Visual Studio 2022\Vs\ProjectCode1\ProjectCode1\SoBer.xlsx")))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                int lastUsedRow = worksheet.Dimension.End.Row;
                for (int row = 1; row <= lastUsedRow; row++)
                {
                    string value = worksheet.Cells[row, 1].Value?.ToString();
                    if (!string.IsNullOrEmpty(value))
                    {
                        this.dataGridView1.Rows.Add(value);
                    }
                }

            }
        }

        private void pictureBox10_Click(object sender, EventArgs e)
        {

        }

        private void addCarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShopCar shopcar = new ShopCar();
            shopcar.ShowDialog();
        }

        private void addToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShopBike shopbike = new ShopBike();
            shopbike.ShowDialog();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
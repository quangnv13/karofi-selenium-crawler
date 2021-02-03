using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
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

namespace Crawler
{
    public partial class Form1 : Form
    {
        ChromeDriver chromeDriver;
        List<DataCraw> lstData = new List<DataCraw>();
        public Form1()
        {
            InitializeComponent();
            chromeDriver = new ChromeDriver();
            chromeDriver.Url = "https://karofi.com/he-thong-phan-phoi";
            chromeDriver.Navigate();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                lstData = new List<DataCraw>();
                var info = chromeDriver.FindElementByClassName("agency-content__list");
                var itemInfos = info.FindElements(By.ClassName("item-info"));
                foreach (var item in itemInfos)
                {
                    string address = item.FindElements(By.ClassName("item-info__line-text"))[0].Text;
                    string phone = item.FindElements(By.ClassName("item-info__line-text"))[1].Text;
                    string locationLink = item.FindElements(By.TagName("a"))[1].GetAttribute("href");
                    DataCraw data = new DataCraw()
                    {
                        Address = address,
                        Phone = phone,
                        LocationLink = locationLink
                    };
                    lstData.Add(data);
                }
                dataGridView1.DataSource = lstData;
            }
            catch
            {
                MessageBox.Show("Có lỗi xảy ra! Vui lòng thử lại hoặc restart tool");
            }

            string filePath = "";
            // tạo SaveFileDialog để lưu file excel
            SaveFileDialog dialog = new SaveFileDialog();

            // chỉ lọc ra các file có định dạng Excel
            dialog.Filter = "Excel | *.xlsx | Excel 2003 | *.xls";

            // Nếu mở file và chọn nơi lưu file thành công sẽ lưu đường dẫn lại dùng
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                filePath = dialog.FileName;
            }

            // nếu đường dẫn null hoặc rỗng thì báo không hợp lệ và return hàm
            if (string.IsNullOrEmpty(filePath))
            {
                MessageBox.Show("Đường dẫn báo cáo không hợp lệ");
                return;
            }

            try
            {
                using (ExcelPackage p = new ExcelPackage())
                {
                    // đặt tên người tạo file
                    p.Workbook.Properties.Author = "QuangNV";

                    // đặt tiêu đề cho file
                    p.Workbook.Properties.Title = "Data craw";

                    //Tạo một sheet để làm việc trên đó
                    p.Workbook.Worksheets.Add("Data");

                    // lấy sheet vừa add ra để thao tác
                    ExcelWorksheet ws = p.Workbook.Worksheets[0];

                    // đặt tên cho sheet
                    ws.Name = "Data";
                    // fontsize mặc định cho cả sheet
                    ws.Cells.Style.Font.Size = 11;
                    // font family mặc định cho cả sheet
                    ws.Cells.Style.Font.Name = "Calibri";

                    // Tạo danh sách các column header
                    string[] arrColumnHeader = {
                                                "Địa chỉ",
                                                "Số điện thoại",
                                                "Vị trí"
                };

                    // lấy ra số lượng cột cần dùng dựa vào số lượng header
                    var countColHeader = arrColumnHeader.Count();

                    int colIndex = 1;
                    int rowIndex = 1;

                    //tạo các header từ column header đã tạo từ bên trên
                    foreach (var item in arrColumnHeader)
                    {
                        var cell = ws.Cells[rowIndex, colIndex];

                        //set màu thành gray
                        var fill = cell.Style.Fill;
                        fill.PatternType = ExcelFillStyle.Solid;
                        fill.BackgroundColor.SetColor(Color.LightBlue);

                        //căn chỉnh các border
                        var border = cell.Style.Border;
                        border.Bottom.Style =
                            border.Top.Style =
                            border.Left.Style =
                            border.Right.Style = ExcelBorderStyle.Thin;

                        //gán giá trị
                        cell.Value = item;

                        colIndex++;
                    }

                    // với mỗi item trong danh sách sẽ ghi trên 1 dòng
                    foreach (var item in lstData)
                    {
                        // bắt đầu ghi từ cột 1. Excel bắt đầu từ 1 không phải từ 0
                        colIndex = 1;

                        // rowIndex tương ứng từng dòng dữ liệu
                        rowIndex++;

                        //gán giá trị cho từng cell                      
                        ws.Cells[rowIndex, colIndex++].Value = item.Address;
                        ws.Cells[rowIndex, colIndex++].Value = item.Phone;
                        ws.Cells[rowIndex, colIndex++].Value = item.LocationLink;
                    }

                    //Lưu file lại
                    Byte[] bin = p.GetAsByteArray();
                    File.WriteAllBytes(filePath, bin);
                }
                MessageBox.Show("Xuất excel thành công!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi lưu file!", ex.Message);
            }
        }
    }

    public class DataCraw
    {
        public string Address { get; set; }
        public string Phone { get; set; }
        public string LocationLink { get; set; }
    }
}

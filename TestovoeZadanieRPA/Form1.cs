using OfficeOpenXml;
using System;
using System.ComponentModel;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestovoeZadanieRPA
{
    public partial class Form1 : Form
    {
        String urlUSD = "https://www.moex.com/ru/derivatives/currency-rate.aspx?currency=USD_RUB";
        String urlEUR = "https://www.moex.com/ru/derivatives/currency-rate.aspx?currency=EUR_RUB";

        String server = "smtp.gmail.com";
        int port = 587;

        public Form1()
        {
            InitializeComponent();

            saveFile.Filter = "Text files(*.xlsx)|*.xlsx";
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            rateTable.Columns.Add("usdDate", "Дата");
            rateTable.Columns.Add("usdRate", "Курс");
            rateTable.Columns.Add("usdСhanges", "Изменения");
            rateTable.Columns.Add("eurDate", "Дата");
            rateTable.Columns.Add("eurRate", "Курс");
            rateTable.Columns.Add("eurСhanges", "Изменения");

            foreach (DataGridViewColumn column in rateTable.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            Thread getRates = new Thread(SetRates);
            getRates.Start();
        }

        void SetRates()
        {
            MatchCollection usdCollection = GetRates(urlUSD);
            MatchCollection eurCollection = GetRates(urlEUR);

            if (usdCollection != null)
            {
                Invoke((MethodInvoker)(() =>
                {
                    rateTable.Rows.Clear();

                    double previousValueUSD = 0;
                    double previousValueEUR = 0;

                    for (int i = 0; i < (eurCollection.Count >= usdCollection.Count ? eurCollection.Count : usdCollection.Count); i++)
                    {
                        rateTable.Rows.Add();
                        rateTable.Rows.Add();

                        previousValueUSD = SetData(i, 0, previousValueUSD, usdCollection, rateTable);
                        previousValueEUR = SetData(i, 3, previousValueEUR, eurCollection, rateTable);
                    }
                }));
            }
        }

        double SetData(int i, int cellNumber, double previousValue, MatchCollection collection, DataGridView dataGridView)
        {
            try
            {
                if (collection.Count > i)
                {
                    var date = Regex.Match(collection[i].Value, @"(?<=>).*?(?=<)");
                    var values = Regex.Matches(collection[i].Value, @"(?<=<td align='center'>).*?(?=<\/td)");

                    i = i * 2;

                    if (values[1].Value != "-")
                        dataGridView.Rows[i].Cells[cellNumber].Value = DateTime.Parse($"{date} {values[1]}");
                    else
                        dataGridView.Rows[i].Cells[cellNumber].Value = DateTime.Parse($"{date}");

                    if (values[3].Value != "-")
                        dataGridView.Rows[i + 1].Cells[cellNumber].Value = DateTime.Parse($"{date} {values[3]}");
                    else
                        dataGridView.Rows[i + 1].Cells[cellNumber].Value = DateTime.Parse($"{date}");

                    if (values[0].ToString() == "-")
                        dataGridView.Rows[i].Cells[cellNumber + 1].Value = "-";
                    else
                        dataGridView.Rows[i].Cells[cellNumber + 1].Value = double.Parse(values[0].ToString());
                    if (values[2].ToString() == "-")
                        dataGridView.Rows[i + 1].Cells[cellNumber + 1].Value = "-";
                    else
                        dataGridView.Rows[i + 1].Cells[cellNumber + 1].Value = double.Parse(values[2].ToString());

                    if (values[0].ToString() != "-")
                    {
                        previousValue = CalculationChange(dataGridView, i, cellNumber, previousValue, double.Parse(values[0].ToString()));

                        if (values[2].ToString() != "-")
                            previousValue = CalculationChange(dataGridView, i + 1, cellNumber, previousValue, double.Parse(values[2].ToString()));
                        else
                            dataGridView.Rows[i + 1].Cells[cellNumber + 2].Value = "-";

                        return previousValue;
                    }
                    else if (values[2].ToString() != "-")
                    {
                        dataGridView.Rows[i].Cells[cellNumber + 2].Value = "-";

                        return CalculationChange(dataGridView, i + 1, cellNumber, previousValue, double.Parse(values[2].ToString()));
                    }
                    else
                    {
                        dataGridView.Rows[i].Cells[cellNumber + 2].Value = "-";
                        dataGridView.Rows[i + 1].Cells[cellNumber + 2].Value = "-";
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            return previousValue;
        }

        double CalculationChange(DataGridView dataGridView, int i, int cellNumber, double previousValue, double currentValue)
        {
            var change = currentValue - previousValue;

            if (previousValue == 0)
                dataGridView.Rows[i].Cells[cellNumber + 2].Value = "-";
            else
                SetChangeInCell(dataGridView, i, cellNumber + 2, change);

            return currentValue;
        }

        void SetChangeInCell(DataGridView dataGridView, int row, int cell, double value)
        {
            dataGridView.Rows[row].Cells[cell].Value = value;
            dataGridView.Rows[row].Cells[cell].Style.ForeColor = value >= 0 ? System.Drawing.Color.Green : System.Drawing.Color.Red;
        }

        MatchCollection GetRates(String url)
        {
            try
            {
                System.Net.WebClient webClient = new System.Net.WebClient();
                var responce = webClient.DownloadString(url);
                return Regex.Matches(responce, @"<tr class=tr(1|0) align=right>[\S\s]*?<\/tr>");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return null;
            }
        }

        private void update_Click(object sender, EventArgs e)
        {
            Thread getRates = new Thread(SetRates);
            getRates.Start();
        }

        private async void excel_Click(object sender, EventArgs e)
        {
            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                var filename = saveFile.FileName;
                var file = new FileInfo(filename);

                try
                {
                    await ExportToExcel(rateTable, file);
                }
                catch (Exception err)
                {
                    MessageBox.Show(err.Message);
                }
            }
            return;
        }

        private static async Task ExportToExcel(DataGridView dataGrid, FileInfo file)
        {
            DeleteIfExist(file);

            using (var package = new ExcelPackage(file))
            {
                await DataSetToExcel(package, dataGrid).SaveAsync();
            }            
        }

        private static ExcelPackage DataSetToExcel(ExcelPackage package, DataGridView dataGrid)
        {
            var ws = package.Workbook.Worksheets.Add("Курс");
            ws.Cells[1, 1].Value = "Дата";
            ws.Cells[1, 2].Value = "Курс";
            ws.Cells[1, 3].Value = "Изменения";
            ws.Cells[1, 4].Value = "Дата";
            ws.Cells[1, 5].Value = "Курс";
            ws.Cells[1, 6].Value = "Изменения";
            ws.Cells[1, 7].Value = "Частное";



            for (int row = 0; row < dataGrid.RowCount; row++)
            {
                for (int column = 0; column < 3; column++)
                {
                    ws.Cells[row + 2, column + 1].Value = dataGrid.Rows[row].Cells[column].Value;

                    if (column != 0 && column != 3 && column != 6)
                        ws.Cells[row + 2, column + 1].Style.Numberformat.Format = "#,##0.00 ₽";
                    else if (column != 6)
                        ws.Cells[row + 2, column + 1].Style.Numberformat.Format = "dd/mm/yyyy hh:mm ";
                }

                for (int equals = 0; equals < dataGrid.RowCount; equals++)
                {
                    if (ws.Cells[row + 2, 1].Value != null && dataGrid.Rows[equals].Cells[3].Value != null)
                        if (ws.Cells[row + 2, 1].Value.ToString() == dataGrid.Rows[equals].Cells[3].Value.ToString())
                        {
                            for (int column = 3; column < 6; column++)
                            {
                                ws.Cells[row + 2, column + 1].Value = dataGrid.Rows[equals].Cells[column].Value;

                                if (column != 0 && column != 3 && column != 6)
                                    ws.Cells[row + 2, column + 1].Style.Numberformat.Format = "#,##0.00 ₽";
                                else if (column != 6)
                                    ws.Cells[row + 2, column + 1].Style.Numberformat.Format = "dd/mm/yyyy hh:mm ";
                            }
                        }
                }
                ws.Cells[row + 2, 7].Formula = $"=E{row + 2}/B{row + 2}";
            }

            for (int i = 1; i < 8; i++)
            {
                ws.Cells[1, i].Style.Font.Size = 14;
                ws.Cells[1, i].Style.Font.Name = "Arial";
                ws.Cells[1, i].Style.Font.Bold = true;
                ws.Cells[1, i].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                ws.Column(i).AutoFit();
            }

            return package;
        }

        private static void DeleteIfExist(FileInfo file)
        {
            if (file.Exists)
                file.Delete();
        }

        private void saveFileDialog1_FileOk(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }

        private async void email_Click(object sender, EventArgs e)
        {
            var file = new FileInfo(Directory.GetCurrentDirectory() + @"\rate.xlsx");

            try
            {
                await ExportToExcel(rateTable, file);
            }
            catch (Exception err)
            {
                MessageBox.Show(err.Message);
            }

            SmtpClient client;

            if (Regex.IsMatch(login.Text.Trim(), @"^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$") &&
                Regex.IsMatch(emailTo.Text.Trim(), @"^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$"))
            {
                try
                {
                    NetworkCredential user = new NetworkCredential(login.Text, password.Text);

                    client = new SmtpClient(server, port);
                    client.Credentials = user;
                    client.EnableSsl = true;

                    MailMessage Mail = new MailMessage(login.Text, emailTo.Text.Trim());
                    Mail.Subject = "Курсы валют";

                    if (rateTable.RowCount % 10 == 1 && rateTable.RowCount % 100 != 11)
                        Mail.Body = $"В файле {rateTable.RowCount} строка с курсами валют";
                    else if ((rateTable.RowCount % 10 > 1) && (rateTable.RowCount % 10 < 5) && (rateTable.RowCount % 100 > 11) && (rateTable.RowCount % 100 < 15))
                        Mail.Body = $"В файле {rateTable.RowCount} строки с курсами валют";
                    else
                        Mail.Body = $"В файле {rateTable.RowCount} строк с курсами валют";

                    Attachment attach = new Attachment(Directory.GetCurrentDirectory() + @"\rate.xlsx", MediaTypeNames.Application.Octet);
                    Mail.Attachments.Add(attach);
                    Mail.IsBodyHtml = true;
                    Mail.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure;

                    client.SendCompleted += new SendCompletedEventHandler(SendCompletedCallback);
                    client.Send(Mail);

                    MessageBox.Show("Сообщение  отправлено", "Message", MessageBoxButtons.OK);
                }
                catch
                {
                    MessageBox.Show("Ошибка ввода данных", "Message", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Логин почты неверный!");
            }
        }
        private static void SendCompletedCallback(object sender, AsyncCompletedEventArgs e)
        {
            if (e.Cancelled)
                MessageBox.Show(string.Format("Отправка отменена.", e.UserState), "Message", MessageBoxButtons.OK);
            if (e.Error != null)
                MessageBox.Show(string.Format("Отправка отменена", e.UserState, e.Error), "Message", MessageBoxButtons.OK);
        }
    }
}

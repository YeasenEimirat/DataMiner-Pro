using System;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;
using System.Text;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;


namespace DataMiner_Pro
{
    public partial class Form1 : Form
    {
        private static readonly HttpClient client = new HttpClient();
        private List<string[]> dataList = new List<string[]>();

        public Form1()
        {
            InitializeComponent();

            client.DefaultRequestHeaders.Add("Accept", "application/json");
        }

        // =========================
        // GET COORDINATES
        // =========================
        public async Task<(double lat, double lng)> GetCoordinates(string city)
        {
            try
            {
                string url = $"https://nominatim.openstreetmap.org/search?q={city}&format=json";

                client.DefaultRequestHeaders.UserAgent.ParseAdd("GeoLeadsApp");

                string response = await client.GetStringAsync(url);

                JArray json = JArray.Parse(response);

                if (json.Count == 0)
                    return (0, 0);

                return (
                    Convert.ToDouble(json[0]["lat"]),
                    Convert.ToDouble(json[0]["lon"])
                );
            }
            catch
            {
                return (0, 0);
            }
        }

        // =========================
        // START BUTTON
        // =========================
        private async void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.Rows.Clear();
                dataList.Clear();

                // =========================
                // INPUT VALIDATION
                // =========================
                if (string.IsNullOrWhiteSpace(txtCity.Text))
                     
                {
                    MessageBox.Show("Please fill all fields");
                    return;
                }

                string city = txtCity.Text.Trim();
                string search = cmbSearchType.SelectedItem.ToString();
                string country = txtCountry.Text.Trim();

                // =========================
                // GET COORDINATES
                // =========================
                var coords = await GetCoordinates(city);

                if (coords.lat == 0 || coords.lng == 0)
                {
                    MessageBox.Show("City not found!");
                    return;
                }

                 string apiKey = "apify_api_xxxxxxxxxxxxxxxxx";

 

                if (string.IsNullOrWhiteSpace(apiKey))
                {
                    MessageBox.Show("Missing API Key!");
                    return;
                }
 
                string url =
                    $"https://api.apify.com/v2/acts/compass~crawler-google-places/run-sync-get-dataset-items?token={apiKey}";
 
                var body = new JObject
                {
                    ["searchStringsArray"] = new JArray($"{search} {city}"),
                    ["lat"] = coords.lat,
                    ["lng"] = coords.lng,
                    ["zoom"] = 14,
                    ["language"] = "ar",
                    ["countryCode"] = country,
                    ["maxCrawledPlacesPerSearch"] = 50,
                    ["searchTightLocation"] = true,
                    ["skipClosedPlaces"] = false
                };

                var content = new StringContent(body.ToString(), Encoding.UTF8, "application/json");

                var response = await client.PostAsync(url, content);
                string result = await response.Content.ReadAsStringAsync();

                if (!result.TrimStart().StartsWith("["))
                {
                    MessageBox.Show("API Error:\n" + result);
                    return;
                }

                JArray data = JArray.Parse(result);

                foreach (var item in data)
                {
                    string name = item["title"]?.ToString();
                    string phone = item["phoneUnformatted"]?.ToString();
                    string cityName = item["city"]?.ToString();
                    string ratingText = item["totalScore"]?.ToString();
                    string website = item["website"]?.ToString();

                    if (string.IsNullOrWhiteSpace(name))
                        continue;

                    double rating = 0;
                    double.TryParse(ratingText, out rating);

                     if (rating < 3.5)
                        continue;

                     if (dataList.Any(x => x[0] == name))
                        continue;

                    dataGridView1.Rows.Add(name, phone, cityName, ratingText);

                    dataList.Add(new string[]
                    {
                        name,
                        phone ?? "",
                        cityName ?? "",
                        ratingText ?? "",
                        website ?? ""
                    });
                }

                MessageBox.Show($"Done! Found {dataList.Count} results.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error:\n" + ex.Message);
            }
        }

  
 
private void btnExport_Click(object sender, EventArgs e)
    {
        try
        {
             btnExport.Enabled = false;
            this.Cursor = Cursors.WaitCursor;

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "Excel Files (*.xlsx)|*.xlsx";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    var workbook = new XLWorkbook();
                    var worksheet = workbook.Worksheets.Add("Data");

                    
                    worksheet.Cell(1, 1).Value = "Name";
                    worksheet.Cell(1, 2).Value = "Phone";
                    worksheet.Cell(1, 3).Value = "City";
                    worksheet.Cell(1, 4).Value = "Rating";
                    worksheet.Cell(1, 5).Value = "Website";

                    // تنسيق الهيدر (اختياري)
                    var header = worksheet.Range("A1:E1");
                    header.Style.Font.Bold = true;

                    
                    for (int i = 0; i < dataList.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = dataList[i][0];

                        // 🔥 مهم: رقم الهاتف كنص
                        worksheet.Cell(i + 2, 2).Value = dataList[i][1];
                            var cell = worksheet.Cell(i + 2, 2);
                            cell.Value = dataList[i][1];
                            cell.Style.NumberFormat.Format = "@";
                            worksheet.Cell(i + 2, 3).Value = dataList[i][2];
                        worksheet.Cell(i + 2, 4).Value = dataList[i][3];
                        worksheet.Cell(i + 2, 5).Value = dataList[i][4];
                    }

            
                    worksheet.Columns().AdjustToContents();

                   
                    workbook.SaveAs(sfd.FileName);
                }
            }

            MessageBox.Show("Export Done!");
        }
        catch (Exception ex)
        {
            MessageBox.Show("Error:\n" + ex.Message);
        }
        finally
        {
            btnExport.Enabled = true;
            this.Cursor = Cursors.Default;
        }
    }
    private void SetupComboBox()
        {
            cmbSearchType.Items.Add("Restaurants");
            cmbSearchType.Items.Add("Hotels");
            cmbSearchType.Items.Add("Cafes");
            cmbSearchType.Items.Add("Clinics");
            cmbSearchType.Items.Add("Supermarkets");
            cmbSearchType.Items.Add("Pharmacies");

            cmbSearchType.SelectedIndex = 0;
        }
        private void SetupDataGridView()
        {
            dataGridView1.Columns.Clear();

            dataGridView1.Columns.Add("Name", "Name");
            dataGridView1.Columns.Add("Phone", "Phone");
            dataGridView1.Columns.Add("City", "City");
            dataGridView1.Columns.Add("Rating", "Rating");
            dataGridView1.Columns.Add("Website", "Website");
            dataGridView1.Columns["Name"].Width = 180;
            dataGridView1.Columns["Phone"].Width = 150;
            dataGridView1.Columns["City"].Width = 150;
            dataGridView1.Columns["Rating"].Width = 80;
            dataGridView1.Columns["Website"].Width = 300;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            SetupComboBox();
            SetupDataGridView();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
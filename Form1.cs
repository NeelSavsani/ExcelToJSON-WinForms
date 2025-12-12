using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using ExcelDataReader;          // Excel reader
using Newtonsoft.Json;          // JSON

namespace SampleProject1
{
    public partial class Form1 : Form
    {
        private string _selectedExcelPath;
        private string _selectedJsonPath;

        // extra UI elements created in code
        private Panel pnlCard;
        private Label lblTitle;
        private Label lblSubtitle;
        private Label lblHint;

        // JSON picker controls
        private Label lblJsonPath;
        private Button btnBrowseJson;
        private Button btnCallApi;

        // API URL + header
        private const string ApiUrl = "https://orderapp.neosoftservice.in/api/Sync/SyncProduct";
        private const string NeoSoftCodeHeaderName = "NeoSoftCode";
        private const string NeoSoftCodeHeaderValue = "b4a482ca-6660-4cb9-96af-f1adb0b69dd1";

        public Form1()
        {
            InitializeComponent();
            SetupModernUi();
        }

        // ===========================================================
        // MODERN UI LAYOUT
        // ===========================================================
        private void SetupModernUi()
        {
            // ----- Form look -----
            this.Text = "Excel → JSON Converter & API Uploader";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.Gainsboro;
            this.Font = new Font("Segoe UI", 10F);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;

            // ----- Main card panel -----
            pnlCard = new Panel();
            pnlCard.BackColor = Color.White;
            pnlCard.Size = new Size(700, 340);
            pnlCard.BorderStyle = BorderStyle.FixedSingle;
            CenterCardPanel();
            this.Controls.Add(pnlCard);

            // Re-center when form resizes
            this.Resize += (s, e) => CenterCardPanel();

            // ----- Title -----
            lblTitle = new Label();
            lblTitle.Text = "Excel → JSON Converter & API Uploader";
            lblTitle.Font = new Font("Segoe UI", 14F, FontStyle.Bold);
            lblTitle.Dock = DockStyle.Top;
            lblTitle.Height = 40;
            lblTitle.TextAlign = ContentAlignment.MiddleCenter;
            pnlCard.Controls.Add(lblTitle);

            // ----- Subtitle -----
            lblSubtitle = new Label();
            lblSubtitle.Text = "Convert Excel to JSON (merged) and send products to API in batches of 1000.";
            lblSubtitle.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            lblSubtitle.Dock = DockStyle.Top;
            lblSubtitle.Height = 30;
            lblSubtitle.TextAlign = ContentAlignment.MiddleCenter;
            lblSubtitle.ForeColor = Color.DimGray;
            pnlCard.Controls.Add(lblSubtitle);

            int baseY = 80;

            // ================== EXCEL PICKER ==================
            lblFilePath.AutoSize = false;
            lblFilePath.BorderStyle = BorderStyle.FixedSingle;
            lblFilePath.BackColor = Color.White;
            lblFilePath.TextAlign = ContentAlignment.MiddleLeft;
            lblFilePath.Text = "No file selected";
            lblFilePath.Width = 420;
            lblFilePath.Height = 28;
            lblFilePath.Location = new Point(60, baseY);
            lblFilePath.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            pnlCard.Controls.Add(lblFilePath); // move into panel

            button1.Text = "Browse Excel";
            button1.Width = 130;
            button1.Height = 28;
            button1.Location = new Point(lblFilePath.Right + 15, baseY);
            button1.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            button1.FlatStyle = FlatStyle.Flat;
            button1.FlatAppearance.BorderSize = 0;
            button1.BackColor = Color.DodgerBlue;
            button1.ForeColor = Color.White;
            button1.Font = new Font("Segoe UI", 9F, FontStyle.Bold);

            pnlCard.Controls.Add(button1); // move into panel

            lblHint = new Label();
            lblHint.AutoSize = true;
            lblHint.Text = "Select Excel (.xlsx, .xls) to convert → merged JSON (scheme logic, longest expiry, etc.)";
            lblHint.Font = new Font("Segoe UI", 8.5F);
            lblHint.ForeColor = Color.Gray;
            lblHint.Location = new Point(lblFilePath.Left + 4, baseY + 32);
            pnlCard.Controls.Add(lblHint);

            ConvertBtn.Text = "Convert Excel to JSON";
            ConvertBtn.Width = 220;
            ConvertBtn.Height = 34;
            ConvertBtn.FlatStyle = FlatStyle.Flat;
            ConvertBtn.FlatAppearance.BorderSize = 0;
            ConvertBtn.BackColor = Color.MediumSeaGreen;
            ConvertBtn.ForeColor = Color.White;
            ConvertBtn.Font = new Font("Segoe UI", 10F, FontStyle.Bold);

            int convertX = (pnlCard.Width - ConvertBtn.Width) / 2;
            int convertY = baseY + 65;
            ConvertBtn.Location = new Point(convertX, convertY);
            ConvertBtn.Anchor = AnchorStyles.Top;
            pnlCard.Controls.Add(ConvertBtn);

            // ================== JSON PICKER & API CALL ==================
            int jsonBaseY = convertY + 55;

            lblJsonPath = new Label();
            lblJsonPath.AutoSize = false;
            lblJsonPath.BorderStyle = BorderStyle.FixedSingle;
            lblJsonPath.BackColor = Color.White;
            lblJsonPath.TextAlign = ContentAlignment.MiddleLeft;
            lblJsonPath.Text = "No file selected";
            lblJsonPath.Width = 420;
            lblJsonPath.Height = 28;
            lblJsonPath.Location = new Point(60, jsonBaseY);
            lblJsonPath.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            pnlCard.Controls.Add(lblJsonPath);

            btnBrowseJson = new Button();
            btnBrowseJson.Text = "Browse JSON";
            btnBrowseJson.Width = 130;
            btnBrowseJson.Height = 28;
            btnBrowseJson.Location = new Point(lblJsonPath.Right + 15, jsonBaseY);
            btnBrowseJson.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            btnBrowseJson.FlatStyle = FlatStyle.Flat;
            btnBrowseJson.FlatAppearance.BorderSize = 0;
            btnBrowseJson.BackColor = Color.SteelBlue;
            btnBrowseJson.ForeColor = Color.White;
            btnBrowseJson.Font = new Font("Segoe UI", 9F, FontStyle.Bold);

            btnBrowseJson.Click += BtnBrowseJson_Click;
            pnlCard.Controls.Add(btnBrowseJson);

            btnCallApi = new Button();
            btnCallApi.Text = "Call API (Upload JSON)";
            btnCallApi.Width = 220;
            btnCallApi.Height = 34;
            btnCallApi.FlatStyle = FlatStyle.Flat;
            btnCallApi.FlatAppearance.BorderSize = 0;
            btnCallApi.BackColor = Color.DarkOrange;
            btnCallApi.ForeColor = Color.White;
            btnCallApi.Font = new Font("Segoe UI", 10F, FontStyle.Bold);

            int apiX = (pnlCard.Width - btnCallApi.Width) / 2;
            int apiY = jsonBaseY + 40;
            btnCallApi.Location = new Point(apiX, apiY);
            btnCallApi.Anchor = AnchorStyles.Top;
            btnCallApi.Click += BtnCallApi_Click;

            pnlCard.Controls.Add(btnCallApi);
        }

        private void CenterCardPanel()
        {
            if (pnlCard == null) return;

            int x = (this.ClientSize.Width - pnlCard.Width) / 2;
            int y = (this.ClientSize.Height - pnlCard.Height) / 2;
            if (x < 0) x = 0;
            if (y < 0) y = 0;
            pnlCard.Location = new Point(x, y);
        }

        // ===========================================================
        // FORM LOAD
        // ===========================================================
        private void Form1_Load(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(lblFilePath.Text))
            {
                lblFilePath.Text = "No file selected";
            }
        }

        // ===========================================================
        // SELECT EXCEL BUTTON
        // ===========================================================
        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Select Excel File";
                ofd.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string ext = Path.GetExtension(ofd.FileName).ToLower();
                    if (ext != ".xlsx" && ext != ".xls")
                    {
                        MessageBox.Show("Only .xlsx and .xls files are allowed.",
                            "Invalid file", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        _selectedExcelPath = null;
                        lblFilePath.Text = "No file selected";
                        return;
                    }

                    _selectedExcelPath = ofd.FileName;
                    lblFilePath.Text = _selectedExcelPath;
                }
                else
                {
                    _selectedExcelPath = null;
                    lblFilePath.Text = "No file selected";
                }
            }
        }

        // ===========================================================
        // SELECT JSON BUTTON
        // ===========================================================
        private void BtnBrowseJson_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Title = "Select JSON File";
                ofd.Filter = "JSON Files (*.json)|*.json";

                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    string ext = Path.GetExtension(ofd.FileName).ToLower();
                    if (ext != ".json")
                    {
                        MessageBox.Show("Only .json files are allowed.",
                            "Invalid file", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                        _selectedJsonPath = null;
                        lblJsonPath.Text = "No file selected";
                        return;
                    }

                    _selectedJsonPath = ofd.FileName;
                    lblJsonPath.Text = _selectedJsonPath;
                }
                else
                {
                    _selectedJsonPath = null;
                    lblJsonPath.Text = "No file selected";
                }
            }
        }

        // ===========================================================
        // CONVERT EXCEL → JSON BUTTON
        // ===========================================================
        private void ConvertBtn_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_selectedExcelPath) || !File.Exists(_selectedExcelPath))
            {
                MessageBox.Show("Please select an Excel file first.",
                    "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ProgressForm progress = new ProgressForm();

            try
            {
                progress.Show();
                progress.UpdateStatus("Reading Excel file...");
                Application.DoEvents();

                // 1. Read + MERGE
                List<ProductMaster> products = ReadProductsFromExcel(_selectedExcelPath);

                if (products.Count == 0)
                {
                    MessageBox.Show("No data rows found in Excel.",
                        "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                progress.UpdateStatus("Converting to JSON...");
                Application.DoEvents();

                string json = JsonConvert.SerializeObject(products, Formatting.Indented);

                using (SaveFileDialog sfd = new SaveFileDialog())
                {
                    sfd.Title = "Save JSON File";
                    sfd.Filter = "JSON Files (*.json)|*.json";
                    sfd.FileName = Path.GetFileNameWithoutExtension(_selectedExcelPath) + ".json";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        progress.UpdateStatus("Saving JSON file...");
                        Application.DoEvents();

                        File.WriteAllText(sfd.FileName, json, Encoding.UTF8);

                        MessageBox.Show("JSON file created:\n" + sfd.FileName,
                            "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during conversion:\n" + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (!progress.IsDisposed)
                    progress.Close();
            }
        }

        // ===========================================================
        // CALL API BUTTON (UPLOAD JSON IN 1000-CHUNKS)
        // ===========================================================
        private async void BtnCallApi_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_selectedJsonPath) || !File.Exists(_selectedJsonPath))
            {
                MessageBox.Show("Please select a JSON file first.",
                    "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ProgressForm progress = new ProgressForm();

            try
            {
                progress.Show();
                progress.UpdateStatus("Reading JSON file...");
                Application.DoEvents();

                string jsonText = File.ReadAllText(_selectedJsonPath, Encoding.UTF8);

                var products = JsonConvert.DeserializeObject<List<ProductMaster>>(jsonText);

                if (products == null || products.Count == 0)
                {
                    MessageBox.Show("JSON file does not contain any products.",
                        "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                progress.UpdateStatus("Uploading to API in batches of 1000...");
                Application.DoEvents();

                await SendProductsToApiAsync(products, progress);

                MessageBox.Show("All batches uploaded successfully!",
                    "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error during API upload:\n" + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (!progress.IsDisposed)
                    progress.Close();
            }
        }

        // ===========================================================
        // EXCEL → LIST<ProductMaster>   (WITH MERGE STEP)
        // ===========================================================
        private List<ProductMaster> ReadProductsFromExcel(string excelPath)
        {
            var rawList = new List<ProductMaster>();

            int headerRowsToSkip = 4;   // skip first 4 rows
            int footerRowsToSkip = 3;   // skip last 3 rows (footer)

            using (var stream = File.Open(excelPath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream))
            {
                var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = false
                    }
                });

                if (dataSet.Tables.Count == 0)
                    return new List<ProductMaster>();

                DataTable table = dataSet.Tables[0];

                int totalRows = table.Rows.Count;
                int endRow = totalRows - footerRowsToSkip;

                for (int r = headerRowsToSkip; r < endRow; r++)
                {
                    DataRow row = table.Rows[r];

                    // Skip completely empty rows
                    if (row.ItemArray.All(x =>
                        x == null || x == DBNull.Value || string.IsNullOrWhiteSpace(x.ToString())))
                        continue;

                    var product = new ProductMaster
                    {
                        Code = GetString(row, 0),              // A
                        Name = GetString(row, 1),              // B
                        ShortName = GetString(row, 1),         // same as Name
                        MfgCompany = null,
                        MfgCode = CleanMfg(GetString(row, 17)),// R
                        GST = null,
                        MRP = GetDecimal(row, 10),             // K
                        SalesRate = GetDecimal(row, 12),       // M
                        PurchaseRate = GetDecimal(row, 11),    // L
                        Packing = null,
                        IsHide = null,
                        Stock = GetDecimal(row, 3),            // D
                        Discount = null,

                        // Scheme from E + F
                        Scheme = BuildScheme(row),

                        ExpDate = FormatDate(GetDate(row, 18)),// S
                        Generic = null
                    };

                    rawList.Add(product);
                }
            }

            // Merge duplicates (Code + Name + MfgCode)
            return MergeProducts(rawList);
        }

        // ===========================================================
        // MERGE LOGIC
        // ===========================================================
        private List<ProductMaster> MergeProducts(List<ProductMaster> input)
        {
            var result = new List<ProductMaster>();

            var groups = input.GroupBy(p => new
            {
                Code = NormalizeKey(p.Code),
                Name = NormalizeKey(p.Name),
                MfgCode = NormalizeKey(p.MfgCode)
            });

            foreach (var g in groups)
            {
                if (g.Count() == 1)
                {
                    result.Add(g.First());
                    continue;
                }

                // pick product with longest expiry; if tie, first one wins
                ProductMaster best = g.First();
                DateTime bestExp = ParseExpiry(best.ExpDate);

                foreach (var p in g.Skip(1))
                {
                    DateTime exp = ParseExpiry(p.ExpDate);
                    if (exp > bestExp)
                    {
                        best = p;
                        bestExp = exp;
                    }
                }

                // sum stock
                decimal totalStock = 0m;
                bool anyStock = false;
                foreach (var p in g)
                {
                    if (p.Stock.HasValue)
                    {
                        totalStock += p.Stock.Value;
                        anyStock = true;
                    }
                }

                var merged = new ProductMaster
                {
                    Code = best.Code,
                    Name = best.Name,
                    ShortName = best.ShortName,
                    MfgCompany = best.MfgCompany,
                    MfgCode = best.MfgCode,
                    GST = best.GST,
                    MRP = best.MRP,
                    SalesRate = best.SalesRate,
                    PurchaseRate = best.PurchaseRate,
                    Packing = best.Packing,
                    IsHide = best.IsHide,
                    Stock = anyStock ? (decimal?)totalStock : null,
                    Discount = best.Discount,
                    Scheme = best.Scheme,
                    ExpDate = best.ExpDate,
                    Generic = best.Generic
                };

                result.Add(merged);
            }

            return result;
        }

        private string NormalizeKey(string value)
        {
            return string.IsNullOrWhiteSpace(value)
                ? string.Empty
                : value.Trim().ToUpperInvariant();
        }

        private DateTime ParseExpiry(string expDate)
        {
            if (string.IsNullOrWhiteSpace(expDate))
                return DateTime.MinValue;

            if (DateTime.TryParse(expDate, out var dt))
                return dt;

            return DateTime.MinValue;
        }

        // ===========================================================
        // API UPLOAD (BATCHES OF 1000)
        // ===========================================================
        private async Task SendProductsToApiAsync(
            List<ProductMaster> products,
            ProgressForm progress)
        {
            const int batchSize = 1000;
            int total = products.Count;
            int totalBatches = (int)Math.Ceiling(total / (double)batchSize);

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Clear();
                client.DefaultRequestHeaders.Accept.Add(
                    new MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Add(NeoSoftCodeHeaderName, NeoSoftCodeHeaderValue);

                for (int i = 0; i < total; i += batchSize)
                {
                    int batchNumber = (i / batchSize) + 1;
                    var batch = products
                        .Skip(i)
                        .Take(batchSize)
                        .ToList();

                    progress.UpdateStatus(
                        $"Sending batch {batchNumber} of {totalBatches} ({batch.Count} items)...");
                    Application.DoEvents();

                    string bodyJson = JsonConvert.SerializeObject(batch);
                    using (var content = new StringContent(bodyJson, Encoding.UTF8, "application/json"))
                    {
                        HttpResponseMessage response = await client.PostAsync(ApiUrl, content);

                        if (!response.IsSuccessStatusCode)
                        {
                            string respText = await response.Content.ReadAsStringAsync();
                            throw new Exception(
                                $"API error on batch {batchNumber}/{totalBatches}: " +
                                $"{(int)response.StatusCode} {response.ReasonPhrase}\r\n{respText}");
                        }
                    }
                }
            }
        }

        // ===========================================================
        // HELPER METHODS
        // ===========================================================
        private string GetString(DataRow row, int col)
        {
            if (col < 0 || col >= row.Table.Columns.Count) return null;
            var v = row[col];
            return (v == null || v == DBNull.Value) ? null : v.ToString().Trim();
        }

        private decimal? GetDecimal(DataRow row, int col)
        {
            if (col < 0 || col >= row.Table.Columns.Count) return null;

            var v = row[col];
            if (v == null || v == DBNull.Value) return null;

            decimal d;
            if (decimal.TryParse(v.ToString(), out d)) return d;

            return null;
        }

        private DateTime? GetDate(DataRow row, int col)
        {
            if (col < 0 || col >= row.Table.Columns.Count) return null;

            var v = row[col];
            if (v == null || v == DBNull.Value) return null;

            DateTime d;
            if (DateTime.TryParse(v.ToString(), out d)) return d;

            return null;
        }

        private string FormatDate(DateTime? dt)
        {
            if (dt == null) return null;
            return dt.Value.ToString("yyyy-MM-dd");
        }

        private string CleanMfg(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return null;

            string v = value.Trim();

            // Normalize values like "-", "--", "- -", "-   -"
            if (v == "-" || v == "--" || v.Replace(" ", "") == "--")
                return null;

            return v;
        }

        /// <summary>
        /// Build scheme from column E + F:
        /// E = main qty (col 4), F = free qty (col 5)
        /// Special rule: if both numeric 0 → "0"
        /// </summary>
        private string BuildScheme(DataRow row)
        {
            string eVal = GetString(row, 4); // column E
            string fVal = GetString(row, 5); // column F

            bool eEmpty = string.IsNullOrWhiteSpace(eVal);
            bool fEmpty = string.IsNullOrWhiteSpace(fVal);

            // both empty → no scheme
            if (eEmpty && fEmpty)
                return null;

            // both numeric and 0 → "0"
            if (decimal.TryParse(eVal, out decimal eNum) &&
                decimal.TryParse(fVal, out decimal fNum))
            {
                if (eNum == 0 && fNum == 0)
                    return "0";
            }

            // both have values → "E+F"
            if (!eEmpty && !fEmpty)
                return $"{eVal}+{fVal}";

            // only one side has value
            return !eEmpty ? eVal : fVal;
        }

        private void lblFilePath_Click(object sender, EventArgs e)
        {
            // no action
        }
    }
}

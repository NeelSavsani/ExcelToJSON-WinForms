using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using ExcelDataReader;          // Excel reader
using Newtonsoft.Json;          // JSON

namespace SampleProject1
{
    public partial class Form1 : Form
    {
        private string _selectedExcelPath;

        // extra UI elements created in code
        private Panel pnlCard;
        private Label lblTitle;
        private Label lblSubtitle;
        private Label lblHint;

        public Form1()
        {
            InitializeComponent();
            SetupModernUi();      // build modern layout in code
        }

        // ===========================================================
        // MODERN UI LAYOUT
        // ===========================================================
        private void SetupModernUi()
        {
            // ----- Form look -----
            this.Text = "Excel to JSON Converter";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.Gainsboro;
            this.Font = new Font("Segoe UI", 10F);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;

            // ----- Main card panel -----
            pnlCard = new Panel();
            pnlCard.BackColor = Color.White;
            pnlCard.Size = new Size(700, 260);
            pnlCard.BorderStyle = BorderStyle.FixedSingle;
            CenterCardPanel();
            this.Controls.Add(pnlCard);

            // Re-center when form resizes
            this.Resize += (s, e) => CenterCardPanel();

            // ----- Title -----
            lblTitle = new Label();
            lblTitle.Text = "Excel → JSON Converter";
            lblTitle.Font = new Font("Segoe UI", 14F, FontStyle.Bold);
            lblTitle.Dock = DockStyle.Top;
            lblTitle.Height = 40;
            lblTitle.TextAlign = ContentAlignment.MiddleCenter;
            pnlCard.Controls.Add(lblTitle);

            // ----- Subtitle -----
            lblSubtitle = new Label();
            lblSubtitle.Text = "Select an Excel file and convert it into JSON product master data.";
            lblSubtitle.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            lblSubtitle.Dock = DockStyle.Top;
            lblSubtitle.Height = 30;
            lblSubtitle.TextAlign = ContentAlignment.MiddleCenter;
            lblSubtitle.ForeColor = Color.DimGray;
            pnlCard.Controls.Add(lblSubtitle);

            // Base Y coord for controls under subtitle
            int baseY = 90;

            // ----- File path label (acts like read-only textbox) -----
            lblFilePath.AutoSize = false;
            lblFilePath.BorderStyle = BorderStyle.FixedSingle;
            lblFilePath.BackColor = Color.White;
            lblFilePath.TextAlign = ContentAlignment.MiddleLeft;
            lblFilePath.Text = "No file selected";
            lblFilePath.Width = 480;
            lblFilePath.Height = 30;
            lblFilePath.Location = new Point(60, baseY);
            lblFilePath.Anchor = AnchorStyles.Top | AnchorStyles.Left;
            pnlCard.Controls.Add(lblFilePath); // move into panel

            // ----- Browse button -----
            button1.Text = "Browse";
            button1.Width = 110;
            button1.Height = 30;
            button1.Location = new Point(lblFilePath.Right + 15, baseY);
            button1.Anchor = AnchorStyles.Top | AnchorStyles.Right;

            button1.FlatStyle = FlatStyle.Flat;
            button1.FlatAppearance.BorderSize = 0;
            button1.BackColor = Color.DodgerBlue;
            button1.ForeColor = Color.White;
            button1.Font = new Font("Segoe UI", 9F, FontStyle.Bold); // smaller font

            pnlCard.Controls.Add(button1); // move into panel

            // ----- Hint text -----
            lblHint = new Label();
            lblHint.AutoSize = true;
            lblHint.Text = "Supported formats: .xlsx, .xls";
            lblHint.Font = new Font("Segoe UI", 8.5F);
            lblHint.ForeColor = Color.Gray;
            lblHint.Location = new Point(lblFilePath.Left + 4, baseY + 38);
            pnlCard.Controls.Add(lblHint);

            // ----- Convert button -----
            ConvertBtn.Text = "Convert to JSON";
            ConvertBtn.Width = 200;
            ConvertBtn.Height = 40;
            ConvertBtn.FlatStyle = FlatStyle.Flat;
            ConvertBtn.FlatAppearance.BorderSize = 0;
            ConvertBtn.BackColor = Color.MediumSeaGreen;
            ConvertBtn.ForeColor = Color.White;
            ConvertBtn.Font = new Font("Segoe UI", 10F, FontStyle.Bold);

            int convertX = (pnlCard.Width - ConvertBtn.Width) / 2;
            int convertY = baseY + 80;
            ConvertBtn.Location = new Point(convertX, convertY);
            ConvertBtn.Anchor = AnchorStyles.Top;
            pnlCard.Controls.Add(ConvertBtn); // move into panel
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
        // SELECT FILE BUTTON
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
        // CONVERT BUTTON
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

                // 1. Read + merge
                List<ProductMaster> products = ReadProductsFromExcel(_selectedExcelPath);

                if (products.Count == 0)
                {
                    MessageBox.Show("No data rows found in Excel.",
                        "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                progress.UpdateStatus("Converting to JSON...");
                Application.DoEvents();

                // 2. JSON
                string json = JsonConvert.SerializeObject(products, Formatting.Indented);

                // 3. Save
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
        // EXCEL → LIST<ProductMaster>   (NOW WITH MERGE STEP)
        // ===========================================================
        private List<ProductMaster> ReadProductsFromExcel(string excelPath)
        {
            var rawList = new List<ProductMaster>();

            int headerRowsToSkip = 4;   // skip first 4 rows
            int footerRowsToSkip = 1;   // skip last row (change if needed)

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
                        Code = GetString(row, 0),
                        Name = GetString(row, 1),
                        ShortName = GetString(row, 1),
                        MfgCompany = null,
                        MfgCode = CleanMfg(GetString(row, 17)),
                        GST = null,
                        MRP = GetDecimal(row, 10),
                        SalesRate = GetDecimal(row, 12),
                        PurchaseRate = GetDecimal(row, 11),
                        Packing = null,
                        IsHide = null,
                        Stock = GetDecimal(row, 3),
                        Discount = null,

                        // Scheme built from E + F (4+5)
                        Scheme = BuildScheme(row),

                        ExpDate = FormatDate(GetDate(row, 18)),
                        Generic = null
                    };

                    rawList.Add(product);
                }
            }

            // 🔁 MERGE DUPLICATES: same Code + Name + MfgCode
            return MergeProducts(rawList);
        }

        // ===========================================================
        // MERGE LOGIC
        // ===========================================================
        /// <summary>
        /// Merge products that are "same" (Code+Name+MfgCode).
        /// - Stock: sum of all stocks
        /// - Expiry: take product with longest expiry
        /// - Other fields: from longest-expiry product
        /// - If multiple have same longest expiry: first one from Excel wins
        /// </summary>
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
                // If only one product in group – no merge needed
                if (g.Count() == 1)
                {
                    result.Add(g.First());
                    continue;
                }

                // 1) Find "best" product = max expiry, breaking ties by order
                ProductMaster best = g.First();                   // start with first row
                DateTime bestExpDate = ParseExpiry(best.ExpDate); // MinValue if null/invalid

                foreach (var p in g.Skip(1))
                {
                    DateTime exp = ParseExpiry(p.ExpDate);

                    // strictly greater → take this as best
                    if (exp > bestExpDate)
                    {
                        best = p;
                        bestExpDate = exp;
                    }
                    // if equal, keep existing 'best' to follow "first product wins"
                }

                // 2) Sum stock across all products
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

                // Clone best so we don't modify original collection accidentally
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

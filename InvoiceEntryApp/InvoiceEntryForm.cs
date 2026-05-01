using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace InvoiceEntryApp
{
    public sealed class InvoiceEntryForm : Form
    {
        private readonly ExcelInvoiceSaver excelInvoiceSaver = new ExcelInvoiceSaver();

        private Panel pnlHeader;
        private Panel pnlFooter;
        private Label lblTitle;
        private GroupBox grpCompany;
        private GroupBox grpCustomer;
        private GroupBox grpInvoiceDetails;
        private Label lblItems;
        private DataGridView dgvItems;
        private Label lblSubTotal;
        private Label lblGst;
        private Label lblTotal;
        private TextBox txtSubTotal;
        private TextBox txtGst;
        private TextBox txtTotal;
        private Button btnSubmit;
        private Button btnCancel;

        private TextBox txtCompanyName;
        private TextBox txtCompanyAddress;
        private TextBox txtCompanyCity;
        private TextBox txtCompanyState;
        private TextBox txtCompanyPinCode;
        private TextBox txtCompanyContactNo;
        private TextBox txtCompanyTin;

        private TextBox txtCustomerName;
        private TextBox txtCustomerAddress;
        private TextBox txtCustomerCity;
        private TextBox txtCustomerState;
        private TextBox txtCustomerPinCode;
        private TextBox txtCustomerContactNo;
        private TextBox txtCustomerTin;

        private TextBox txtInvoiceNo;
        private TextBox txtInvoiceDate;

        public InvoiceEntryForm()
        {
            InitializeComponent();
            ConfigureGrid();
            ResetForm();
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            AutoScaleDimensions = new SizeF(6F, 13F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = Color.FromArgb(245, 245, 245);
            ClientSize = new Size(960, 650);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "InvoiceEntryForm";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Invoice Entry";

            pnlHeader = new Panel
            {
                Name = "pnlHeader",
                BackColor = Color.White,
                Location = new Point(0, 0),
                Size = new Size(960, 78)
            };

            pnlFooter = new Panel
            {
                Name = "pnlFooter",
                BackColor = Color.Black,
                Dock = DockStyle.Bottom,
                Height = 12
            };

            lblTitle = new Label
            {
                Name = "lblTitle",
                Font = new Font("Segoe UI", 24F, FontStyle.Regular, GraphicsUnit.Point),
                Location = new Point(280, 16),
                Size = new Size(400, 42),
                Text = "Invoice Entry",
                TextAlign = ContentAlignment.MiddleCenter
            };
            pnlHeader.Controls.Add(lblTitle);

            grpCompany = new GroupBox
            {
                Name = "grpCompany",
                Text = "Company",
                Location = new Point(32, 88),
                Size = new Size(400, 214),
                TabIndex = 0,
                TabStop = false
            };

            grpCustomer = new GroupBox
            {
                Name = "grpCustomer",
                Text = "Customer",
                Location = new Point(528, 88),
                Size = new Size(400, 214),
                TabIndex = 1,
                TabStop = false
            };

            grpInvoiceDetails = new GroupBox
            {
                Name = "grpInvoiceDetails",
                Text = "Invoice Details",
                Location = new Point(32, 312),
                Size = new Size(896, 52),
                TabIndex = 2,
                TabStop = false
            };

            txtCompanyName = AddLabeledTextBox(grpCompany, "lblCompanyName", "Name", "txtCompanyName", 78, 28, 1, 95, 155);
            txtCompanyAddress = AddLabeledTextBox(grpCompany, "lblCompanyAddress", "Address", "txtCompanyAddress", 78, 54, 2, 95, 155);
            txtCompanyCity = AddLabeledTextBox(grpCompany, "lblCompanyCity", "City", "txtCompanyCity", 78, 80, 3, 95, 155);
            txtCompanyState = AddLabeledTextBox(grpCompany, "lblCompanyState", "State", "txtCompanyState", 78, 106, 4, 95, 155);
            txtCompanyPinCode = AddLabeledTextBox(grpCompany, "lblCompanyPinCode", "Pin Code", "txtCompanyPinCode", 78, 132, 5, 95, 155);
            txtCompanyContactNo = AddLabeledTextBox(grpCompany, "lblCompanyContactNo", "Contact No", "txtCompanyContactNo", 78, 158, 6, 95, 155);
            txtCompanyTin = AddLabeledTextBox(grpCompany, "lblCompanyTin", "Tin", "txtCompanyTin", 78, 184, 7, 95, 155);

            txtCustomerName = AddLabeledTextBox(grpCustomer, "lblCustomerName", "Customer Name", "txtCustomerName", 42, 28, 8, 120, 165);
            txtCustomerAddress = AddLabeledTextBox(grpCustomer, "lblCustomerAddress", "Customer Address", "txtCustomerAddress", 42, 54, 9, 120, 165);
            txtCustomerCity = AddLabeledTextBox(grpCustomer, "lblCustomerCity", "City", "txtCustomerCity", 42, 80, 10, 120, 165);
            txtCustomerState = AddLabeledTextBox(grpCustomer, "lblCustomerState", "State", "txtCustomerState", 42, 106, 11, 120, 165);
            txtCustomerPinCode = AddLabeledTextBox(grpCustomer, "lblCustomerPinCode", "Pin Code", "txtCustomerPinCode", 42, 132, 12, 120, 165);
            txtCustomerContactNo = AddLabeledTextBox(grpCustomer, "lblCustomerContactNo", "Contact No", "txtCustomerContactNo", 42, 158, 13, 120, 165);
            txtCustomerTin = AddLabeledTextBox(grpCustomer, "lblCustomerTin", "Tin", "txtCustomerTin", 42, 184, 14, 120, 165);

            Label lblInvoiceNo = CreateLabel("lblInvoiceNo", "Invoice No", new Point(22, 22), new Size(100, 20));
            txtInvoiceNo = CreateTextBox("txtInvoiceNo", new Point(132, 21), 15, 155);
            Label lblInvoiceDate = CreateLabel("lblInvoiceDate", "Invoice Date", new Point(530, 22), new Size(100, 20));
            txtInvoiceDate = CreateTextBox("txtInvoiceDate", new Point(640, 21), 16, 155);
            grpInvoiceDetails.Controls.Add(lblInvoiceNo);
            grpInvoiceDetails.Controls.Add(txtInvoiceNo);
            grpInvoiceDetails.Controls.Add(lblInvoiceDate);
            grpInvoiceDetails.Controls.Add(txtInvoiceDate);

            lblItems = new Label
            {
                Name = "lblItems",
                AutoSize = true,
                Font = new Font("Segoe UI", 18F, FontStyle.Regular, GraphicsUnit.Point),
                Location = new Point(34, 372),
                Text = "Items"
            };

            dgvItems = new DataGridView
            {
                Name = "dgvItems",
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle,
                ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
                Location = new Point(34, 405),
                Size = new Size(894, 96),
                TabIndex = 17,
                RowTemplate = { Height = 24 },
                MultiSelect = false,
                SelectionMode = DataGridViewSelectionMode.CellSelect,
                AllowUserToOrderColumns = false,
                AllowUserToResizeColumns = false,
                EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2,
                GridColor = Color.Silver,
                CellBorderStyle = DataGridViewCellBorderStyle.Single,
                RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single
            };
            dgvItems.CellEndEdit += dgvItems_CellEndEdit;
            dgvItems.RowsRemoved += dgvItems_RowsRemoved;
            dgvItems.UserDeletedRow += dgvItems_UserDeletedRow;

            lblSubTotal = new Label
            {
                Name = "lblSubTotal",
                Location = new Point(690, 518),
                Size = new Size(88, 23),
                Text = "SubTotal",
                TextAlign = ContentAlignment.MiddleRight
            };

            txtSubTotal = new TextBox
            {
                Name = "txtSubTotal",
                BackColor = Color.White,
                Location = new Point(784, 519),
                Size = new Size(144, 20),
                TabIndex = 18
            };

            lblGst = new Label
            {
                Name = "lblGst",
                Location = new Point(690, 545),
                Size = new Size(88, 23),
                Text = "GST",
                TextAlign = ContentAlignment.MiddleRight
            };

            txtGst = new TextBox
            {
                Name = "txtGst",
                Location = new Point(784, 546),
                Size = new Size(144, 20),
                TabIndex = 19
            };
            txtGst.TextChanged += txtGst_TextChanged;

            lblTotal = new Label
            {
                Name = "lblTotal",
                Location = new Point(690, 572),
                Size = new Size(88, 23),
                Text = "Total",
                TextAlign = ContentAlignment.MiddleRight
            };

            txtTotal = new TextBox
            {
                Name = "txtTotal",
                BackColor = Color.White,
                Location = new Point(784, 573),
                Size = new Size(144, 20),
                TabIndex = 20
            };

            btnSubmit = new Button
            {
                Name = "btnSubmit",
                Location = new Point(690, 604),
                Size = new Size(110, 31),
                TabIndex = 21,
                Text = "Submit",
                UseVisualStyleBackColor = true
            };
            btnSubmit.Click += btnSubmit_Click;

            btnCancel = new Button
            {
                Name = "btnCancel",
                Location = new Point(818, 604),
                Size = new Size(110, 31),
                TabIndex = 22,
                Text = "Cancel",
                UseVisualStyleBackColor = true
            };
            btnCancel.Click += btnCancel_Click;

            Controls.Add(btnCancel);
            Controls.Add(btnSubmit);
            Controls.Add(txtTotal);
            Controls.Add(txtGst);
            Controls.Add(txtSubTotal);
            Controls.Add(lblTotal);
            Controls.Add(lblGst);
            Controls.Add(lblSubTotal);
            Controls.Add(dgvItems);
            Controls.Add(lblItems);
            Controls.Add(grpInvoiceDetails);
            Controls.Add(grpCustomer);
            Controls.Add(grpCompany);
            Controls.Add(pnlFooter);
            Controls.Add(pnlHeader);

            ResumeLayout(false);
            PerformLayout();
        }

        private TextBox AddLabeledTextBox(Control parent, string labelName, string labelText, string textBoxName, int labelX, int textBoxY, int tabIndex, int labelWidth, int textBoxWidth)
        {
            Label label = CreateLabel(labelName, labelText, new Point(labelX, textBoxY - 1), new Size(labelWidth, 20));
            TextBox textBox = CreateTextBox(textBoxName, new Point(labelX + labelWidth + 14, textBoxY), tabIndex, textBoxWidth);
            parent.Controls.Add(label);
            parent.Controls.Add(textBox);
            return textBox;
        }

        private static Label CreateLabel(string name, string text, Point location, Size size)
        {
            return new Label
            {
                Name = name,
                Location = location,
                Size = size,
                Text = text,
                TextAlign = ContentAlignment.MiddleRight
            };
        }

        private static TextBox CreateTextBox(string name, Point location, int tabIndex, int width)
        {
            return new TextBox
            {
                Name = name,
                Location = location,
                Size = new Size(width, 20),
                TabIndex = tabIndex
            };
        }

        private void ConfigureGrid()
        {
            dgvItems.AutoGenerateColumns = false;
            dgvItems.AllowUserToAddRows = true;
            dgvItems.AllowUserToDeleteRows = true;
            dgvItems.AllowUserToResizeRows = false;
            dgvItems.RowHeadersWidth = 24;
            dgvItems.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            dgvItems.Columns.Clear();
            dgvItems.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            dgvItems.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colItemNo",
                HeaderText = "Item No",
                FillWeight = 24,
                MinimumWidth = 140,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });

            dgvItems.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colDescription",
                HeaderText = "Description",
                FillWeight = 28,
                MinimumWidth = 190,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });

            dgvItems.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colQuantity",
                HeaderText = "Quantity",
                FillWeight = 24,
                MinimumWidth = 140,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });

            dgvItems.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "colPrice",
                HeaderText = "Price",
                FillWeight = 24,
                MinimumWidth = 140,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            string validationMessage = TryBuildSubmission(out InvoiceSubmission submission);
            if (!string.IsNullOrWhiteSpace(validationMessage))
            {
                MessageBox.Show(this, validationMessage, "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                excelInvoiceSaver.Save(submission);
                MessageBox.Show(this, "Invoice saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                ResetForm();
            }
            catch (System.IO.IOException)
            {
                MessageBox.Show(this, "The Excel file is currently unavailable or locked. Close the file if it is open and try again.", "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (UnauthorizedAccessException)
            {
                MessageBox.Show(this, "The application could not write to the configured Excel path. Check the folder permissions and try again.", "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Unable to save the invoice to Excel.\r\n" + ex.Message, "Save Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            ResetForm();
        }

        private void dgvItems_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            RecalculateTotals();
        }

        private void dgvItems_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            RecalculateTotals();
        }

        private void dgvItems_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            RecalculateTotals();
        }

        private void txtGst_TextChanged(object sender, EventArgs e)
        {
            RecalculateTotals();
        }

        private string TryBuildSubmission(out InvoiceSubmission submission)
        {
            submission = null;

            if (!TryParseDate(txtInvoiceDate.Text, out DateTime invoiceDate))
            {
                return "Enter a valid Invoice Date.";
            }

            if (!TryParseDecimal(txtGst.Text, out decimal gst))
            {
                return "GST must be numeric.";
            }

            List<InvoiceItem> items = ReadItems();
            decimal subTotalValue = items.Sum(item => item.LineTotal);
            decimal totalValue = subTotalValue + gst;

            InvoiceSubmission candidate = new InvoiceSubmission
            {
                Company = new PartyInfo
                {
                    Name = txtCompanyName.Text.Trim(),
                    Address = txtCompanyAddress.Text.Trim(),
                    City = txtCompanyCity.Text.Trim(),
                    State = txtCompanyState.Text.Trim(),
                    PinCode = txtCompanyPinCode.Text.Trim(),
                    ContactNo = txtCompanyContactNo.Text.Trim(),
                    Tin = txtCompanyTin.Text.Trim()
                },
                Customer = new PartyInfo
                {
                    Name = txtCustomerName.Text.Trim(),
                    Address = txtCustomerAddress.Text.Trim(),
                    City = txtCustomerCity.Text.Trim(),
                    State = txtCustomerState.Text.Trim(),
                    PinCode = txtCustomerPinCode.Text.Trim(),
                    ContactNo = txtCustomerContactNo.Text.Trim(),
                    Tin = txtCustomerTin.Text.Trim()
                },
                InvoiceNo = txtInvoiceNo.Text.Trim(),
                InvoiceDate = invoiceDate,
                SubTotal = subTotalValue.ToString("0.00"),
                Gst = txtGst.Text.Trim(),
                Total = totalValue.ToString("0.00")
            };

            candidate.Items.AddRange(items);

            string validationMessage = InvoiceValidator.Validate(candidate);
            if (!string.IsNullOrWhiteSpace(validationMessage))
            {
                return validationMessage;
            }

            submission = candidate;
            return string.Empty;
        }

        private List<InvoiceItem> ReadItems()
        {
            List<InvoiceItem> items = new List<InvoiceItem>();

            foreach (DataGridViewRow row in dgvItems.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

                string itemNo = Convert.ToString(row.Cells["colItemNo"].Value)?.Trim();
                string description = Convert.ToString(row.Cells["colDescription"].Value)?.Trim();
                string quantityText = Convert.ToString(row.Cells["colQuantity"].Value)?.Trim();
                string priceText = Convert.ToString(row.Cells["colPrice"].Value)?.Trim();

                bool rowHasAnyValue =
                    !string.IsNullOrWhiteSpace(itemNo) ||
                    !string.IsNullOrWhiteSpace(description) ||
                    !string.IsNullOrWhiteSpace(quantityText) ||
                    !string.IsNullOrWhiteSpace(priceText);

                if (!rowHasAnyValue)
                {
                    continue;
                }

                if (!TryParseDecimal(quantityText, out decimal quantity))
                {
                    quantity = 0;
                }

                if (!TryParseDecimal(priceText, out decimal price))
                {
                    price = -1;
                }

                items.Add(new InvoiceItem
                {
                    ItemNo = itemNo,
                    Description = description,
                    Quantity = quantity,
                    Price = price
                });
            }

            return items;
        }

        private void RecalculateTotals()
        {
            decimal subTotal = 0m;

            foreach (InvoiceItem item in ReadItems())
            {
                if (item.Quantity > 0 && item.Price >= 0)
                {
                    subTotal += item.LineTotal;
                }
            }

            txtSubTotal.Text = subTotal == 0m ? string.Empty : subTotal.ToString("0.00");

            decimal gst = 0m;
            if (!string.IsNullOrWhiteSpace(txtGst.Text) && TryParseDecimal(txtGst.Text, out decimal parsedGst) && parsedGst >= 0)
            {
                gst = parsedGst;
            }

            decimal total = subTotal + gst;
            txtTotal.Text = total == 0m ? string.Empty : total.ToString("0.00");
        }

        private void ResetForm()
        {
            txtCompanyName.Text = string.Empty;
            txtCompanyAddress.Text = string.Empty;
            txtCompanyCity.Text = string.Empty;
            txtCompanyState.Text = string.Empty;
            txtCompanyPinCode.Text = string.Empty;
            txtCompanyContactNo.Text = string.Empty;
            txtCompanyTin.Text = string.Empty;
            txtCustomerName.Text = string.Empty;
            txtCustomerAddress.Text = string.Empty;
            txtCustomerCity.Text = string.Empty;
            txtCustomerState.Text = string.Empty;
            txtCustomerPinCode.Text = string.Empty;
            txtCustomerContactNo.Text = string.Empty;
            txtCustomerTin.Text = string.Empty;
            txtInvoiceNo.Text = string.Empty;
            txtInvoiceDate.Text = string.Empty;
            txtSubTotal.Text = string.Empty;
            txtGst.Text = string.Empty;
            txtTotal.Text = string.Empty;
            dgvItems.Rows.Clear();
            txtCompanyName.Focus();
        }

        private static bool TryParseDecimal(string value, out decimal parsedValue)
        {
            return decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out parsedValue) ||
                   decimal.TryParse(value, NumberStyles.Number, CultureInfo.CurrentCulture, out parsedValue);
        }

        private static bool TryParseDate(string value, out DateTime parsedDate)
        {
            string[] formats =
            {
                "dd/MM/yyyy",
                "d/M/yyyy",
                "dd-MM-yyyy",
                "d-M-yyyy",
                "yyyy-MM-dd",
                "MM/dd/yyyy",
                "M/d/yyyy"
            };

            return DateTime.TryParseExact(value, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate) ||
                   DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.None, out parsedDate);
        }
    }
}

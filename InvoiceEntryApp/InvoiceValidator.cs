namespace InvoiceEntryApp
{
    internal static class InvoiceValidator
    {
        public static string Validate(InvoiceSubmission submission)
        {
            if (submission == null)
            {
                return "Invoice data is missing.";
            }

            string companyError = ValidateParty(submission.Company, "Company");
            if (!string.IsNullOrWhiteSpace(companyError))
            {
                return companyError;
            }

            string customerError = ValidateParty(submission.Customer, "Customer");
            if (!string.IsNullOrWhiteSpace(customerError))
            {
                return customerError;
            }

            if (string.IsNullOrWhiteSpace(submission.InvoiceNo))
            {
                return "Invoice No is required.";
            }

            if (submission.InvoiceDate == System.DateTime.MinValue)
            {
                return "Invoice Date is required.";
            }

            if (submission.Items.Count == 0)
            {
                return "Enter at least one valid item.";
            }

            foreach (InvoiceItem item in submission.Items)
            {
                if (string.IsNullOrWhiteSpace(item.ItemNo) || string.IsNullOrWhiteSpace(item.Description))
                {
                    return "Each item must include Item No and Description.";
                }

                if (item.Quantity <= 0)
                {
                    return "Quantity must be greater than zero.";
                }

                if (item.Price < 0)
                {
                    return "Price must be zero or greater.";
                }
            }

            return string.Empty;
        }

        private static string ValidateParty(PartyInfo party, string sectionName)
        {
            if (party == null)
            {
                return sectionName + " details are required.";
            }

            if (string.IsNullOrWhiteSpace(party.Name)) return sectionName + " Name is required.";
            if (string.IsNullOrWhiteSpace(party.Address)) return sectionName + " Address is required.";
            if (string.IsNullOrWhiteSpace(party.City)) return sectionName + " City is required.";
            if (string.IsNullOrWhiteSpace(party.State)) return sectionName + " State is required.";
            if (string.IsNullOrWhiteSpace(party.PinCode)) return sectionName + " Pin Code is required.";
            if (string.IsNullOrWhiteSpace(party.ContactNo)) return sectionName + " Contact No is required.";
            if (string.IsNullOrWhiteSpace(party.Tin)) return sectionName + " Tin is required.";

            return string.Empty;
        }
    }
}

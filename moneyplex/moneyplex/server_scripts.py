import frappe, pandas, openpyxl
from io import BytesIO

@frappe.whitelist()
def moneyplex_export():
    query = """
    SELECT
        `tabPayment Request`.name,
        `tabPayment Request`.company AS "Kontoname",
        `tabPayment Request`.bank_account_no AS "Kontonummer",
        `tabPayment Request`.branch_code AS "Bankleitzahl",
    --    ursprüngliche Variante:
    --    `tabPayment Request`.transaction_date AS "Datum",
    --    `tabPayment Request`.valuta AS "Valuta",
        DATE_FORMAT(`tabPayment Request`.transaction_date, '%d.%m.%Y') AS "Datum",
        DATE_FORMAT(`tabPayment Request`.valuta, '%d.%m.%Y') AS "Valuta",
        `tabPayment Request`.supplier_name AS "Name",
        `tabPayment Request`.party_iban AS "Iban",
        `tabPayment Request`.party_bic AS "Bic",
        `tabPayment Request`.reference_name AS "Zweck",
        `tabPayment Request`.mode_of_payment AS "Kategorie",
    -- ursprüngliche Variante:
    --    `tabPayment Request`.grand_total AS "Betrag",
        FORMAT(`tabPayment Request`.grand_total, 2, 'de_DE') AS "Betrag",
        `tabPayment Request`.currency AS "Währung"
    FROM
        `tabPayment Request`
    WHERE
        `tabPayment Request`.docstatus = 1 AND `tabPayment Request`.für_moneyplex_exportiert = 0
    ORDER BY
        `tabPayment Request`.creation
    """
    result = frappe.db.sql(query, as_dict=True)

    if len(result) == 0:
        frappe.msgprint("No records to exports.")
        return ""

    # save the .xlsx file
    filename = "moneyplex_export.xlsx"
    filepath = "{}/private/files/{}".format(frappe.utils.get_site_path(), filename)
    df = pandas.DataFrame.from_records(result)
    names = list(df["name"])
    df.drop("name", axis=1, inplace=True)
    df.to_excel(filepath, index=False)

    payment_entries = []
    for name in names:
        # create payment entries
        doc = frappe.get_doc("Payment Request", name)
        pe = doc.create_payment_entry(submit=True)
        payment_entries.append(pe.name)
        # set für_moneyplex_exportiert to 1
        frappe.db.set_value("Payment Entry", pe.name, "für_moneyplex_exportiert", 1)
        frappe.db.set_value("Payment Request", name, "für_moneyplex_exportiert", 1)
    
    if payment_entries:
        message = "Payment Entries generated: "
        for pe in payment_entries:
            message += "<a href=/app/payment-entry/{} target='_blank'>{}</a>, ".format(pe, pe)
        frappe.msgprint(message)

    files = frappe.db.get_list("File", filters={"file_url": "/private/files/{}".format(filename)})
    if len(files) == 0:
        doc = frappe.get_doc({
            "doctype": "File", 
            "file_name": filename, 
            "is_private": 1, 
            "file_url": "/private/files/{}".format(filename)
            })
        doc.insert()

    # build the full URL for the file
    full_file_url = frappe.utils.get_site_url(frappe.utils.get_site_base_path()[2:])
    full_file_url = "{}/private/files/{}".format(full_file_url, filename)

    return full_file_url


from erpnext.accounts.doctype.payment_request.payment_request import PaymentRequest

class CustomPaymentRequest(PaymentRequest):
    def on_submit(self):
        """
        this is a copy of the on_submit function from the Payment Request
        with the difference that this function avoids the sending of emails
        """
        
        if self.payment_request_type == "Outward":
            self.db_set("status", "Initiated")
            return
        elif self.payment_request_type == "Inward":
            self.db_set("status", "Requested")

        send_mail = False # force the avoidance of the email to be sent
        ref_self = frappe.get_self(self.reference_selftype, self.reference_name)

        if (
            hasattr(ref_self, "order_type") and getattr(ref_self, "order_type") == "Shopping Cart"
        ) or self.flags.mute_email:
            send_mail = False

        if send_mail and self.payment_channel != "Phone":
            self.set_payment_request_url()
            self.send_email()
            self.make_communication_entry()

        elif self.payment_channel == "Phone":
            self.request_phone_payment()
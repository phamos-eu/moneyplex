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
        TO_CHAR(`tabPayment Request`.transaction_date, 'DD.MM.YYYY') AS "Datum",
        TO_CHAR(`tabPayment Request`.valuta, 'DD.MM.YYYY') AS "Valuta",
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
    """
    result = frappe.db.sql(query, as_dict=True)

    # save the .xlsx file
    filename = "moneyplex_export.xlsx"
    filepath = "{}/private/files/{}".format(frappe.utils.get_site_path(), filename)
    df = pandas.DataFrame.from_records(result, index=['1', '2'])
    names = list(df["name"])
    df.drop("name", axis=1, inplace=True)
    df.to_excel(filepath, index=False)

    # set für_moneyplex_exportiert to 1
    for name in names:
        frappe.db.set_value("Payment Request", name, "für_moneyplex_exportiert", 1)

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
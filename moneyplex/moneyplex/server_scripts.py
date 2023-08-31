import frappe

@frappe.whitelist()
def moneyplex_export():
    prs = frappe.db.get_list("Payment Request", filters={"docstatus": 1})
    return prs
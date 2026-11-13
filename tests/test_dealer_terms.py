from scripts.dealer_terms import (
    find_customer_forbidden_in_text,
    is_builder_contact_text,
    is_dealer_terms_text,
)


def test_drops_builder_contact_block():
    t = "Nathan Miller Family 6470 Township Road 603 PHONE: 330-279-2102 FAX: 844-415-1713"
    assert is_builder_contact_text(t)
    assert is_dealer_terms_text(t)


def test_drops_finance_charge_terms():
    t = "A 2% finance charge may be added to any invoice not paid within terms."
    assert is_dealer_terms_text(t)
    assert "finance charge" in find_customer_forbidden_in_text(t)


def test_drops_past_due_invoice_terms():
    t = "Invoices more than 30 days past due are subject to collection."
    assert is_dealer_terms_text(t)

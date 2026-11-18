from bc import rules

CUST = {"CN10001": {"name": "Shree Ganesh Traders"}, "CN10002": {"name": "Patel Auto Spares"},
        "CW10021": {"name": "Kaveri Lubricants Pvt Ltd"}}


def test_amounts_indian_format_and_dates_ignored():
    a = rules.find_amounts("Balance as on 30.09.2026 is Rs 4,82,350.00 Dr, ours Rs 12,400.00 in 2026")
    assert [x["value"] for x in a] == [482350.0, 12400.0]
    assert a[0]["side"] == "DR"


def test_any_difference_is_dispute_even_paise():
    assert rules.status_for(1000.00, 1000.00) == "CLEAR"
    assert rules.status_for(1000.00, 999.99) == "DISPUTE"
    assert rules.difference(482350, 412350) == 70000.0


def test_no_changes_is_clear_with_our_balance():
    s = rules.suggest("Dear Sir, no changes from our side.", 126900)
    assert s["status"] == "CLEAR" and s["theirs"] == 126900


def test_caveat_blocks_automatic_clear():
    s = rules.suggest("We confirm the balance except bill no. 4471 for Rs 12,400.00", 50000)
    assert s["status"] is None and s["theirs"] is None and s["needs_typing"]
    assert s["flags"]


def test_different_amount_suggests_dispute():
    s = rules.suggest("As per our books the balance is Rs 4,82,350.00. Balance as per our books Rs 4,12,350.00", 482350)
    assert s["theirs"] == 412350.0 and s["status"] == "DISPUTE"


def test_empty_text_needs_typing():
    s = rules.suggest("", 1000)
    assert s["needs_typing"] and s["status"] is None


def test_customer_by_code_with_name_check():
    no, how, ok, _ = rules.find_customer("", "CN10001 Shree Ganesh Traders.pdf", CUST)
    assert no == "CN10001" and ok


def test_code_name_mismatch_is_flagged():
    no, how, ok, _ = rules.find_customer("Customer: Patel Auto Spares", "CN10001 reply.pdf", CUST)
    assert no == "CN10001" and not ok


def test_name_only_fallback():
    no, how, ok, _ = rules.find_customer("From M/s Kaveri Lubricants Private Limited", "confirmation.pdf", CUST)
    assert no == "CW10021" and how.startswith("matched by name")


def test_code_not_matched_inside_longer_code():
    no, *_ = rules.find_customer("ref CN100011", "x.pdf", CUST)
    assert no is None


def test_disagree_with_amount_is_dispute_not_blocked():
    s = rules.suggest("[x] We do not agree. Balance as per our books Rs 4,12,350.00 Dr", 482350)
    assert s["theirs"] == 412350.0 and s["status"] == "DISPUTE"


def test_disagree_without_amount_needs_typing():
    s = rules.suggest("We do not agree with the balance.", 482350)
    assert s["status"] is None and s["needs_typing"]


def test_codes_and_bill_numbers_are_not_amounts():
    a = rules.find_amounts("Customer No. CN10002, bill no. 4471, Invoice No: 99812, balance Rs4,000.00 and Rs 12,400")
    assert [x["value"] for x in a] == [4000.0, 12400.0]

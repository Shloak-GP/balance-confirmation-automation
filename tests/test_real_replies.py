"""Wording patterns from real replies. Customer names, e-mails and phone numbers are replaced with made-up ones."""
from bc import rules

AMBER = "Kindly check.\n\nFrom: X <a@b.in>\nSubject: Re: Customer Balance Confirmation\nAs Per Book Of Account As On Date 30-09-2026 Balance Amount Is Rs. 1,67,75,671/- Credited\n\nFrom: GP\nold Rs. 99,999.00\n360 004 <https://linkprotect.example/url?a=10006&c=E,1,b9>"
NIL = "From: Y\nBalance 01.04.26 to 30.09.26. Total amount-67,38,187.50."
ZEN = "From: Z\nPlease find below the balance as per our books.\n30.09.2026 :--   -3,452,112.00\nM : 9000000001\n"


def test_indian_amount_with_slash_dash_and_cr():
    s = rules.suggest(AMBER, 16775671)
    assert s["status"] == "CLEAR" and s["theirs"] == 16775671


def test_period_total_is_not_taken_as_closing_balance():
    s = rules.suggest(NIL, 18351801)
    assert s["needs_typing"] and s["flags"]


def test_negative_western_grouping_and_phone_ignored():
    s = rules.suggest(ZEN, 3438209)
    assert s["theirs"] == 3452112.0 and s["status"] == "DISPUTE"


CUSTS = {
    "C1": dict(name="Amber Trading Co. (IPOL)", email="ambertrading@example.in"),
    "C2": dict(name="Amber Chemicals Pvt. Ltd.", email=""),
    "C3": dict(name="Zentora India Private Limited", email="accounts@zentora.example"),
    "C4": dict(name="Tyreco Limited (Mumbai)", email="accounts@tyreco.example"),
    "C5": dict(name="Tyreco Limited (Nasik)", email="accounts@tyreco.example"),
    "C6": dict(name="Nilgiri Polymers Pvt. Ltd. (Chinchwad-Unit-I)", email=""),
    "C7": dict(name="Other Rubber Pvt. Ltd.", email="x@otherrubber.com"),
}
FWD = ("Kindly check trailing mail.\n\n________________\nFrom: {frm}\nSent: Tuesday\nTo: Rahul <rahul.mehta@gpglobal.com>\n"
       "Cc: Vikram <vikram.joshi@gpglobal.com>\nSubject: Re: Customer Balance Confirmation\n\n{body}\n\nFrom: Rahul <rahul.mehta@gpglobal.com>\nold mail")


def test_match_by_sender_email_without_customer_number():
    t = FWD.format(frm="AMBER TRADING CO. <ambertrading@example.in>", body="Balance is Rs. 1,67,75,671/-")
    no, how, ok, _ = rules.find_customer(t, "reply.msg", CUSTS, "Rahul Mehta")
    assert no == "C1" and ok and "e-mail" in how


def test_match_by_name_when_email_not_in_list():
    t = FWD.format(frm="Priya Rao <someone@newdomain.com>", body="Zentora India Pvt Ltd books show 3,452,112.00")
    no, how, ok, _ = rules.find_customer(t, "reply.msg", CUSTS, "")
    assert no == "C3" and ok


def test_shared_email_two_accounts_is_not_guessed():
    t = FWD.format(frm="Ravi <accounts@tyreco.example>", body="Please find our balance 5,00,000.00")
    no, how, ok, cands = rules.find_customer(t, "reply.msg", CUSTS, "")
    assert no is None and set(cands) == {"C4", "C5"}
    t2 = FWD.format(frm="Ravi <accounts@tyreco.example>", body="Tyreco Limited Nasik balance 5,00,000.00")
    assert rules.find_customer(t2, "reply.msg", CUSTS, "")[0] == "C5"


def test_gp_names_in_the_thread_do_not_match_a_customer():
    t = FWD.format(frm="Nobody <n@unknown.org>", body="Balance 10,000.00")
    assert rules.find_customer(t, "reply.msg", CUSTS, "")[0] is None


def test_two_accounts_of_one_company_need_the_balance_to_choose():
    cs = {"C1": dict(name="Amber Trading Co. (IPOL)", email="", erp=16775671.0), "C2": dict(name="Amber Trading Co. (RPO)", email="", erp=999.0)}
    t = FWD.format(frm="AMBER TRADING CO. <a@yahoo.co.in>", body="Balance Amount Is Rs. 1,67,75,671/-")
    no, how, ok, cands = rules.find_customer(t, "reply.msg", cs, "")
    assert no is None and set(cands) == {"C1", "C2"}
    assert rules.pick_by_amount(t, cands, cs) == "C1"
    assert rules.pick_by_amount(t.replace("1,67,75,671", "5,00,000"), cands, cs) is None


def test_old_quoted_table_does_not_decide_the_customer():
    t = "Note\n\nFrom: X <x@y.org>\nSubject: Re\nBalance 1,000.00\n" + "x " * 300 + "\n\nFrom: Rahul <k@gpglobal.com>\nC15343 Nilgiri Unit 7 table"
    cs = {"C15343": dict(name="Nilgiri Unit 7", email="")}
    assert rules.find_customer(t, "reply.msg", cs, "")[0] is None


def test_one_reply_for_two_customers_uses_the_line_with_the_name():
    t = ("Note\n\nFrom: A <a@gmail.com>\nSubject: Re\nBalance as per our books.\nK.P. Enterprises (RPO) Rs. 44,26,391/-\n"
         "Sunrise Enterprises (RPO) Rs.1,98,36,248/-\nOn Mon, Oct 19, 2026 at 3:26 PM Anil wrote:\n\tPlease confirm\n\t4,426,391 old table")
    a = rules.suggest(t, 4426391, "K.P. Enterprises (RPO)")
    assert a["status"] == "CLEAR" and a["theirs"] == 4426391
    b = rules.suggest(t, 19836248, "Sunrise Enterprises (RPO)")
    assert b["status"] == "CLEAR" and b["theirs"] == 19836248


def test_our_own_attached_letter_is_not_the_customers_answer():
    t = "Please find attached.\n\nThanks and Regards,\nX" + rules.ATT_MARK + "Confirmation of Balance as on 30/09/2026 Aging Details Total 13,08,384.00"
    s = rules.suggest(t, 1308384)
    assert s["needs_typing"] and s["status"] is None and any("our own" in f for f in s["flags"])


def test_i_confirm_the_same_is_clear():
    s = rules.suggest("I confirm the same.\nOn Fri wrote:\n Dear Sir, balance Rs 12,000.00", 12000)
    assert s["status"] == "CLEAR"

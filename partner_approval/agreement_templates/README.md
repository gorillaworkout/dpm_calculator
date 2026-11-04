# Agreement templates (Partner Approval, Beta)

One Word file per role, named after the role: `CM.docx`, `BM.docx`, `RM.docx`,
`MIB.docx`, `NCIB.docx`, `RD.docx`. `CM`, `BM` and `RM` are here (v1.35).

**Temporary templates (v1.37).** Until their own file exists, a role borrows
one (`TEMPLATE_FALLBACK` in `docs.py`, easy to change):

| Role | Uses   | Why |
|------|--------|-----|
| MIB  | RM.docx | individual sales role: own customers + referral rebate, no team target, no reporting line |
| NCIB | RM.docx | same as MIB |
| RD   | CM.docx | most senior template, with a team target and a reporting clause |

The filler swaps the role wording: the title ("Relationship Manager" →
"Master IB" / "NCIB", "Country Manager" → "Regional Director"), "the RM"/"a
RM's" → "the MIB"/"an MIB's", and for RD the CM line "report directly to the
Regional Director, or in their absence, to the Director of Business
Development" becomes "report directly to the Director of Business
Development". **Not changed (check before real use):** the targets
(CM's $200,000 team target applies to RD), the allowance table and rebates,
CM duties ("Support Relationship Managers and Master IBs"), "RM/BM" in CM's
fraud line, "Employee" in RM's arbitration clause. NCIB has no confirmed long
title, so "NCIB" is used. Admins see "template sementara (dari RM/CM)"; the
partner does not.

Dropping in `MIB.docx`, `NCIB.docx` or `RD.docx` (or pointing
`DW_APPROVAL_TEMPLATES` at another folder) replaces the borrowed one at once,
no code change. Requests already signed keep their PDF. If no template is found
at all, the request is still Approved with "Agreement template for <role> not
ready yet" and a "Retry agreement" button.

The filler looks for these labels, so keep them in a new template:

| Label in the template                         | Filled with                    |
|-----------------------------------------------|--------------------------------|
| `This Agreement is made on ` (CM, BM) or `This agreement was made` (RM) | Effective Date (1 November 2026); " on" is added when the template lacks it |
| `Region/Country : ` / `Name : ` / `Passport/National ID : ` / `Address : ` | Region, Partner Name, Passport, Address |
| table cell next to `Dupoin UID`               | UID                            |
| `The primary location is<tabs>,` (optional: RM has none) | Region             |
| `as of<tabs>,`                                | Effective Date                 |
| table `SIGNED by the Company` / `SIGNED by the <role>` | partner signature in the right-hand cell; the company cell stays blank |
| `Name : ` / `Passport/ID No : ` / `Date : ` under that table | Partner name, passport, sign date |
| paragraph starting `Allowance Scheme`         | a new paragraph "Agreed Allowance Scheme:" + the request text is inserted after it; the template table, rebates and notes are unchanged |
| `Remark: `                                    | Remarks                        |

`PIF.docx` is the **Personal Information Form**, not an agreement: it is never
picked as a role template, is made only for Agreement requests at approval and
is for admins only (never e-mailed, never on a /sign link). The filler writes
into the table cell that starts with each label: `Full Name`, `Address`,
`Passport / National ID Number`, `Position`, `Region / Country`,
`Email Address`, `Dupoin UID`, plus the passport after `Passport / ID No` in the
Declaration. `Nickname`, `Mobile No`, `Attachments`, `Date` and `Signature` have
no form field and stay blank.

Scrub the author in File > Info (or `docProps/core.xml`) before committing a
template: the files here say "Dupoin Markets Ltd".

Templates must not contain real partner data. Only blank templates belong here.

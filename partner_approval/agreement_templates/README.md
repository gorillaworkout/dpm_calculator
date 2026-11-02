# Agreement templates (Partner Approval, Beta)

One Word file per role, named after the role: `CM.docx`, `BM.docx`, `RM.docx`,
`MIB.docx`, `NCIB.docx`, `RD.docx`. Approving an Agreement / Special MIB request
for a role without a file here still works, but the admin sees
"Agreement template for <role> not ready yet" and no signing link is made.
Drop the file in (or point `DW_APPROVAL_TEMPLATES` at another folder) and click
"Retry agreement" on the request.

The filler looks for these labels, so keep them in a new template:

| Label in the template                         | Filled with                    |
|-----------------------------------------------|--------------------------------|
| `This Agreement is made on `                  | Effective Date (1 November 2026) |
| `Region/Country : ` / `Name : ` / `Passport/National ID : ` / `Address : ` | Region, Partner Name, Passport, Address |
| table cell next to `Dupoin UID`               | UID                            |
| `The primary location is<tabs>,`              | Region                         |
| `as of<tabs>,`                                | Effective Date                 |
| table `SIGNED by the Company` / `SIGNED by the <role>` | partner signature in the right-hand cell; the company cell stays blank |
| `Name : ` / `Passport/ID No : ` / `Date : ` under that table | Partner name, passport, sign date |
| paragraph starting `Allowance Scheme`         | a new paragraph "Agreed Allowance Scheme:" + the request text is inserted after it; the template table, rebates and notes are unchanged |
| `Remark: `                                    | Remarks                        |

Templates must not contain real partner data. Only blank templates belong here.

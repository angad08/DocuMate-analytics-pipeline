# Template setup

How to build the two Word templates DocuMate fills. You only need this if you
are changing the certificate layout or pointing DocuMate at a different form.

[Back to the README](../README.md)

---

## Template Setup

DocuMate uses two types of Word templates depending on the rendering engine. Both live in the `documate/files/templates/` folder.

### docxtpl template (v1, v2, v3, Z)

`DOCUMENT_TEMPLATE_FILE.docx` is a standard Word document with Jinja-style placeholders. These are double-curly-brace tags that `docxtpl` replaces with data at render time.

To create or edit the template, open a `.docx` file in Microsoft Word and type the placeholders directly where each value should appear. For example:

```
Name: {{Name}}
Date of Birth: {{When_and_where_born}}
Serial: {{Serial}}
```

No special Word configuration is needed. The file is a normal `.docx` document -- `docxtpl` reads the placeholders and fills them in using Python. This works on any operating system without Microsoft Word installed at runtime.

### Word Mail Merge template (Y, X, O)

`DOCUMENT_TEMPLATE_FILE_MM.docx` is a Word Mail Merge template. Unlike the docxtpl template, this one requires a one-time configuration inside Microsoft Word to link it to a data source and insert merge fields.

#### Steps to create the template

1. Open the base Word template in **Microsoft Word**.
2. Go to the **Mailings** tab in the ribbon.
3. Click **Start Mail Merge** and select **Letters** (or **Normal Word Document**).
4. Click **Select Recipients** and choose **Use an Existing List**.
5. Browse to `documate/files/data/DocuMate_DataFrame.xlsx` and select it.
6. If prompted, select the **DocuMateSRC** sheet.
7. Word now knows which data source the template is linked to.
8. Place your cursor where each field should appear in the document and click **Insert Merge Field** to add the placeholders.

### Available fields

Both templates use the same field names, which correspond to the column names produced by DocuMate's data pipeline:

- `File_Number`
- `Serial`
- `Name`
- `Sex`
- `When_and_where_born`
- `Name_of_the_Father`
- `Name_of_the_Mother`
- `Description_and_residence_of_informant`
- `Registration_date`
- `MHA_File_And_date`
- `Signing_Authority_Name`
- `Signing_Authority_Name_Designation`

In the docxtpl template, these appear as `{{Name}}`, `{{Serial}}`, etc. In the Mail Merge template, these are inserted via Word's **Insert Merge Field** button.

9. Once all merge fields are placed, save the Mail Merge template as `DOCUMENT_TEMPLATE_FILE_MM.docx` in the `documate/files/templates/` folder.

#### Note for DocuMateO

DocuMateO does not use the Excel file as its data source at runtime. It pulls records from PostgreSQL and writes them to a temporary CSV file as a bridge for Word's Mail Merge engine. The template still needs to be configured once with the Excel file (steps above) so that Word recognises the merge fields. At runtime, DocuMateO overrides the data source connection to point at the temp file automatically.

The same applies to Y and X: whatever data source the template was configured with is replaced at runtime by a temporary CSV containing only the rows to process.

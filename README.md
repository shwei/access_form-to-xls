Access Form To XLS
==================

Extracts form fields and labels from MS Access forms, 
matches the fields to corresponding labels, and
exports the results to Excel file.


Usage
------------------
Download the file, form_exportAccessMetadataToExcel-ucd-1.0.0.mdb and double-click on frmExportMSAccessMetadata Form. Recommend: MS Access 2010


Background
------------------
I had to convert several MS Access databases to REDCap (http://project-redcap.org/) projects. I needed a way to put together a data dictionary without too much of manual manipulation. This is what I put together and used. The work was based on Hal Beresford's Export Access Metadata application. Without it, the conversion effort will not be achievable in a short amount of time. Thank you, Hal Beresford.


File Structure
------------------
form_exportAccessMetadataToExcel-ucd-1.0.0.mdb  (is the program)
    src
        Form_frmExportMSAccessMetadata.cls      (Form functions)
        modFuzzy.txt                            (Module to do fuzzy string match)


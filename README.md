# Bulk-convert-WORD-to-PDF-for-free
Bulk convert WORD to PDF for free

Step to Bulk Convert Word to PDF
Step 1: Organizing word files
Move all your Ms Word documents, that needs to be converted to PDF in one folder say “fldr“. You can name it anything you like.

Step 2: Code to convert Word to PDF
Open notepad/notepad++ and copy the following code and save file as “SaveAsPDF.js” in a folder say“fldr”. Note: Don’t use Ms Word to save it.

Following code converts Word document to PDF file.

var obj = new ActiveXObject("Scripting.FileSystemObject");

var docPath = WScript.Arguments(0);

docPath = obj.GetAbsolutePathName(docPath);

var pdfPath = docPath.replace(/\.doc[^.]*$/, ".pdf");

var objWord = null;

try

{

    objWord = new ActiveXObject("Word.Application");

objWord.Visible = false;

var objDoc = objWord.Documents.Open(docPath);

var format = 17;

objDoc.SaveAs(pdfPath, format);

objDoc.Close();

WScript.Echo("Saving '" + docPath + "' as '" + pdfPath + "'...");

}

finally

{

if (objWord != null)

{

objWord.Quit();

}

}

Step 3: Create Batch file to bulk convert Word to PDF
Open notepad or notepad++ and copy the following code. Save the file with “.bat” extension say “bulk-convert-Word2PDF.bat” in folder “fldr”.


echo off

for %%X in (*.docx) do cscript.exe //nologo SaveAsPDF.js "%%X"

for %%X in (*.doc) do cscript.exe //nologo SaveAsPDF.js "%%X"

Step 4: Running Batch file
Double click batch file “bulk-convert-Word2PDF.bat” created in above step and relax. All the word documents with .docx and .doc extension in folder “fldr” is saved to the pdf file with the same name.

Summary
Simple code shown in the blog helps automate task of converting Ms Word document into pdf files. To bulk convert Word to PDF, just double click on “bulk-convert-Word2PDF.bat” and Relax!!!

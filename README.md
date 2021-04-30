<!--Category:PowerShell--> 
 <p align="right">
    <a href="https://www.powershellgallery.com/packages/ProductivityTools.PSGetOneDriveDirectory/"><img src="Images/Header/Powershell_border_40px.png" /></a>
    <a href="http://productivitytools.tech/get-onedrivedirectory/"><img src="Images/Header/ProductivityTools_green_40px_2.png" /><a> 
    <a href="https://github.com/pwujczyk/ProductivityTools.PSGetOneDriveDirectory"><img src="Images/Header/Github_border_40px.png" /></a>
</p>
<p align="center">
    <a href="http://http://productivitytools.tech/">
        <img src="Images/Header/LogoTitle_green_500px.png" />
    </a>
</p>

   
# Convert Documents

Module allows to convert document from Docx and MD format into html or markdown format.

<!--more-->

Module uses **Pandoc** application to perform conversion. It contains zipped version of the Pandoc and first step before running the script is to validate if archive is extracted.

![PandocExtraction](images/PandocExtraction.png)

## Methods
- ConvertFrom-Docx - allows to convert documents to **Html** and **Markdown**
- ConvertFrom-Markdown - allows to convert documents to **Html** and **Docx**

## Parameters

- Path - path to source document (Docx or Markdown)
- TargetFormat - target format of the document **Html**, **Docx** or **Markdown**
- OutputDirectory - directory where output file will be placed
- OutputFileName - name of the output file
- ExtractImages - if used images will be extracted from the document to files

```powershell
ConvertFrom-Markdown -verbose -Path "README.md" -TargetFormat Html -OutputDirectory D:\Trash\ArticlesHTML\MD
```

<!--ogimage-->
![PandocExtraction](images/Example.png)

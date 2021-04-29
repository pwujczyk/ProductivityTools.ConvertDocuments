clear
cd $PSScriptRoot
Import-Module .\ProductivityTools.ConvertDocuments.psm1 -Force
ConvertFrom-Docx -verbose -Path "d:\OneDrive\Dokumenty\Pawel.docx" -TargetFormat Html -OutputDirectory D:\Trash\ArticlesHTML\Docx

cd d:\Trash\Readmes\ProductivityTools.CloneGithubRepositories\
ConvertFrom-Markdown -verbose -Path "README.md" -TargetFormat Html -OutputDirectory D:\Trash\ArticlesHTML\MD

#=== articles from doc to html
#$articles2=Get-ChildItem D:\OneDrive\Dokumenty\Articles\*.docx -Recurse | Where-Object {$_.Directory -notlike "*Arch*"}
#
#foreach($article in $articles2)
#{
#	$sourceName=$article.FullName
#	$destinationDirectoryName=$article.Directory.BaseName
#	write-host $destinationDirectoryName
#	$destinationDirectoryFullName=Join-Path "d:\Trash\ArticlesHTML\" $destinationDirectoryName
#	ConvertFrom-Docx -verbose -Path $sourceName -TargetFormat Html -OutputDirectory $destinationDirectoryFullName
#}
#====
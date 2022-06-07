Import-Module .\AMHC_Report\PSWriteWord\PSWriteWord.psm1 -Force

### Define the Word file
$FilePath = ".\AMHC_Report\MilestoneTB2.docx"
$WordDocument = Get-WordDocument $FilePath

### Define JSON file
$JPath = ".\output.json"

### Test from JSON
$in = Get-Content $JPath | ConvertFrom-Json
$MilestoneXProtectVersion_Title = $in.MilestoneXProtectVersion.Product
$MilestoneXProtectVersion_SLC = $in.MilestoneXProtectVersion.SLC


### Headings & Contents
#$MilestoneXProtectVersion_Title = "Checked Item: Milestone XProtect Version"
$MilestoneXProtectVersion_wwc_text = "Why we check: `r`nStaying up to date on Milestone XProtect versions allows users to take advantage of the latest features, capabilities, and performance optimizations. New features allow users to realize increased reliability, increased security, and reduced total cost of ownership."
#$MilestoneXProtectVersion_Result_text = "Result: `r`nSoftware License Code: M01-C01-131-01-6C42B0 `r`nProduct: XProtect Corporate 2019 R1 (13.1a) `r`n `r`nLicense has been partially upgraded"
$MilestoneXProtectVersion_Result_text = "Result: `r`nSoftware License Code: " +$MilestoneXProtectVersion_SLC +"`r`nProduct: XProtect Corporate 2019 R1 (13.1a) `r`n `r`nLicense has been partially upgraded"
$MilestoneXProtectVersion_Recomendation_text = "Recommendation: `r`nA newer software version is available, XProtect Corporate 2020 R1. It is recommended to consider updating to this version that allows having new functions and system improvements."

$MilestoneCareStatus_Title = "Checked Item: Milestone Care Status"
$MilestoneCareStatus_wwc_text = "Why we check: `r`nMaintaining active Milestone Care Plus status allows users to upgrade to the latest version of Milestone XProtect as soon as it is released, with no out-of-pocket expense. Maintaining active Care Premium status allows users direct access to Milestone Technical Support 24/7. Users also receive increased priority for telephone calls and all Care Premium support cases receive a Service Level Agreement (SLA) for first response time and status update frequency."
$MilestoneCareStatus_Result_text = "Result: `r`nSoftware License Code: " +$MilestoneXProtectVersion_SLC +"`r`n `r`nCare Plus: Valid `r`n `r`nExpiration Date:"

### Replace content
foreach ($Paragraph in $WordDocument.Paragraphs) {
    $Paragraph.ReplaceText('###Title###','Micro Health Check Prepared for Engineer Name')
	$Paragraph.ReplaceText('###MilestoneXProtectVersion_Title###',$MilestoneXProtectVersion_Title)
	$Paragraph.ReplaceText('###MilestoneXProtectVersion_wwc_text###',$MilestoneXProtectVersion_wwc_text)
    $Paragraph.ReplaceText('###MilestoneXProtectVersion_Result_text###',$MilestoneXProtectVersion_Result_text)
	$Paragraph.ReplaceText('###MilestoneXProtectVersion_Recomendation_text###',$MilestoneXProtectVersion_Recomendation_text)
	
	$Paragraph.ReplaceText('###MilestoneCareStatus_Title###',$MilestoneCareStatus_Title)
	$Paragraph.ReplaceText('###MilestoneCareStatus_wwc_text###',$MilestoneCareStatus_wwc_text)
}
### Save document
Save-WordDocument $WordDocument

### Table of Content update
$word = New-Object -ComObject Word.Application
$doc = $word.Documents.Open($FilePath)
$toc = $doc.TablesOfContents
$toc.item(1).update()
$doc.save()
$doc.close()
$word.Quit()

### Bold

$wd=New-Object -ComObject Word.Application
#$wd.Visible=$true
$doc = $wd.Documents.Open($newfile)


$BoldText = 'Why we check:', 'Recommendation:', 'Result:'
Foreach ($i in $BoldText)
{
$searchText=$replaceText= $i
$newfile=$FilePath
$wdReplaceAll  = 2

$wd.Selection.Find.Replacement.Font.Bold = $true
$wd.Selection.Find.Execute($searchText, $false, $true, $false, $false, $false, $true, $false, $true, $replaceText, $wdReplaceAll)
$doc.save()
#$doc.close()
#$wd.Quit()
}

### Start Word with file
Invoke-Item $FilePath

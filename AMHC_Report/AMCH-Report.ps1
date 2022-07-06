$MyDir = [System.IO.Path]::GetDirectoryName($myInvocation.MyCommand.Definition)
$Module = $MyDir + "\PSWriteWord\PSWriteWord.psm1"
Import-Module $Module -Force
### Define the Word file

$destination = Split-Path -Path $MyDir -Parent
$FilePath = $destination + "\Template.docx"
$WordDocument = Get-WordDocument $FilePath

### Define JSON file
$JPath = $destination + "\output.json"

### Retrieve from JSON
$in = Get-Content $JPath | ConvertFrom-Json
$MilestoneXProtectVersion = $in.MilestoneXProtectVersion.Product
$MilestoneXProtectVersion_SLC = $in.MilestoneXProtectVersion.SLC
$MilestoneXProtectVersion_EXP = $in.MilestoneCareStatus.ExpirationDate
$MilestoneMediaDeletionLow_List = $in.MediaDeletionDuetoLowDiskSpace.DeletionDuetoLowDiskSpaceErrorList -join "`r`n`r`n" 
$MilestoneMediaDeletionOverflow_List = $in.MediaDeletionDuetoOverflow.MediaDeletionDuetoOverflowList -join "`r`n`r`n" 
$MSRAM_0 = $in.SystemRAMutilization.Samples.Value[0]
$MSRAM_1 = $in.SystemRAMutilization.Samples.Value[1]
$MSRAM_2 = $in.SystemRAMutilization.Samples.Value[2]
$MSRAM_3 = $in.SystemRAMutilization.Samples.Value[3]
$MSRAM_4 = $in.SystemRAMutilization.Samples.Value[4]
$MSRAM_M = $in.SystemRAMutilization.Max
$MSRAM_T = "$MSRAM_0 Kb - $MSRAM_1 Kb -  $MSRAM_2 Kb - $MSRAM_3 Kb - $MSRAM_4 Kb - Total $MSRAM_M Kb -> " + ((([Int]$MSRAM_0+[Int]$MSRAM_1+[Int]$MSRAM_2+[Int]$MSRAM_3+[Int]$MSRAM_4)/5/[Int]$MSRAM_M)).tostring("P")
$MSCPU_0 = $in.SystemCPUutilization.Samples.Value[0]
$MSCPU_1 = $in.SystemCPUutilization.Samples.Value[1]
$MSCPU_2 = $in.SystemCPUutilization.Samples.Value[2]
$MSCPU_3 = $in.SystemCPUutilization.Samples.Value[3]
$MSCPU_4 = $in.SystemCPUutilization.Samples.Value[4]
$MSCPU_M = $in.SystemCPUutilization.Max
$MSCPU_T = "$MSCPU_0% $MSCPU_1% $MSCPU_2% $MSCPU_3% $MSCPU_4% -> " + ((([Int]$MSCPU_0+[Int]$MSCPU_1+[Int]$MSCPU_2+[Int]$MSCPU_3+[Int]$MSCPU_4)/5/[Int]$MSCPU_M)).tostring("P")

$GPU1_0 = $in.HardwareaAcelerationCapability[0].Name
$GPU1_1 = $in.HardwareaAcelerationCapability[0].DriverVersion
$GPU1_2 = $in.HardwareaAcelerationCapability[0].DriverDate

$GPU2_0 = $in.HardwareaAcelerationCapability[1].Name
$GPU2_1 = $in.HardwareaAcelerationCapability[1].DriverVersion
$GPU2_2 = $in.HardwareaAcelerationCapability[1].DriverDate


# TODO: Read Languaje 

# TODO: Read Location json 

# TODO: Replace hardcoded text with the one readed from json 

### Headings & Contents
$MilestoneXProtectVersion_Title = "Checked Item: Milestone XProtect Version"
$MilestoneXProtectVersion_wwc_text = "Why we check: `r`nStaying up to date on Milestone XProtect versions allows users to take advantage of the latest features, capabilities, and performance optimizations. New features allow users to realize increased reliability, increased security, and reduced total cost of ownership."
#$MilestoneXProtectVersion_Result_text = "Result: `r`nSoftware License Code: M01-C01-131-01-6C42B0 `r`nProduct: XProtect Corporate 2019 R1 (13.1a) `r`n `r`nLicense has been partially upgraded"
$MilestoneXProtectVersion_Result_text = "Result: `r`nSoftware License Code: " +$MilestoneXProtectVersion_SLC +"`r`nProduct: "+$MilestoneXProtectVersion# +"`r`n `r`nLicense has been partially upgraded"
$MilestoneXProtectVersion_Recommendation_text = "Recommendation: `r`n[Text here is pending...]"

$MilestoneCareStatus_Title = "Checked Item: Milestone Care Status"
$MilestoneCareStatus_wwc_text = "Why we check: `r`nMaintaining active Milestone Care Plus status allows users to upgrade to the latest version of Milestone XProtect as soon as it is released, with no out-of-pocket expense. Maintaining active Care Premium status allows users direct access to Milestone Technical Support 24/7. Users also receive increased priority for telephone calls and all Care Premium support cases receive a Service Level Agreement (SLA) for first response time and status update frequency."
$MilestoneCareStatus_Result_text = "Result: `r`nSoftware License Code: " +$MilestoneXProtectVersion_SLC +"`r`n `r`nCare Plus: Valid `r`n `r`nExpiration Date: " +$MilestoneXProtectVersion_EXP
$MilestoneCareStatus_Recommendation_text = "Recommendation:" +"`r`nAn active Care agreement gives you access to updates when they are released without any additional costs, keeping you on the latest and greatest that our software has to offer. `r`nCare Premium also gives you additional entitlements like SLA, 24/7 access to support, and priority in the phone queue. You can find more information regarding Care at [$linkcareplusinfo]"

$MilestoneComulativeUpdates_Title = "Checked Item: Milestone Cumulative Updates"
$MilestoneComulativeUpdates_wwc_text = "No available on this release"
$MilestoneComulativeUpdates_Result_text = "No available on this release"
$MilestoneComulativeUpdates_Recommendation_text = "No available on this release"

$MilestoneMediaDeletionLow_Title = "Checked Item: Media Deletion Due to Low Disk Space "
$MilestoneMediaDeletionLow_wwc_text = "Why we check: `r`nWhen free disk space in XProtect recording paths fall below certain thresholds, recorded media may be at risk of deletion. Users often wish to avoid this scenario, as unexpected deletions could mean a loss of situational awareness or regulatory compliance. Adjusting recording configuration or storage hardware can prevent this scenario.   "


$MilestoneMediaDeletionLow_Result_text = "Result: `r`n" + $MilestoneMediaDeletionLow_List 
$MilestoneMediaDeletionLow_Recommendation_text = ""

$MilestoneMediaDeletionOverflow_Title = "Checked Item: Media Deletion Due to Overflow "
$MilestoneMediaDeletionOverflow_wwc_text = "Why we check: `r`nWhen the amount of inbound recorded media exceeds the disk write performance capability of the recording and storage hardware, media may be at risk of deletion due to a scenario called 'Media Overflow'. Users often wish to avoid this scenario, as unexpected deletions could mean a loss of situational awareness or regulatory compliance. Adjusting recording configuration or storage hardware can prevent this scenario. "
$MilestoneMediaDeletionOverflow_Result_text = "Result: `r`n" +$MilestoneMediaDeletionOverflow_List
$MilestoneMediaDeletionOverflow_Recommendation_text = ""

$MilestoneSystemRAM_Title = "Checked Item: System RAM utilization"
$MilestoneSystemRAM_wwc_text = "Why we check: `r`nMaintaining appropriate system RAM utilization is important for optimal performance and reliability. Most XProtect applications have minimum requirements of 2 GB of system RAM, but additional RAM may be necessary for larger systems. Typical recommended amount for best performance are 16 or 32 GB. Recommended utilization should be below 70% of max utilization.  "
$MilestoneSystemRAM_Result_text = "Result: `r`n" +[string]$MSRAM_T 
$MilestoneSystemRAM_Recommendation_text = ""

$MilestoneSystemCPU_Title = "Checked Item: System CPU utilization "
$MilestoneSystemCPU_wwc_text = "Why we check: `r`nMaintaining appropriate system CPU utilization is important for optimal performance and reliability. XProtect will install on most modern Intel or AMD x86 64-bit architecture CPUs (or virtual equivalent), however larger systems will require CPUs with higher performance. Recommended utilization should be below 70% of max utilization.  "
$MilestoneSystemCPU_Result_text = "Result: `r`n" +[string]$MSCPU_T 
$MilestoneSystemCPU_Recommendation_text = ""
	
$MilestoneFailOver_Title = "Checked Item: Failover Configuration"
$MilestoneFailOver_wwc_text = "No available on this release"
$MilestoneFailOver_Result_text = "No available on this release"
$MilestoneFailOver_Recommendation_text = "No available on this release"

$MilestoneHardwareAcceleration_Title = "Checked Item: Hardware Acceleration Capability"
$MilestoneHardwareAcceleration_wwc_text = "XProtect software can offload Recording Server motion detection tasks to certain Intel or Nvidia GPUs. This offloading of compute tasks frees up performance resources on the CPU and helps reduce the total cost of ownership by allowing users to run more cameras on a Recording Server, without reaching the performance limitations of their CPU. Intel CPUs with Quick Sync Video technology or Nvidia GPUs with Keplar and newer chipsets are required to enable hardware accelerated motion detection."
$MilestoneHardwareAcceleration_Result_text = "Result: `r`n" + [string]$GPU1_0 + " - " + [string]$GPU1_1 + " - " + [string]$GPU1_2 +  "`r`n" + [string]$GPU2_0 + " - " + [string]$GPU2_1 + " - " + [string]$GPU2_2
$MilestoneHardwareAcceleration_Recommendation_text = "No available on this release"
	
$MilestoneAntivirus_Title = "Checked Item: Antivirus Presence"
$MilestoneAntivirus_wwc_text = "No available on this release"
$MilestoneAntivirus_Result_text = "No available on this release"
$MilestoneAntivirus_Recmmendation_text = "No available on this release"

### Replace content
foreach ($Paragraph in $WordDocument.Paragraphs) {
    $Paragraph.ReplaceText('###Title###','Micro Health Check Prepared for Engineer Name')
	$Paragraph.ReplaceText('###MilestoneXProtectVersion_Title###',$MilestoneXProtectVersion_Title)
	$Paragraph.ReplaceText('###MilestoneXProtectVersion_wwc_text###',$MilestoneXProtectVersion_wwc_text)
    $Paragraph.ReplaceText('###MilestoneXProtectVersion_Result_text###',$MilestoneXProtectVersion_Result_text)
	$Paragraph.ReplaceText('###MilestoneXProtectVersion_Recommendation_text###',$MilestoneXProtectVersion_Recommendation_text)
	
	$Paragraph.ReplaceText('###MilestoneCareStatus_Title###',$MilestoneCareStatus_Title)
	$Paragraph.ReplaceText('###MilestoneCareStatus_wwc_text###',$MilestoneCareStatus_wwc_text)
	$Paragraph.ReplaceText('###MilestoneCareStatus_Result_text###',$MilestoneCareStatus_Result_text)
	$Paragraph.ReplaceText('###MilestoneCareStatus_Recommendation_text###',$MilestoneCareStatus_Recommendation_text)
	
	$Paragraph.ReplaceText('###MilestoneComulativeUpdates_Title###',$MilestoneComulativeUpdates_Title)
	$Paragraph.ReplaceText('###MilestoneComulativeUpdates_wwc_text###',$MilestoneComulativeUpdates_wwc_text)
	$Paragraph.ReplaceText('###MilestoneComulativeUpdates_Result_text###',$MilestoneComulativeUpdates_Result_text)
	$Paragraph.ReplaceText('###MilestoneComulativeUpdates_Recommendation_text###',$MilestoneComulativeUpdates_Recommendation_text)
	
	$Paragraph.ReplaceText('###MilestoneMediaDeletionLow_Title###',$MilestoneMediaDeletionLow_Title)
	$Paragraph.ReplaceText('###MilestoneMediaDeletionLow_wwc_text###',$MilestoneMediaDeletionLow_wwc_text)
	$Paragraph.ReplaceText('###MilestoneMediaDeletionLow_Result_text###',$MilestoneMediaDeletionLow_Result_text)
	$Paragraph.ReplaceText('###MilestoneMediaDeletionLow_Recommendation_text###',$MilestoneMediaDeletionLow_Recommendation_text)
	
	$Paragraph.ReplaceText('###MilestoneMediaDeletionOverflow_Title###',$MilestoneMediaDeletionOverflow_Title)
	$Paragraph.ReplaceText('###MilestoneMediaDeletionOverflow_wwc_text###',$MilestoneMediaDeletionOverflow_wwc_text)
	$Paragraph.ReplaceText('###MilestoneMediaDeletionOverflow_Result_text###',$MilestoneMediaDeletionOverflow_Result_text)
	$Paragraph.ReplaceText('###MilestoneMediaDeletionOverflow_Recommendation_text###',$MilestoneMediaDeletionOverflow_Recommendation_text)
	
	$Paragraph.ReplaceText('###MilestoneSystemRAM_Title###',$MilestoneSystemRAM_Title)
	$Paragraph.ReplaceText('###MilestoneSystemRAM_wwc_text###',$MilestoneSystemRAM_wwc_text)
	$Paragraph.ReplaceText('###MilestoneSystemRAM_Result_text###',$MilestoneSystemRAM_Result_text)
	$Paragraph.ReplaceText('###MilestoneSystemRAM_Recommendation_text###',$MilestoneSystemRAM_Recommendation_text)

	$Paragraph.ReplaceText('###MilestoneSystemCPU_Title###',$MilestoneSystemCPU_Title)
	$Paragraph.ReplaceText('###MilestoneSystemCPU_wwc_text###',$MilestoneSystemCPU_wwc_text)
	$Paragraph.ReplaceText('###MilestoneSystemCPU_Result_text###',$MilestoneSystemCPU_Result_text)
	$Paragraph.ReplaceText('###MilestoneSystemCPU_Recommendation_text###',$MilestoneSystemCPU_Recommendation_text)
	
	$Paragraph.ReplaceText('###MilestoneFailOver_Title###',$MilestoneFailOver_Title)
	$Paragraph.ReplaceText('###MilestoneFailOver_wwc_text###',$MilestoneFailOver_wwc_text)
	$Paragraph.ReplaceText('###MilestoneFailOver_Result_text###',$MilestoneFailOver_Result_text)
	$Paragraph.ReplaceText('###MilestoneFailOver_Recommendation_text###',$MilestoneFailOver_Recommendation_text)

	$Paragraph.ReplaceText('###MilestoneHardwareAcceleration_Title###',$MilestoneHardwareAcceleration_Title)
	$Paragraph.ReplaceText('###MilestoneHardwareAcceleration_wwc_text###',$MilestoneHardwareAcceleration_wwc_text)
	$Paragraph.ReplaceText('###MilestoneHardwareAcceleration_Result_text###',$MilestoneHardwareAcceleration_Result_text)
	$Paragraph.ReplaceText('###MilestoneHardwareAcceleration_Recommendation_text###',$MilestoneHardwareAcceleration_Recommendation_text)
	
	$Paragraph.ReplaceText('###MilestoneAntivirus_Title###',$MilestoneAntivirus_Title)
	$Paragraph.ReplaceText('###MilestoneAntivirus_wwc_text###',$MilestoneAntivirus_wwc_text)
	$Paragraph.ReplaceText('###MilestoneAntivirus_Result_text###',$MilestoneAntivirus_Result_text)
	$Paragraph.ReplaceText('###MilestoneAntivirus_Recmmendation_text###',$MilestoneAntivirus_Recmmendation_text)	
	
}
### Save document
Save-WordDocument $WordDocument -FilePath "$destination\Report.docx" -Supress $true

# ### Table of Content update
# $word = New-Object -ComObject Word.Application
# $doc = $word.Documents.Open($FilePath)
# $toc = $doc.TablesOfContents
# $toc.item(1).update()
# $doc.save()
# $doc.close()
# $word.Quit()

# ### Bold

# $wd=New-Object -ComObject Word.Application
# #$wd.Visible=$true
# $doc = $wd.Documents.Open($newfile)


# $BoldText = 'Why we check:', 'Recommendation:', 'Result:'
# Foreach ($i in $BoldText)
# {
# $searchText=$replaceText= $i
# $newfile=$FilePath
# $wdReplaceAll  = 2

# $wd.Selection.Find.Replacement.Font.Bold = $true
# $wd.Selection.Find.Execute($searchText, $false, $true, $false, $false, $false, $true, $false, $true, $replaceText, $wdReplaceAll)
# $doc.save()
# #$doc.close()
# #$wd.Quit()
# }

### Start Word with file

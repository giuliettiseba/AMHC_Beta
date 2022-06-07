function Get-WordParagraphs {
    [CmdletBinding()]
    param(
        [parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Document.NET.Container] $WordDocument
    )
    $Paragraphs = @()
    foreach ($p in $WordDocument.Paragraphs) {
        #Write-Verbose "Get-WordParagraphs - $p"
        $Paragraphs += $p
    }
    return $Paragraphs
}
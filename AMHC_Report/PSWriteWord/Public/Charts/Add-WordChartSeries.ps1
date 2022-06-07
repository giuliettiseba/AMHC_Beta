function Add-WordChartSeries {
    [CmdletBinding()]
    param (
        [string] $ChartName = 'Legend',
        [string[]] $Names,
        [int[]] $Values
    )

    [Array] $rNames = foreach ($Name in $Names) {
        $Name
    }
    [Array] $rValues = foreach ($value in $Values) {
        $value
    }
    [Xceed.Document.NET.Series] $series = [Xceed.Document.NET.Series]::new($ChartName)
    $Series.Bind($rNames, $rValues)
    return $Series
}
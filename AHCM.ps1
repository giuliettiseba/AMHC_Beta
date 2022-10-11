Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Milestone Micro Healt Check'
$form.Size = New-Object System.Drawing.Size(340,300)
$form.StartPosition = 'CenterScreen'

$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(30,20)
$label.Size = New-Object System.Drawing.Size(100,15)
$label.Text = 'Select Language:'
$form.Controls.Add($label)

$Language = New-Object System.Windows.Forms.ComboBox
$Language.Location = New-Object System.Drawing.Point(140,20)
$Language.Size = New-Object System.Drawing.Size(120,15)

[String[]]$Languages = "English",
    "Spanish",
    "French",
    "Portuguese"
$Language.Items.AddRange($Languages);
$Language.SelectedIndex = 0;
$form.Controls.Add($Language)

$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(30,50)
$label2.Size = New-Object System.Drawing.Size(100,15)
$label2.Text = 'Engineer:'
$form.Controls.Add($label2)

$Engineer = New-Object System.Windows.Forms.TextBox
$Engineer.Location = New-Object System.Drawing.Point(140,50)
$Engineer.Size = New-Object System.Drawing.Size(120,20)
$form.Controls.Add($Engineer)

$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(30,80)
$label3.Size = New-Object System.Drawing.Size(100,15)
$label3.Text = 'Account Name:'
$form.Controls.Add($label3)

$AccountName = New-Object System.Windows.Forms.TextBox
$AccountName.Location = New-Object System.Drawing.Point(140,80)
$AccountName.Size = New-Object System.Drawing.Size(120,20)
$form.Controls.Add($AccountName)

$label4 = New-Object System.Windows.Forms.Label
$label4.Location = New-Object System.Drawing.Point(30,110)
$label4.Size = New-Object System.Drawing.Size(100,15)
$label4.Text = 'Contact Name:'
$form.Controls.Add($label4)

$ContactName = New-Object System.Windows.Forms.TextBox
$ContactName.Location = New-Object System.Drawing.Point(140,110)
$ContactName.Size = New-Object System.Drawing.Size(120,20)
$form.Controls.Add($ContactName)

$label5 = New-Object System.Windows.Forms.Label
$label5.Location = New-Object System.Drawing.Point(30,140)
$label5.Size = New-Object System.Drawing.Size(100,15)
$label5.Text = 'Report Number:'
$form.Controls.Add($label5)

$ReportNumber = New-Object System.Windows.Forms.TextBox
$ReportNumber.Location = New-Object System.Drawing.Point(140,140)
$ReportNumber.Size = New-Object System.Drawing.Size(120,20)
$form.Controls.Add($ReportNumber)

$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(160,220)
$okButton.Size = New-Object System.Drawing.Size(75,23)
$okButton.Text = 'OK'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $okButton
$form.Controls.Add($okButton)

$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(80,220)
$cancelButton.Size = New-Object System.Drawing.Size(75,23)
$cancelButton.Text = 'Cancel'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$form.CancelButton = $cancelButton
$form.Controls.Add($cancelButton)

$form.Topmost = $true

$form.Add_Shown({$Language.Select()})
$result = $form.ShowDialog()

if ($result -eq [System.Windows.Forms.DialogResult]::OK)
{
    $Language.SelectedItem
    $Engineer.Text
    $AccountName.Text
    $ContactName.Text
    $ReportNumber.Text
}
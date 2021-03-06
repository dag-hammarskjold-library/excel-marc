.\env\Scripts\activate
Add-Type -AssemblyName System.Windows.Forms

$chooser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ InitialDirectory = [Environment]::GetFolderPath('Desktop') }
$chooser.ShowDialog()
$fn = $chooser.filename
$defaults = ''
$out = "$($fn).mrc"

if (test-path 'defaults.xlsx' -PathType leaf) {
	$defaults = 'defaults.xlsx'
	write-output "Applying defaults from $($defaults)" 
}

excel_marc --connect=$env:mdb --file=$fn --type=bib --format=mrc --check=191a --defaults=$defaults --out=$out

if (!$?) {
	powershell -noexit
}

write-output $out | set-clipboard 
write-output "MARC file: " $out `n` "The file path has been added to your clipboard" `n`

powershell -noexit

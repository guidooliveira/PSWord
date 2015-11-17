$fileList = @(
'DocX.dll',
'MadMilkman.Docx.dll',
'MadMilkman.Docx.xml',
'PSWord.dll',
'PSWord.psd1',
'en-US/PSWord.dll-Help.xml'
)


$InstallDirectory = Join-Path -Path "$([Environment]::GetFolderPath('MyDocuments'))\WindowsPowershell\Modules" -ChildPath PSWord

if (!(Test-Path $InstallDirectory))
{
    $null = mkdir $InstallDirectory
    $null = mkdir $InstallDirectory\en-US
}

$wc = new-object System.Net.WebClient
$fileList | ForEach-Object {
    Try
    {
        $wc.DownloadFile("https://github.com/guidooliveira/PSWord/raw/master/Release/$_", "$installDirectory\$_") 
    }
    catch
    {
        Write-Error -Message "Error Downloading $_"
    }
}
$Logfile = Read-Host("LOG FILE:")
$read_path = Read-Host("SOURCE PATH:")
$write_path = Read-Host("DESTINATION PATH:")
$error_path = Read-Host("EXCEPTION PATH:")
$passwd = Read-Host("PASSWORD:")
$counter=1
$WordObj = New-Object -ComObject Word.Application

Function LogWrite
{
   Param ([string]$logstring)

   Add-content $Logfile -value $logstring
}

foreach ($file in $count=Get-ChildItem $read_path -Filter *.doc*)
{
    try{
        $WordObj.Visible = $false
        Write-Host("UNLOCKING: " + $file.Name)
        $WordDoc = $WordObj.Documents.Open($file.FullName, $null, $false, $null, $passwd, $passwd)
        $WordDoc.Activate()
        $WordDoc.Password=""
        $WordDoc.SaveAs($write_path + $file)
        $WordDoc.Close()
        Write-Host("FINISHED: "+$counter+" of "+$count.Length)
        Remove-Item -Path ($read_path + $file)
        $counter++
    } catch {
        LogWrite("An Error Occured: " + $_)
        LogWrite("File: " + $file.Name + " could not be Unlocked.")
        Move-Item -Path ($read_path + $file) -Destination ($error_path + $file)
    }
}
$WordObj.Application.Quit()

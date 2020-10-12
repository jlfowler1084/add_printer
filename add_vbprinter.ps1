 function Add-vbPrinter {
    Param(
    [CmdletBinding()]
    [Parameter(Mandatory=$true,
               ValueFromPipeline=$true,
               Position=1,
               HelpMessage='Enter the name of the new printer here')]
            [string]$PrinterName,

    [Parameter(Mandatory=$true,
               ValuefromPipeline=$true,
               Position=2,
               HelpMessage='Enter the IP Address of the Printer here')] 
            [ipaddress]$IPAddress   
    )


 $OfficeMenu=@"

----------------------------------------------

1. Office1 
2. Office2
3. Office3
4. Office4
5. Office5

Q. Quit

----------------------------------------------

Select the Office by number or Q to quit
"@

Write-Host 'Select Office Location for new printer'

$r = Read-Host $OfficeMenu


Switch ($r){

"1"{
    write-host "Setting Server to Office 1 Print Server" -ForegroundColor Green
    $Server = 'Office1PrintServer'
    }

"2"{
    write-host "Setting Server to Office 2 Print Server" -ForegroundColor Green
    $Server = 'Office2PrintServer'
    }

"3"{
    write-host "Setting Server to Office 3 Print Server" -ForegroundColor Green
    $Server = 'Office3PrintServer'
    }

"4"{
    write-host "Setting Server to "Office 4 Print Server" -ForegroundColor Green
    $Server = 'Office4PrintServer'
    }

"5"{
    write-host "Setting Server to Office 5 Print Server" -ForegroundColor Green
    $Server = 'Office5PrintServer'
    }

"Q" {
    Write-Host "Quitting" -ForegroundColor Green
    break
    }
        
}

$DriverMenu=@"

----------------------------------------------

1. HP LaserJet 600 M601 M602 M603
2. HP LaserJet 4350
3. HP LaserJet M606
4. Ricoh MultiFunction Device (MFP)
5. HP Color LaserJet M750 PCL 6
6. HP LaserJet M604 PCL 6
7. HP LaserJet M605 PCL 6
8. HP Universal Printing PCL 6 (v6.6.5)

Q. Quit

----------------------------------------------

Select the Print Model by number or Q to quit

"@

$r= Read-Host $DriverMenu

Switch ($r) {

"1" {
    write-host 'Configure for HP LaserJet 600 Series' -ForegroundColor Green
    $DriverName = 'HP LaserJet 600 M601 M602 M603 PCL6'
    }

"2" {
    write-host 'Configure for HP LaserJet 4350 Series' -ForegroundColor Green
    $DriverName = 'HP LaserJet 4350 PCL 6'
    }

"3" {
    write-host 'Configure for HP LaserJet M606 Series' -ForegroundColor Green
    $DriverName = 'HP LaserJet M606 PCL 6'
    }
    
"4" {
    write-host 'Configure for RICOH Universal MFP' -ForegroundColor Green
    $DriverName = 'PCL6 Driver for Universal Print'
    }

"5" {
    write-host 'HP Color LaserJet M750 PCL 6' -ForegroundColor Green
    $DriverName = 'HP Color LaserJet M750 PCL 6'
    } 

"6" {
    write-host 'HP LaserJet M604 PCL 6' -ForegroundColor Green
    $DriverName = 'HP LaserJet M604 PCL 6'
    }
    
"7" {
    write-host 'HP LaserJet M605 PCL 6' -ForegroundColor Green
    $DriverName = 'HP LaserJet M605 PCL 6'
    }
    
"8" {
    write-host 'HP Universal Printing PCL 6 (v6.6.5)' -ForegroundColor Green
    $DriverName = 'HP Universal Printing PCL 6 (v6.6.5)'
    }      
    
"Q" {
    Write-Host "Quitting" -ForegroundColor Green
    break
    } 
}

 $Port = "$PrinterName.venable.com"
 
 switch($Server){
    Office1PrintServer {Add-DnsServerResourceRecord -ComputerName office1_DC -A -Name $PrinterName -ZoneName "venable.com" -IPv4Address $IPAddress.ipaddressToString}
    Office2PrintServer {Add-DnsServerResourceRecord -ComputerName office2_DC -A -Name $PrinterName -ZoneName "venable.com" -IPv4Address $IPAddress.ipaddressToString}
    Office3PrintServer {Add-DnsServerResourceRecord -ComputerName office3_DC -A -Name $PrinterName -ZoneName "venable.com" -IPv4Address $IPAddress.ipaddressToString}
    Office4PrintServer {Add-DnsServerResourceRecord -ComputerName office4_DC -A -Name $PrinterName -ZoneName "venable.com" -IPv4Address $IPAddress.ipaddressToString}
    Office5PrintServer {Add-DnsServerResourceRecord -ComputerName office5_DC -A -Name $PrinterName -ZoneName "venable.com" -IPv4Address $IPAddress.ipaddressToString}
 }


 write-host "Please Wait for DNS Cache to Flush" -ForegroundColor Green

 Start-Sleep -Seconds 10

 Invoke-WmiMethod -class Win32_process -name Create -ArgumentList ("cmd.exe /c ipconfig /flushdns") -ComputerName $Server 
 
 Start-Sleep -seconds 10

 Invoke-WmiMethod -class Win32_process -name Create -ArgumentList ("cmd.exe /c ipconfig /flushdns") -ComputerName $Server

 Start-Sleep -scon 10

 Add-PrinterPort -ComputerName $Server -Name $Port -PrinterHostAddress $Port
 
 Add-Printer -ComputerName $Server -Name $PrinterName -DriverName $DriverName -PortName $Port -Published:$true
 
 }

 Add-vbPrinter



 

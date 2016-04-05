$urls_coconut = Import-Csv "C:\Users\blankenr\Desktop\ISAAC\scripts\coconut.csv" -Delimiter ";"
$urls = Import-Csv "C:\Users\blankenr\Desktop\ISAAC\scripts\ISAAC.csv" -Delimiter ";"

Open-IETabs -object $urls_coconut 

Open-IETabs_ISAAC -object $urls

Function Open-IETabs {
    [cmdletbinding()]
    param (
        $object
    )
    begin {
        $Ie = New-Object -ComObject InternetExplorer.Application
    }

    process {
        foreach ($url in $object) {
            $nummer = [array]::IndexOf($object, $url)
            Write-Debug  "nummer is: $nummer"
            Write-Debug "Gaat nu naar :  $url.URLS"

            if($nummer -eq 0){
                $Ie.Navigate($url.URLS)
                }
            else{
                $Ie.Navigate2($url.URLS, 0x1000)
            }

            while($ie.ReadyState -ne 4) {start-sleep -m 100} 
            Write-host "heeft nu de pagina geladen :  + $url.URLS"

            Write-host "De URL ID is:"  $url.ID
            if($url.ID){ Write-host "Gaat de username invoegen :" + $url.USERNAME
                $ie.document.getElementById($url.ID).value = $url.USERNAME
            }
        }
    }

    end {
        $Ie.Visible = $true
    }
}


Function Open-IETabs_ISAAC {
    [cmdletbinding()]
    param (
        $object
    )
    begin {
        $Ie = New-Object -ComObject InternetExplorer.Application
    }

    process {
        foreach ($url in $object) {
            Write-host "Gaat nu naar :  $url.URLS"

            $Ie.Navigate($url.URLS)

            Write-host "heeft nu de pagina geladen :  + $url.URLS"
            Write-host "Gaat nu even slapen"
            start-sleep -m 1000 

            Write-host "De URL ID is: $url.ID"
            if($url.ID){ Write-host "Gaat de username invoegen :" + $url.USERNAME
                $ie.document.getElementById($url.ID).value = $url.USERNAME
            }
           
            if($url.URLS -eq "https://nwo.acc.isaac.spinozanet.nl/nl/home"){
                Write-host "begint nu aan ISAAC"
                $shell = New-Object -ComObject Shell.Application
                $ieTabs = $shell.Windows()
                $ie_ISAAC = $ieTabs | ? {$_.LocationURL -eq "https://nwo.acc.isaac.spinozanet.nl/nl/home"} 
                $ie_ISAAC.Document.getElementById($url.ID).value = $url.USERNAME

                Write-host "begint nu aan Alfresco"
                $shell = New-Object -ComObject Shell.Application
                $ieTabs = $shell.Windows()
                $ie_ISAAC = $ieTabs | ? {$_.LocationURL -eq "https://nwo.acc.isaac.spinozanet.nl/nl/home"} 
                $ie_ISAAC[0].Navigate("https://nwoaapp01.acc.isaac.spinozanet.nl/share/page/", 2048)
     
            }
        }
    }

    end {
        $Ie.Visible = $true
    }
}



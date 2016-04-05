$urls = Import-Csv "Desktop\ISAAC\scripts\urls.csv" -Delimiter ";"


$IE=new-object -com internetexplorer.application
$IE.visible=$true
  $IE.navigate($url.URLS)
  $nummer = [array]::IndexOf($urls, $url)
foreach($url in $urls){

    $IE.navigate2($url.URLS,0x1000)
    while($ie.ReadyState -ne 4) {start-sleep -m 100} 
    $ie.document.getElementById($url.ID).value = $url.USERNAME

}

Function Open-IETabs {
    [cmdletbinding()]
    param (
        $object
    )
    begin {
        $Ie = New-Object -ComObject InternetExplorer.Application
    }

    process {
      $Ie.Visible = $true

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

            Write-Debug "heeft nu de pagina geladen :  + $url.URLS"
            while($ie.ReadyState -ne 4) {write-host $ie.ReadyState
            start-sleep -m 100} 

            Write-Debug $url.ID
            if($url.ID){ Write-host "Gaat de username invoegen :" + $url.USERNAME
            $ie.document.getElementById($url.ID).value = $url.USERNAME
            }
        }
    }

    end {
        $Ie.Visible = $true
    }
}

Open-IETabs -object $urls -Debug


Write-host "begint nu met nieuwe shell"
$shell = New-Object -ComObject Shell.Application
$ieTabs = $shell.Windows()
$ie_ = $ieTabs | ? {$_.LocationURL -eq "https://nwo.acc.isaac.spinozanet.nl/nl/beheer1"} 

# De sjabloon kolom bevat nooit <br>. En de Type kolom wel. Hier is dus de scheiding op gemaakt. 
$knoppen = $ie_.Document.documentElement.getElementsByTagName("div") |   Where {$_.IHTMLElement_className -eq 'aq-answer-holder '-and $_.IHTMLElement_outerText -ne "" -and ($_.IHTMLElement_outerHTML -match "<br>")} 
Write-Host "Aantal velden gevonden: " + $knoppen.count

#$knoppen |  select -first 3 | select -Last 1

for($i=1;$i -lt $knoppen.count ; ( $i ++)){
    while(! ($ie_.Document.documentElement.getElementsByTagName("div") |   Where {$_.IHTMLElement_className -eq 'aq-answer-holder '})) {
        Start-Sleep -m 100
    }
    # knoppen opnieuw inladen want anders werkt het niet
    $knoppen = $ie_.Document.documentElement.getElementsByTagName("div") |   Where {$_.IHTMLElement_className -eq 'aq-answer-holder '-and $_.IHTMLElement_outerText -ne "" -and ($_.IHTMLElement_outerHTML -match "<br>")} 

    # de juiste knoppakken
    $knop = $knoppen | select -First $i  | select -Last 1
    $sjabloon = $knoppen | select -First ($i + 1) | select -Last 1
    $sjabloon.IHTMLElement_outerText 
      
    Write-Host "Gaat nu op knop drukken:" + $knop.IHTMLElement_outerHTML + $knop.outerText -ForegroundColor Cyan
    $knop.click()

    while(!($ie_.document.getElementById("zoektermSjabloon"))){
        write-host "Kan nu nog geen waarde invullen"
        start-sleep -m 100
    }
  
    # zoekwaarde invullen
    Write-Host "Vult nu de zoekwaarde in" -ForegroundColor Cyan
    $ie_.document.getElementById("zoektermSjabloon").value = "test"

    # zoek knop indrukken
    $knop_search  = $ie_.Document.getElementsByName("search") | select -First 1
    $knop_search.click()

    while(!($ie_.Document.getElementsByName("select") | select -First 2 | select -last 1)){
        Write-Host "kan nu nog geen resultaat selecteren"
        start-sleep -m 100
    }

    # als zoekresultaten geladen zijn dan eerste in resultaat indrukken
    $knop_sjabloon = $ie_.Document.getElementsByName("select") | select -First 2 | select -last 1
    $knop_sjabloon.click()  
}



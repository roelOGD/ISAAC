Write-Host "begint nu met nieuwe shell"
$shell = New-Object -ComObject Shell.Application
$ieTabs = $shell.Windows()
$ie_ = $ieTabs | ? {$_.LocationURL -eq "https://nwo.acc.isaac.spinozanet.nl/nl/beheer1"} 

$newKnoppen_initieel = newKnoppen

function newKnoppen(){
    # De sjabloon kolom bevat nooit <br>. En de Type kolom wel. Hier is dus de scheiding op gemaakt. 
    #alleen op het moment dat sjabloon niet is ingevuld, dan moet die meegenomen worden
    [System.Collections.Generic.List[System.Object]]$knoppen = $ie_.Document.documentElement.getElementsByTagName("div") |   Where {$_.IHTMLElement_className -eq 'aq-answer-holder '} 
    $newKnoppen =@()
    for($i = 0; $i -lt $knoppen.count ; $i++){
        if(  (($knoppen[$i].IHTMLElement_outerHTML -match "<br>")  -and ($knoppen[($i +1)].IHTMLElement_outerHTML -match "<br>")) -or  ( ($i -eq ($knoppen.count -1 )) -and ($knoppen[($i)].IHTMLElement_outerHTML -match "<br>") ) ){
          $newKnoppen += $knoppen[$i]
           Write-Host $i
        }
        else{
            Write-Host $i -ForegroundColor Red
        }
   
    }

    Write-Host "Aantal velden gevonden: " + $Knoppen.count
    Write-Host "Aantal velden gevonden die gedaan moeten worden: " + $newKnoppen.count

    return $newKnoppen
}



for($i=1;$i -le $newKnoppen_initieel.count ; ( $i ++)){
    write-host "$i / " $newKnoppen_initieel.count -ForegroundColor Yellow 
    while(! ($ie_.Document.documentElement.getElementsByTagName("div") |   Where {$_.IHTMLElement_className -eq 'aq-answer-holder '})) {
        Start-Sleep -m 100
    }
    # knoppen opnieuw inladen want anders werkt het niet
    $newKnoppen = newKnoppen

    # de juiste knoppakken
    $knop = $newKnoppen | select -First 1

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



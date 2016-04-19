$username = "piet"
$password = ""

# Een object met IE tab
$shell = New-Object -ComObject Shell.Application
$ieTabs = $shell.Windows()
$ie_ = $ieTabs | ? {$_.LocationURL -eq "https://nwo.acc.isaac.spinozanet.nl/nl/home"} 

If(!$ie_ ){
    write-host "Er moet een IE tablat met de volgende URL opens staan: https://nwo.acc.isaac.spinozanet.nl/nl/home" -ForegroundColor Cyan
    Write-Host "Start nu een IE browser op!"
     $ie_ = New-Object -ComObject InternetExplorer.Application
     $ie_.Visible = $true
     $ie_.Navigate("https://nwo.acc.isaac.spinozanet.nl/nl/home")
     start-sleep 3
    
    # Een object met IE tab
    $shell = New-Object -ComObject Shell.Application
    $ieTabs = $shell.Windows()
    $ie_ = $ieTabs | ? {$_.LocationURL -eq "https://nwo.acc.isaac.spinozanet.nl/nl/home"} 
}

#inloggen als het nodig is
if($ie_.document.getElementById("_58_login")){
    Write-Host "Moet eerst inloggen... Vult nu de credentials in!" -ForegroundColor Cyan
    $tijd_voor_inloggen = Measure-Command {
        $ie_.document.getElementById("_58_login").value = $username
        $ie_.document.getElementById("_58_password").value = $password 
        $ie_.Document.documentElement.getElementsByTagName("Aanmelden")
        $inlogKnop =  $ie_.document.getElementById("aui_3_4_0_1_270")  
        $inlogKnop = $ie_.Document.documentElement.getElementsByTagName("INPUT") | Where {$_.IHTMLInputButtonElement_type -match 'submit'}
        $inlogKnop.click()
            
        while( ! ($ie_.Document.documentElement.getElementsByTagName("h2") | where {$_.outerHTML -match 'Welkom'}) ){
            write-host "Kan nu nog geen waarde invullen"
            start-sleep -m 100
            #Write-Host $ie_.busy
        }
    }    
    Write-Host "De totale tijd voor het inloggen was:" $tijd_voor_inloggen.Seconds "seconden"  en $tijd_voor_inloggen.Milliseconds "Miliseconden" -ForegroundColor Cyan  
        
    # schrijf data weg naar een CSV bestand       
    #$array = @{
    #    
    #}
    #Add-Content -Path "H:\output.csv" $array
}
else {
    Write-Host "User is reeds ingelogd" -ForegroundColor Cyan
}

# navigeer naar tabje ´Aanvragen´
    $tijd_voor_aanvragen_klikken = Measure-Command {
    $menu_aanvragen = $ie_.Document.documentElement.getElementsByTagName("a")  | Where {$_.innerText -match 'Aanvragen'}
    $menu_aanvragen.click()

    while($ie_.busy) {
        write-host "Pagina nog niet geladen"
        start-sleep -m 100
    }
       
    Write-Host "De totale tijd voor het laden van de aanvraagpagina was:" $tijd_voor_aanvragen_klikken.Seconds "seconden"  en $tijd_voor_aanvragen_klikken.Milliseconds "Miliseconden" -ForegroundColor Cyan 
    # schrijf data weg naar een CSV bestand       
    #$array = @{
    #    
    #}
    #Add-Content -Path "H:\output.csv" $array 
}



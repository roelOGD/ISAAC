$failedfiles = Get-content c:\temp\viruslist.txt
$log = 'C:\temp\restorefromshadowcopy2.log'
import-module "C:\Temp\lib\net40\AlphaFS.dll"
#new-psdrive -PSProvider Filesystem -Name V -Root \\localhost\E$\@GMT-2016.03.25-11.00.10 |out-null
#new-psdrive -PSProvider Filesystem -Name W -Root \\localhost\E$\@GMT-2016.03.25-06.00.05 |out-null

function write-log ($logfile,$status,$message,$data,$progress){
	$logstring = [string](get-date -format u)
	$logstring = $logstring.substring(0,$logstring.length - 1)
	$logstring += " ,$status"
	$logstring = $logstring.padright(39)
	$logstring += " ,$message"
	$logstring = $logstring.padright(69)
	$logstring += " ,[$data]"
	$logstring += " ,[$progress]"
	$logstring | out-file -filepath $logfile -append
}


$count = $failedfiles.count
[int]$progress =  0
[int]$filenumber = 0
foreach($failedfile in $failedfiles){
	$destination = $failedfile.substring(0,$failedfile.Length -8)
	$filenumber ++
	$progress = ($filenumber * 100) / $count
	write-progress -id 1 -activity 'restoring files' -percentcomplete $progress -currentoperation $destination -status 'Progress'
	$fileroot = $destination | split-path -noqualifier
	$sourceA = "\\localhost\E$\@GMT-2016.03.25-11.00.10\${fileroot}"
    $sourceB = "\\localhost\E$\@GMT-2016.03.25-06.00.05\${fileroot}"
    
	write-log -logfile $log -Status Info -Message "trying to restore from source a" -data "$destination" -progress $progress
    $copied=$false
	try{
		$isavailable = $false
		#$isavailable = test-path $sourceA
        $isavailable = [Alphaleonis.Win32.Filesystem.File]::Exists($sourceA)
		if($isavailable){
			$isblocked = $false
			#$isblocked = test-path $destination
            $isblocked = [Alphaleonis.Win32.Filesystem.File]::Exists($destination)
			if($isblocked){
				write-log -logfile $log -Status Warning -Message "file already present and not copied" -data "$destination" -progress $progress
                $copied = $true
			} else {
				#copy-item -path $sourceA -destination $destination
                [Alphaleonis.Win32.Filesystem.File]::Copy($sourceA, $destination) 
				write-log -logfile $log -Status Info -Message "file copied from source A" -data "$destination" -progress $progress
                $copied = $true
			}
		} else {
			write-log -logfile $log -Status warning -Message "File not available in source a" -data "$destination" -progress $progress
		}
	} catch {
		write-log -logfile $log -Status Error -Message "Could not copy file from source a" -data "${$destination}:${$_.exception}" -progress $progress
	}
	if(-not ($copied)){
        write-log -logfile $log -Status Info -Message "trying to restore from source b" -data "$destination" -progress $progress
        try{
	    	$isavailable = $false
	    	#$isavailable = test-path $sourceB
            $isavailable = [Alphaleonis.Win32.Filesystem.File]::Exists($sourceB)
	    	if($isavailable){
	    		$isblocked = $false
	    		#$isblocked = test-path $destination
                $isblocked = [Alphaleonis.Win32.Filesystem.File]::Exists($destination)

	    		if($isblocked){
	    			write-log -logfile $log -Status Warning -Message "file already present and not copied" -data "$destination" -progress $progress
                    $copied = $true
	    		} else {
	    			#copy-item -path $sourceB -destination $destination 
                    [Alphaleonis.Win32.Filesystem.File]::Copy($sourceB, $destination) 
	    			write-log -logfile $log -Status Info -Message "file copied from source B" -data "$destination" -progress $progress
                    $copied = $true
	    		}
	    	} else {
	    		write-log -logfile $log -Status Error -Message "File not available in shadowcopy B" -data "$destination" -progress $progress
	    	}
	    } catch {
	    	write-log -logfile $log -Status Error -Message "Could not copy file from sahdowcopy B" -data "${$destination}:${$_.exception}" -progress $progress
	    }
    }
    if($copied){
        write-log -logfile $log -Status Info -Message "deleting encrypted file" -data "$failedfile" -progress $progress
        #$ispresentinfected = test-path $failedfile
        $ispresentinfected = [Alphaleonis.Win32.Filesystem.File]::Exists($failedfile)
        if($ispresentinfected){
        try{
            #remove-item $failedfile -force
            [Alphaleonis.Win32.Filesystem.File]::Delete($failedfile,[Alphaleonis.Win32.Filesystem.PathFormat]::FullPath)
            write-log -logfile $log -Status Info -Message "deleted encrypted file" -data "$failedfile" -progress $progress
        }catch{
            write-log -logfile $log -Status Error -Message "Could not delete encrypted file" -data "$failedfile" -progress $progress
        }
    }else{
        write-log -logfile $log -Status Info -Message "Infected file not present" -data "$failedfile" -progress $progress
    }
    }
	
}	

#remove-psdrive -name W
#remove-psdrive -name V
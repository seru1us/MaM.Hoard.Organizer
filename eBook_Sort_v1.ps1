#Take a peek inside ColdStorage, exclude organized folders starting with !
$Source ='Z:\eBooks'
$Dest ='Z:\eBooks'
$eBooks = Get-ChildItem $Source | ? {$_.FullName -notmatch '!'} | sort-object | select-object Name

#Create the hashtable for mapping categories to folder names.
#A few notes on this: The first entry is for Audiobooks, and the second is for eBooks. "61" is Comics (ignore) and 118 is Mixed Collections (too many false positives)
$dirtable = @{
"39"="!Action Adventure";
"60"="!Action Adventure";
"49"="!Art";
"71"="!Art";
"50"="!Biographical";
"72"="!Biographical";
"83"="!Business";
"90"="!Business";
"61"="";
"51"="!Computer Internet";
"73"="!Computer Internet";
"97"="!Crafts";
"101"="!Crafts";
"40"="!Crime Thriller";
"62"="!Crime Thriller";
"41"="!Fantasy";
"63"="!Fantasy";
"106"="!Food";
"107"="!Food";
"42"="!General Fiction";
"64"="!General Fiction";
"52"="!General Non Fiction";
"74"="!General Non Fiction";
"98"="!Historical Fiction";
"102"="!Historical Fiction";
"54"="!History";
"76"="!History";
"55"="!Home Garden";
"77"="!Home Garden";
"43"="!Horror";
"65"="!Horror";
"99"="!Humor";
"103"="!Humor";
"115"="!Illusion Magic";
"84"="!Instructional";
"91"="!Instructional";
"44"="!Juvenile";
"66"="!Juvenile";
"56"="!Language";
"78"="!Language";
"45"="!Literary Classics";
"67"="!Literary Classics";
"79"="!Magazines Newspapers";
"57"="!Math Science Tech";
"80"="!Math Science Tech";
"85"="!Medical";
"92"="!Medical";
"118"="";
"87"="!Mystery";
"94"="!Mystery";
"119"="!Nature";
"120"="!Nature";
"88"="!Philosophy";
"95"="!Philosophy";
"58"="!Politics Sociology Religion";
"81"="!Politics Sociology Religion";
"59"="!Recreation";
"82"="!Recreation";
"46"="!Romance";
"68"="!Romance";
"47"="!Science Fiction";
"69"="!Science Fiction";
"53"="!Self Help";
"75"="!Self Help";
"89"="!Travel Adventure";
"96"="!Travel Adventure";
"100"="!True Crime";
"104"="!True Crime";
"108"="!Urban Fantasy";
"109"="!Urban Fantasy";
"48"="!Western";
"70"="!Western";
"111"="!Young Adult";
"112"="!Young Adult"
}

#Start IE. For the first time you run this, you will have to change "visible" to "true" so you can log in. I'll automate this at some point.
$ie = new-object -ComObject "InternetExplorer.Application"
$ie.visible = $false
$ie.silent = $true

#Search MAM for category ID, and then add the category to the object
foreach ($obj in $eBooks) {

    #Remove extra crap so we can search for this file
    $lostbook = $obj.name | %{ $_.split('(')[0] } | %{ $_.split('[')[0] } | %{ $_.split('.')[0] }
    $lostbook = $lostbook.Replace("_"," ")
    Write-Host "Search term: " $lostbook

    #Search MAM
    $requestUri = "https://www.myanonamouse.net/tor/browse.php?tor%5Btext%5D=" + $lostbook + "&tor%5BsrchIn%5D%5Btitle%5D=true&tor%5BsrchIn%5D%5Bdescription%5D=true&tor%5BsrchIn%5D%5Btags%5D=true&tor%5BsrchIn%5D%5Bauthor%5D=true&tor%5BsrchIn%5D%5Bnarrator%5D=true&tor%5BsrchIn%5D%5Bseries%5D=true&tor%5BsrchIn%5D%5BfileTypes%5D=true&tor%5BsrchIn%5D%5Bfilenames%5D=true&tor%5BsearchType%5D=all&tor%5BsearchIn%5D=torrents&tor%5Bcat%5D%5B%5D=0&tor%5BbrowseFlagsHideVsShow%5D=0&tor%5Bhash%5D=&tor%5BsortType%5D=&tor%5BstartNumber%5D=0"
    $ie.navigate($requestUri)
    while($ie.Busy) { Start-Sleep -Milliseconds 100 } 
    
    #If there were no search results, add the unknown property and put it into the Uncategorized folder.   
    $mamSearchResult = $ie.Document.getElementByID('searchResults') | select textContent
    $nomamSearchResult = "Search ResultsNothing returned, out of 0"
    Start-Sleep -Milliseconds 1000

    #If there were no search results, set to "unknown" category
    if ($mamSearchResult.textContent -eq $nomamSearchResult) {
        $obj | Add-Member -NotePropertyName category -NotePropertyValue "" -Force
    }

    #Otherwise follow the link for the first result on the page
    Else {
        $catCode = $ie.Document.documentElement.getElementsByClassName('newCatLink') | select href | select-object -first 2
        $catCode = $catCode.href -replace "[^0-9]"
        
        #Replace cocurance of "Audiobooks" with "eBooks", we are doing ebooks here. Also check if it is null or not
        If ($catCode -and $catCode -ne "") {
            #If the first result is "Mixed Collections", try the second one
            If ($catCode[0] -eq "118") {
                $obj | Add-Member -NotePropertyName category -NotePropertyValue $catCode[1] -Force
            }
            Else {
            $obj | Add-Member -NotePropertyName category -NotePropertyValue $catCode[0] -Force
            }
        }
    }
    Write-Host "Filename:" $obj.Name
    Write-Host "Category:" $obj.category
}


#Now pick out the ones that we found categories for
foreach ($objj in $eBooks) {
    if ($objj.category -and $objj.category -ne "") {
        $Dir = $dirtable.Get_Item($objj.category)
        $objjName = $objj.Name

        Write-Host "Filename:" $objjName
        Write-Host "Directory:" $Dir

        #And move them
        If ($Dir -and $Dir -ne "") {
            Move-Item -LiteralPath "$Source\$objjName" -Destination "$Dest\$Dir" -force -Verbose
        }
    }
    
}


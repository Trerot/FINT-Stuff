#Requires -Modules ImportExcel
#Requires -Version 7.0
<#
    .SYNOPSIS
    creates an excel file that contains a FINT dump of: 
    all students and their relations to classes. 
    all schoolresources and their relation to schools.
    all schools and the amount of students.
    .DESCRIPTION
    does 8 larger API requests. then localy matches instead of doing a gazillon api requests which would take a long time.
    splitt into three functions
    Get-FintToken
    this function when run fills in $global:headers with an oauth token for use with api requests later on.
    You should fill inn creds on line 53-59

    Invoke-FintApiRequest
    gets all the data.

    Running Get-SchoolDataList does the stuff. 

    Data is stored in the following files once its done processing data
    "$path\UndervisningsgruppeXElev.csv"
    "$path\skole.csv"
    "$path\laerer.csv"
    "$path\Skoledataliste.xlsx"

    Default path is documents folder for the user in the powershell session.
    for another path, run "Get-SchoolDataList -path 'C:\Program Files\examplePath'" instead.

    the excel file contains the same info as the csv files split into each its own tab.
    once everything is done it will do a cleanup of the temp files. Get-ChildItem $env:TEMP\Alle*.tmp | remove-item
    if you have trouble with API stability or they just take to much time. you can remove the cleanup at the bottom of this file inside end
    and also remove the -force part of "Invoke-FintApiRequest -force" in the beginning of Get-SchoolDataList

    .EXAMPLE
        Stores info in $home\documents
    Get-SchoolDataList

        For another path than $home\documents run
    Get-SchoolDataList -path 'C:\Program Files\examplePath'

    .Notes
        FunctionName : Get-SchoolDataList
        Created by   : david.heim@mrfylke.no
        Date         : 2022-02-09
        GitHub       : https://github.com/Trerot

        Feel free to ask about stuff.
#>
function Get-FintToken {
    param (
        [switch]$Force
    )
    begin {
        #set all your credentials here
        $grant_type = "password"
        $client_id = "ID"
        $client_secret = "Secret"
        $username = "david@stuff.no"
        $password = $null
        $scope = "fint-client"
        $idp_url = "https://idp.felleskomponent.no/nidp/oauth/nam/token"
        #params for the token request.
        $parameters = @{
            uri    = $idp_url
            Method = 'Post'
            Body   = "grant_type=$grant_type&client_id=$client_id&client_secret=$client_secret&username=$username&password=$password&scope=$scope"
        }     
        #just to stop those who forget to paste in creds
        if ($password -eq $null) {
            Write-Warning "Password is blank. fill inn your api creds on line 44-50"
            break
        }
    }
    process {
        if ($force) {
            #create token
            "Asking for token."
            $Global:Token = invoke-restmethod @parameters
            $global:TokenExpireTime = (get-date).AddSeconds(($Global:token).expires_in)
            $Global:AccessToken = $global:Token.access_token
        }
        else {
            if (-not$global:Token) {
                "Cannot find Token. Asking for one now."
                #create token
                $Global:Token = invoke-restmethod @parameters
                $global:TokenExpireTime = (get-date).AddSeconds(($Global:token).expires_in)
                $Global:AccessToken = $global:Token.access_token
            }
            else {
                #check if token is to old 
                if (((get-date).AddMinutes(5)) -gt ($global:TokenExpireTime)) {
                    "Token to old. Asking for a new one."
                    #create token
                    $Global:Token = invoke-restmethod @parameters
                    #setting token expire time
                    $global:TokenExpireTime = (get-date).AddSeconds(($Global:token).expires_in)
                    $Global:AccessToken = $global:Token.access_token
                }
            }
        }
        # header for the request
        $global:headers = @{
            Authorization = "Bearer $Global:AccessToken"
        }
    }
}
function Invoke-FintApiRequest {
    param (
        [switch]$Force
    )
    begin {
        #getting oauth token. the function below fils the $global:headers which gets used further down.
        Get-FintToken
        # removing previousfiles if -force
        if($Force){
            Get-ChildItem $env:TEMP\Alle*.tmp | remove-item
        }

        #  creating an array that works with parallelized workloads.
        $Costs = [System.Collections.Concurrent.ConcurrentBag[psobject]]::new()
        
        #list all links
        $ListLinks = @(
            [pscustomobject]@{
                Name = 'AlleSkoler';
                Uri  = 'https://api.felleskomponent.no/utdanning/utdanningsprogram/skole/'
            }
            [pscustomobject]@{
                Name = 'AlleElever'; 
                Uri  = 'https://api.felleskomponent.no/utdanning/elev/elev/'
            }
            [pscustomobject]@{
                Name = 'AlleElevPersoner';
                Uri  = 'https://api.felleskomponent.no/utdanning/elev/person'
            }
            [pscustomobject]@{
                Name = 'AlleUndervisningsgrupper'; 
                Uri  = 'https://api.felleskomponent.no/utdanning/timeplan/undervisningsgruppe'
            }
            [pscustomobject]@{
                Name = 'AlleUndervisningsForhold';
                Uri  = 'https://api.felleskomponent.no/utdanning/elev/undervisningsforhold'
            }
            [pscustomobject]@{
                Name = 'AlleSkoleRessurser'; 
                Uri  = 'https://api.felleskomponent.no/utdanning/elev/skoleressurs'
            }
            [pscustomobject]@{
                Name = 'AlleFag';
                Uri  = 'https://api.felleskomponent.no/utdanning/timeplan/fag'
            }
            [pscustomobject]@{
                Name = 'AlleElevForhold'; 
                Uri  = 'https://api.felleskomponent.no/utdanning/elev/elevforhold/'
            }
            [pscustomobject]@{
                Name = 'AlleAnsatte';
                Uri  = 'https://api.felleskomponent.no/administrasjon/personal/person'
            }
            [pscustomobject]@{
                Name = 'AlleAnsattPersoner'; 
                Uri  = 'https://api.felleskomponent.no/administrasjon/personal/personalressurs'
            }
            [pscustomobject]@{
                Name = 'AlleArbeidsforhold';
                Uri  = 'https://api.felleskomponent.no/administrasjon/personal/arbeidsforhold/'
            }
            [pscustomobject]@{
                Name = 'AlleBasisGrupper'; 
                Uri  = 'https://api.felleskomponent.no/utdanning/elev/basisgruppe'
            }
            [pscustomobject]@{
                Name = 'AlleUtdanningsProgram'; 
                Uri  = 'https://api.felleskomponent.no/utdanning/utdanningsprogram/utdanningsprogram'
            }
            [pscustomobject]@{
                Name = 'AlleUtdanningsProgramOmrade'; 
                Uri  = 'https://api.felleskomponent.no/utdanning/utdanningsprogram/programomrade'
            }

        ) 
    }
    process {
        $ListLinks | ForEach-Object -ThrottleLimit 10 -Parallel {
            $item = $_
            #setting the filename for the item. 
            $FileName = "$env:TEMP\$($item.name).tmp"

            #creating an array and an object to add all filelocations to
            $ItemObject = [pscustomobject]@{
                Name     = "$($item.Name)"; 
                Uri      = "$($item.Uri)"
                FileName = $FileName
            }
            $fintfilearray = $using:costs
            $fintfilearray.add($ItemObject)
    
            #testing if file exists or if date is older than 8 hours.
            if (!(test-path $filename) -or (Get-ChildItem $FileName -ErrorAction SilentlyContinue).LastWriteTime -le (get-date).AddHours(-8) -or (Get-ChildItem -Path $filename -ErrorAction SilentlyContinue).Length -eq "0" ) {
                try {
                    #New-Item -Path $FileName -ItemType File -Force
                    $countist = 0
                    while ((get-childitem -path $filename -ErrorAction SilentlyContinue).length -eq "0") {
                        $countist ++
                        "try $countist for $filename"
                        Invoke-RestMethod -Headers $using:headers -uri $($item.uri) -ErrorAction Stop -OutFile $FileName
                    }
                }
                catch {
                    $errormessage = ($error[0] | Select-Object *).Exception.Response.statuscode
                    $autherror = "Unauthorized"
                    $wronglink = "NotFound"
                    $none = "none" # not sure if this is what it lookslike. status code is probably blank when i get non error.
    
                    switch ($errormessage) {
                        $autherror { 
                            Get-FintToken 
                        }
                        $wronglink {
                            "link is wrong, fix it"
                        }
                        $none {
                            start-sleep -Seconds 10
                            # then try again after sleep.
                        }
                        Default { <#this should just try again #> }
                    }
                }
            }
        }
        $Global:FintFileArray = $Costs
    }
}
function Get-SchoolDataList {
    [CmdletBinding()]
    param (
        
    )
    
    begin {

        Invoke-FintApiRequest -force
        # import all the stuff into arrays. 
        $Global:FintFileArray.foreach({
                $FintFile = $_
                $info = get-content -Path $FintFile.FileName | ConvertFrom-Json
                $hashtable = @{}
                $info._embedded._entries.foreach({
                        $item = $_
                        $key = $null
                        # testing for systemid
                        $key = ($item._links.self.href | select-string "systemid").line

                        if ($null -eq $key  ) {
                            # all the ones without system id
                            $key = ($item._links.self.href | select-string "fodselsnummer").line
                            if ($null -eq $key ) {
                                $key = ($item._links.self.href | select-string "ansattnummer").line
                                $hashtable.add($key, $item)
                            }
                            else {
                                $hashtable.add($key, $item)
                            }
                        }
                        else {
                            $hashtable.add($key, $item)
                        }          
                    })

                New-Variable -Name "$($FintFile.name)" -Value $info -Force
                New-Variable -Name "Hash$($FintFile.name)" -Value $hashtable -Force
            })
        #arraylist to sort the students into 
        $ElevInElevPersoner = New-Object -TypeName System.Collections.ArrayList
        #arraylist for the donestuff
        $DoneStuff = New-Object -TypeName System.Collections.ArrayList
        $item = $null
        # creating an arraylist to store all stuff in.
        $UndervisningsgruppeXElevArrayList = New-Object -TypeName System.Collections.ArrayList
        $SkoleArrayList = New-Object -TypeName System.Collections.ArrayList
        $LaererArrayList = New-Object -TypeName System.Collections.ArrayList
    }
    
    process {
        "Local processing data for UndervisningsgruppeXElev. Takes some minutes."
        $AlleElevForhold._embedded._entries.foreach({
                $studentprogramomrade = $null
                $utdanningstrinn = $null
                $studentelevforhold = $null
                $numFromProgramomrade = $null
                $utdanningstrinn = $null
                $programomrade = $null
                $programnavn = $null
                $basisgruppe = $null

                $StudentClassesLink = $_._links.undervisningsgruppe
    
                $studentsystemid = $_._links.elev.href
                # $StudentInfo = $AlleElever._embedded._entries.where({ $_._links.self[2].href -eq $StudentLink.href })
                $studentinfo = $hashAlleElever.$studentsystemid
                # some students have multiple school relations. should get the one where hovedskole = true
                if ($studentinfo._links.elevforhold.href.count -gt 1) {
                    if ($HashAlleElevForhold.$($studentinfo._links.elevforhold.href[0]).hovedskole -eq $true ) {
                        $studentelevforhold = $HashAlleElevForhold.$($studentinfo._links.elevforhold.href[0])
                    }
                    if ($HashAlleElevForhold.$($studentinfo._links.elevforhold.href[1]).hovedskole -eq $true ) {
                        $studentelevforhold = $HashAlleElevForhold.$($studentinfo._links.elevforhold.href[1])
                    }
                }
                else {
                    $studentelevforhold = $HashAlleElevForhold.$($studentinfo._links.elevforhold.href)
                }

                $studentprogramomrade = $HashAlleUtdanningsProgramOmrade.$($studentelevforhold._links.programomrade.href)
                $basisgruppe = $HashAlleBasisGrupper.$($studentelevforhold._links.basisgruppe.href)
                
                
                $utdanningstrinn = $basisgruppe._links.trinn.href -replace '.*/'


                $programomrade = $studentprogramomrade.beskrivelse
                $programnavn = $studentprogramomrade.navn
                $feidenavn = $studentinfo.feidenavn.identifikatorverdi
                $elevnummer = $StudentInfo.elevnummer.identifikatorverdi
    
                # $elevpersoninfo = $AlleElevPersoner._embedded._entries.where({ $_._links.self.href -eq $studentinfo._links.person.href })
                $elevpersoninfo = $HashAlleElevPersoner."$($studentinfo._links.person.href)"
                $elevfornavn = $elevpersoninfo.navn.fornavn + " " + $elevpersoninfo.navn.mellomnavn
                $elevetternavn = $elevpersoninfo.navn.etternavn
    
                $StudentClassesLink.foreach({
                        $class = $_
                        # $undervisningsgruppe = $alleundervisningsgrupper._embedded._entries.where({ $_._links.self.href -eq $class.href })
                        $undervisningsgruppe = $HashAlleUndervisningsgrupper."$($class.href)"
                        # you can have more than one undervisningsforhold. this selects the first hit and bypasses "null array" errors.
                        if ($undervisningsgruppe._links.undervisningsforhold.href.count -gt 1) {
                            # $undervisningsforhold = $AlleUndervisningsForhold._embedded._entries.where({ $_._links.self.href -eq $undervisningsgruppe._links.undervisningsforhold.href[0] })
                            $undervisningsforhold = $HashAlleUndervisningsForhold."$($undervisningsgruppe._links.undervisningsforhold.href[0])"
                        }
                        else {
                            # $undervisningsforhold = $AlleUndervisningsForhold._embedded._entries.where({ $_._links.self.href -eq $undervisningsgruppe._links.undervisningsforhold.href })
                            $undervisningsforhold = $HashAlleUndervisningsForhold."$($undervisningsgruppe._links.undervisningsforhold.href)"
                        }
                        # $skoleressurs = $AlleSkoleRessurser._embedded._entries.where({ $_._links.self.href[1] -eq $undervisningsforhold._links.skoleressurs.href })
                        $skoleressurs = $HashAlleSkoleRessurser."$($undervisningsforhold._links.skoleressurs.href)"
                        $teachername = $skoleressurs.feidenavn.identifikatorverdi
                        $classname = $undervisningsgruppe.navn
                        # $classfullname = ($AlleFag._embedded._entries.where({ $_._links.self.href -eq $undervisningsgruppe._links.fag.href })).navn
                        $Fag = $hashallefag."$($undervisningsgruppe._links.fag.href)"
                        $fagkode = $fag.systemid.identifikatorverdi
                        $classfullname = $Fag.navn

                        #finding programmområde and 
                        #$ProgramOmrade = $hashAlleUtdanningsProgramOmrade."$($fag._links.programomrade)"
    
    
                        # $schoolinfo = $alleskoler._embedded._entries.where({ $_._links.self[1].href -eq $undervisningsgruppe._links.skole.href })
                        $SchoolInfo = $HashAlleSkoler."$($undervisningsgruppe._links.skole.href)"
                        $SchoolName = $schoolinfo.navn
                        #adding to list
                        [void]$UndervisningsgruppeXElevArrayList.add("$schoolname;$classname;$classfullname;$feidenavn;$elevnummer;$elevetternavn;$elevfornavn;$teachername;$programomrade;$programnavn;$utdanningstrinn;$fagkode")
                    
                    })
            })

        #export to excel here
        $StudentPSobject = $UndervisningsgruppeXElevArrayList | ConvertFrom-Csv -Delimiter ";" -Header "skolenavn", "undervisningsgruppenavn", "undervisningsgruppebeskrivelse", "ElevFeideNavn", "Elevnummer", "ElevEtternavn", "ElevFornavn", "LærerFeideNavn", "Programområde", "programnavn", "utdanningstrinn", "Fagkode"
        $date = (get-date).ToFileTime()
        $StudentPSobject | Export-Excel -Path "$env:USERPROFILE\schooldatalist\Skoledatalist-$date.xlsx" -WorksheetName "UndervisningsgruppeXElev"
    }
    
    end {
        #cleaning up temp files
        Get-ChildItem $env:TEMP\Alle*.tmp | remove-item
    }
}
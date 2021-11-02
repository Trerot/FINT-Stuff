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
    splitt into two functions
    Get-FintToken
    this function when run fills in $global:headers with an oauth token for use with api requests later on.
    You should fill inn creds on line 45-50

    Running Get-SchoolDataList does the stuff. 

    Data is stored in the following files once its done processing data
    "$path\UndervisningsgruppeXElev.csv"
    "$path\skole.csv"
    "$path\laerer.csv"
    "$path\Skoledataliste.xlsx"

    Default path is documents folder for the user in the powershell session.
    for another path, run "Get-SchoolDataList -path 'C:\Program Files\examplePath'" instead.

    the excel file contains the same info as the csv files split into each its own tab.

    .EXAMPLE
        Stores info in $home\documents
    Get-SchoolDataList

        For another path than $home\documents run
    Get-SchoolDataList -path 'C:\Program Files\examplePath'

    .Notes
        FunctionName : Get-SchoolDataList
        Created by   : david.heim@mrfylke.no
        Date         : 2022-11-02
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
function Get-SchoolDataList {
    param (
        # Default path is home folder for the user. 
        [string]$path = "$home\Documents"
    )
    begin {
        Get-FintToken -Force
        #testing for files
        if (test-path "$path\UndervisningsgruppeXElev.csv") {
            Write-Warning "$path\UndervisningsgruppeXElev.csv exists."
            $test = $true
        }
        if (test-path "$path\skole.csv") {
            Write-Warning "$path\skole.csv exists."
            $test = $true
        }
        if (test-path "$path\laerer.csv") {
            Write-Warning "$path\laerer.csv exists."
            $test = $true
        }
        if (test-path "$path\Skoledataliste.xlsx") {
            Write-Warning "$path\Skoledataliste.xlsx exists."
            $test = $true
        }
        if ($test) {
            "delete these files before you continue"
            Read-Host -Prompt "press Y and  enter to continue once you have deleted the files."
        }
        # All API requests and hash table creations(for speed)
        "Getting Data"
        $AlleElever = Invoke-RestMethod -Headers $headers -uri https://api.felleskomponent.no/utdanning/elev/elev/
        $hashAlleElever = @{}
        $AlleElever._embedded._entries.foreach({ $hashAlleElever.add($_.systemid.identifikatorverdi, $_) })
        start-sleep -Seconds 3
        "1/8"
        $AlleSkoler = Invoke-RestMethod -Headers $headers -uri https://api.felleskomponent.no/utdanning/utdanningsprogram/skole/
        $HashAlleSkoler = @{}
        $AlleSkoler._embedded._entries.foreach({ $HashAlleSkoler.add($_._links.self.href[1], $_) })
        start-sleep -Seconds 3
        "2/8"
        $AlleElevPersoner = Invoke-RestMethod -Headers $headers -uri https://api.felleskomponent.no/utdanning/elev/person
        $HashAlleElevPersoner = @{}
        $AlleElevPersoner._embedded._entries.foreach({ $HashAlleElevPersoner.add($_._links.self.href, $_) })
        start-sleep -Seconds 3
        "3/8"
        $alleundervisningsgrupper = Invoke-RestMethod -Headers $headers -uri https://api.felleskomponent.no/utdanning/timeplan/undervisningsgruppe
        $HashAlleUndervisningsgrupper = @{}
        $alleundervisningsgrupper._embedded._entries.foreach({ $HashAlleUndervisningsgrupper.add($_._links.self.href, $_) })
        start-sleep -Seconds 3
        "4/8"
        $AlleUndervisningsForhold = Invoke-RestMethod -Headers $headers -uri https://api.felleskomponent.no/utdanning/elev/undervisningsforhold
        $HashAlleUndervisningsForhold = @{}
        $AlleUndervisningsForhold._embedded._entries.foreach({ $HashAlleUndervisningsForhold.add($_._links.self.href, $_) })
        start-sleep -Seconds 3
        "5/8"
        $AlleSkoleRessurser = Invoke-RestMethod -Headers $headers -uri https://api.felleskomponent.no/utdanning/elev/skoleressurs
        $HashAlleSkoleRessurser = @{}
        $AlleSkoleRessurser._embedded._entries.foreach({ $HashAlleSkoleRessurser.add("https://api.felleskomponent.no/utdanning/elev/skoleressurs/systemid/$($_.systemid.identifikatorverdi)", $_) })
        start-sleep -Seconds 3
        "6/8"
        $AlleFag = Invoke-RestMethod -Headers $headers -uri https://api.felleskomponent.no/utdanning/timeplan/fag
        $HashAlleFag = @{}
        $AlleFag._embedded._entries.foreach({ $HashAlleFag.add($_._links.self.href, $_) })
        start-sleep -Seconds 3
        "7/8"

        $AlleElevForhold = Invoke-RestMethod -Headers $headers -Uri https://api.felleskomponent.no/utdanning/elev/elevforhold/
        start-sleep -Seconds 3
        "8/8"
        # dont need these so commented out
        # $AlleAnsatte = Invoke-RestMethod -Headers $headers -uri https://api.felleskomponent.no/administrasjon/personal/person
        # $AlleArbeidsforhold = Invoke-RestMethod -Headers $headers -uri https://api.felleskomponent.no/administrasjon/personal/arbeidsforhold/
        
        # creating an arraylist to store all stuff in.
        $UndervisningsgruppeXElevArrayList = New-Object -TypeName System.Collections.ArrayList
        $SkoleArrayList = New-Object -TypeName System.Collections.ArrayList
        $LaererArrayList = New-Object -TypeName System.Collections.ArrayList
    }
    process {
        "Local processing data for UndervisningsgruppeXElev. Takes some minutes."
        $AlleElevForhold._embedded._entries.foreach({
                $StudentClassesLink = $_._links.undervisningsgruppe
    
                $studentsystemid = $_._links.elev.href -replace "https://api.felleskomponent.no/utdanning/elev/elev/systemid/"
                # $StudentInfo = $AlleElever._embedded._entries.where({ $_._links.self[2].href -eq $StudentLink.href })
                $studentinfo = $hashAlleElever.$studentsystemid
    
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
                        $classfullname = $Fag.navn
    
    
                        # $schoolinfo = $alleskoler._embedded._entries.where({ $_._links.self[1].href -eq $undervisningsgruppe._links.skole.href })
                        $SchoolInfo = $HashAlleSkoler."$($undervisningsgruppe._links.skole.href)"
                        $SchoolName = $schoolinfo.navn
    
                        [void]$UndervisningsgruppeXElevArrayList.add("$schoolname;$classname;$classfullname;$feidenavn;$elevnummer;$elevetternavn;$elevfornavn;$teachername")
                    })
            })
        "Done"
        "Local processing data for Skole. should be almost instant"
        foreach ($item in $AlleSkoler._embedded._entries) {
            $SchoolName = $item.navn
            $SchoolNumber = $item.skolenummer.identifikatorverdi
            $SchoolStudentCount = ($item._links.elevforhold | measure-object).Count
            [void]$SkoleArrayList.add("$SchoolName;$SchoolNumber;$SchoolStudentCount") 
        }
        "done"
        "Local processing data for Laerer"
        $AlleSkoleRessurser._embedded._entries.foreach({
                $item = $_
                $TeacherPerson = $HashAlleElevPersoner."$($item._links.person.href)"
                $teacherSchool = $HashAlleSkoler."$($item._links.skole.href)"
                $TeacherFirstName = "$($TeacherPerson.navn.fornavn)" + "$($TeacherPerson.navn.mellomnavn)"
                $TeacherLastName = $TeacherPerson.navn.etternavn
                $TeacherSchoolName = $teacherSchool.navn
                $feidenavn = $item.feidenavn.identifikatorverdi
                $TeacherEmail = $TeacherPerson.kontaktinformasjon.epostadresse
                [void]$LaererArrayList.add("$TeacherFirstName;$TeacherLastName;$TeacherSchoolName;$feidenavn;$TeacherEmail")
            })
        "Data processed. converting to object with headers."
        # should proapply just have created psobject to begin with, if i feel like it i/someone feels like it i should run some performance tests and clean up some of this.
        $StudentPSobject = $UndervisningsgruppeXElevArrayList | ConvertFrom-Csv -Delimiter ";" -Header "skolenavn", "undervisningsgruppenavn", "undervisningsgruppebeskrivelse", "ElevFeideNavn", "Elevnummer", "ElevEtternavn", "ElevFornavn", "LærerFeideNavn"
        $SkolePSobject = $SkoleArrayList | ConvertFrom-Csv -Delimiter ";" -Header "Skolenavn", "Skolenummer", "Elevantall"
        $LaererPSobject = $LaererArrayList | ConvertFrom-Csv -Delimiter ";" -Header "Fornavn", "Etternavn", "Skole", "FeideID", "E-post"
        "Exporting to excel and CSV. following filenames"
        "$path\UndervisningsgruppeXElev.csv"
        "$path\skole.csv"
        "$path\laerer.csv"
        "$path\Skoledataliste.xlsx"
        #CSV files
        $StudentPSobject | Export-Csv -path $path\UndervisningsgruppeXElev.csv -force
        $SkolePSobject | Export-Csv -Path $path\skole.csv -Force
        $LaererPSobject | export-csv -Path $path\laerer.csv -Force
        # adding to excel
        $StudentPSobject | Export-Excel -Path $path\Skoledataliste.xlsx -WorksheetName "UndervisningsgruppeXElev"
        $SkolePSobject | Export-Excel -Path $path\Skoledataliste.xlsx -WorksheetName "Skole"
        $LaererPSobject | Export-Excel -path $path\Skoledataliste.xlsx -WorksheetName "Laerer"
        "done"
    }
    end {
        # nothing to do here realy as far as i can tell.
    }
}
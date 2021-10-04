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
    In order for this to work you should have an oathtoken on hand stored in $accesstoken which in turn is stored in $headers like so

$AccessToken = "Paste in oath token here"
$headers = @{
Authorization = "Bearer $AccessToken"
}
       
    Example to fill in $headers with powershell using Curl (https://curl.se/windows/)

$CurlCommand = @"
curl -s https://idp.felleskomponent.no/nidp/oauth/nam/token ``
-u "somestring-off-text-with-stuff:longeruninterupetstringoftext" ``
-d grant_type=password ``
-d username="davidheim@client.mrfylke.no" ``
-d password="shouldhaveyourpasswordhere" ``
-d scope="fint-client"
"@

$Token = Invoke-Expression $CurlCommand | ConvertFrom-Json
$AccessToken = $Token.access_token
$headers = @{
    Authorization = "Bearer $AccessToken"
}

    Data is stored in the following files once its done processing data
    "C:\UndervisningsgruppeXElev.csv"
    "C:\skole.csv"
    "C:\laerer.csv"
    "C:\Skoledataliste.xlsx"

    the excel file contains the same info as the csv files split into each its own tab.

    .EXAMPLE
    Get-SchoolDataList

    .Notes
        FunctionName : Get-SchoolDataList
        Created by   : david.heim@mrfylke.no
        Date         : 04/10/2022 13:43
        GitHub       : https://github.com/Trerot

        Feel free to ask about stuff.
#>
function Get-SchoolDataList {
    begin {
        if ($headers -eq $null) {
            $string = @'
$headers is empty!
you need to have an oauth token in $headers for this to run
I couldn't get oauth token through powershell only so you have to install curl for windows(if its not allready installed) for this to work 
https://curl.se/windows/.  
you could also just do $accesstoken = "paste inn your token here"
then remove all the other stuff apart from the $headers part.

   Example to fill in $headers 

$CurlCommand = @"
curl -s https://idp.felleskomponent.no/nidp/oauth/nam/token ``
-u "somestring-off-text-with-stuff:longeruninterupetstringoftext" ``
-d grant_type=password ``
-d username="davidheim@client.mrfylke.no" ``
-d password="shouldhaveyourpasswordhere" ``
-d scope="fint-client"
"@

$Token = Invoke-Expression $CurlCommand | ConvertFrom-Json
$AccessToken = $Token.access_token
$headers = @{
    Authorization = "Bearer $AccessToken"
}
'@
            Write-Warning $string
            break
        }
        #testing for files
        if (test-path "C:\UndervisningsgruppeXElev.csv") {
            Write-Warning "C:\UndervisningsgruppeXElev.csv exists."
            $test = $true
        }
        if (test-path "C:\skole.csv") {
            Write-Warning "C:\skole.csv exists."
            $test = $true
        }
        if (test-path "C:\laerer.csv") {
            Write-Warning "C:\laerer.csv exists."
            $test = $true
        }
        if (test-path "C:\Skoledataliste.xlsx") {
            Write-Warning "C:\Skoledataliste.xlsx exists."
            $test = $true
        }
        if ($test) {
            "delete these files before you continue"
            Read-Host -Prompt "press Y and  enter to continue once you have deleted the files."
        }
        
        # All API requests and hash table creations(for speed)
        "Getting Data"
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
        $StudentPSobject = $UndervisningsgruppeXElevArrayList | ConvertFrom-Csv -Delimiter ";" -Header "skolenavn", "undervisningsgruppenavn", "undervisningsgruppebeskrivelse", "ElevFeideNavn", "Elevnummer", "ElevEtternavn", "ElevFornavn", "LærerFeideNavn"
        $SkolePSobject = $SkoleArrayList | ConvertFrom-Csv -Delimiter ";" -Header "Skolenavn", "Skolenummer", "Elevantall"
        $LaererPSobject = $LaererArrayList | ConvertFrom-Csv -Delimiter ";" -Header "Fornavn", "Etternavn", "Skole", "FeideID", "E-post"
        "Exporting to excel and CSV. following filenames"
        "C:\UndervisningsgruppeXElev.csv"
        "C:\skole.csv"
        "C:\laerer.csv"
        "C:\Skoledataliste.xlsx"
        #CSV files
        $StudentPSobject | Export-Csv -path C:\UndervisningsgruppeXElev.csv -force
        $SkolePSobject | Export-Csv -Path C:\skole.csv -Force
        $LaererPSobject | export-csv -Path c:\laerer.csv -Force
        # adding to excel
        $StudentPSobject | Export-Excel -Path C:\Skoledataliste.xlsx -WorksheetName "UndervisningsgruppeXElev"
        $SkolePSobject | Export-Excel -Path C:\Skoledataliste.xlsx -WorksheetName "Skole"
        $LaererPSobject | Export-Excel -path C:\Skoledataliste.xlsx -WorksheetName "Laerer"
        "done"
    }
    end {
        # nothing to do here realy as far as i can tell.
    }
}
Function Import-BulkeDocs {
    [cmdletbinding()]

    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet('EDOCS', 'EDOCS_TEST')]
        [string]$Server,

        [Parameter(Mandatory = $true)]
        [String]$Log
    )
    #region Script setup
    Import-Module "\\Path\To\API\eDOCSAPI.psm1" -Force
    $csvLoc = "\\Path\To\Document.csv"
    $CSV = import-csv $csvLoc
    $taskID = 0

    Add-eDOCSScriptLogEntry -ErrMsg "Import started" -Status Info -Severity INFO -Type Script -Action "Import" -Resource "$csvLoc"
    #endregion Script setup

    #region Login Section
    $user = 'eDOCS_User'
    $password = 'Password'

    # Sometimes logging into EDOCS_TEST fails, who knows why, this gives us a few goes to get in.
    $loginAttempts = 0
    $maxLoginAttempts = 20
    $loggedIn = $false

    while (-not $loggedIn -and $loginAttempts -lt $maxLoginAttempts) {
        $login = Get-eDOCSLogon -dmUser $user -dmPass $password -dmLibrary $Server -logPath $Log

        if ($login.success) {
            $loggedIn = $true
        } else {
            return
        }
        $loginAttempts++
    }
    #endregion Login Section
    
    #region Extensions
    $extToAppID = @{
        ".mdb"  = "MS ACCESS";
        ".PDF"  = "PDF";
        ".DOT"  = "MS WORD";
        ".DOCX" = "MS WORD";
        ".DOC"  = "MS WORD";
        ".ODT"  = "MS WORD";
        ".RTF"  = "MS WORD";
        ".WBK"  = "MS WORD";
        ".docm" = "MS WORD";
        ".XLS"  = "MS EXCEL";
        ".XLSX" = "MS EXCEL";
        ".CSV"  = "MS EXCEL";
        ".xlsm" = "MS EXCEL";
        ".MSG"  = "MS OUTLOOK";
        ".PPT"  = "MS POWERPOINT";
        ".PPTX" = "MS POWERPOINT";
        ".PUB"  = "MS PUBLISHER";
        ".TIF"  = "PUB TIF";
        ".GIF"  = "PUB GIF";
        ".TIFF"  = "PUB TIF";
        ".JPG"  = "PUB JPG";
        ".JPEG" = "PUB JPG";
        ".PNG"  = "PUB PNG";
        ".TXT"  = "NOTEPAD";
        ".ZIP"  = "WINDOWS ZIP";
        ".HTM"  = "PUB BROWSER";
        ".HTML" = "PUB BROWSER";
        ".XML"  = "PUB XML";
        ".MOV"  = "MPEG";
        ".MP3"  = "MPEG";
        ".MP4"  = "MPEG";
        ".MPG"  = "MPEG";
        ".WMA"  = "WMV";
        ".BMP"  = "PUB BMP";
        ".WAV"  = "WMV";
        ".ASF"  = "WMV";
        ".AVI"  = "WMV";
        ".WMV"  = "WMV";
        ".SWF"  = "PUB BROWSER";
        ".MHT"  = "PUB BROWSER";
        ".INDD" = "ADOBE INDESIGN";
        ".DWG" = "AUTOCAD";
        ".REP" = "BIQ REPORTS";
        ".CIT" = "CITERITE";
        ".WDF" = "DELTAVIEW";
        ".DCM" = "DICOM";
        ".EML" = "EML";
        ".TOA" = "FAWIN";
        ".ZZZ" = "FOLDER";
        ".KML" = "GOOGLEEARTH";
        ".123" = "L123-97";
        ".LND" = "LOTUS NOTES DOC";
        ".LNE" = "LOTUS NOTES EMAIL";
        ".DXL" = "LOTUS NOTES FORM";
        ".LWP" = "LOTUS WORD PRO";
        ".MPEG" = "MPEG";
        ".MPP" = "MS PROJECT";
        ".SHW" = "PRESENTATIONS";
        ".XER" = "PRIMAVERA";
        ".PST" = "PST";
        ".QPW" = "QPW";
        ".QRP" = "RM VIEW";
        ".VSD" = "VISIO";
        ".WPP" = "WEB PUBLISHING";
        ".WPD" = "WORDPERFECT";
        ".dif" = "MS EXCEL";
        ".ods" = "MS EXCEL";
        ".prn" = "MS EXCEL";
        ".slk" = "MS EXCEL";
        ".xla" = "MS EXCEL";
        ".xlam" = "MS EXCEL";
        ".xlsb" = "MS EXCEL";
        ".xlt" = "MS EXCEL";
        ".xltm" = "MS EXCEL";
        ".xltx" = "MS EXCEL";
        ".xps" = "MS EXCEL";
        ".emf" = "MS POWERPOINT";
        ".odp" = "MS POWERPOINT";
        ".potm" = "MS POWERPOINT";
        ".potx" = "MS POWERPOINT";
        ".ppa" = "MS POWERPOINT";
        ".ppam" = "MS POWERPOINT";
        ".pps" = "MS POWERPOINT";
        ".ppsm" = "MS POWERPOINT";
        ".ppsx" = "MS POWERPOINT";
        ".pptm" = "MS POWERPOINT";
        ".thmx" = "MS POWERPOINT";
        ".wmf" = "MS POWERPOINT";
        ".dotm" = "MS WORD";
        ".dotx" = "MS WORD";
        ".wps" = "MS WORD";
        ".vsdx" = "VISIO";
    }
    #endregion Extensions
    
    #region Import checks
    Foreach($Document in $csv) {
        $appID = $null
        $appID = $extToAppID["$($Document.Extension)"]
    
        If($appID) {
            If(($Document.DateOfRequest).Length -eq 7) {
                $DateOfRequest = [DateTime]::ParseExact("0$($Document.DateOfRequest)","MMddyyyy",$null).ToString("dd/MM/yyyy")
            } Elseif(($Document.DateOfRequest).Length -eq 8) {
                $DateOfRequest = [DateTime]::ParseExact("$($Document.DateOfRequest)","MMddyyyy",$null).ToString("dd/MM/yyyy")
            } Else {
                $DateOfRequest = "N/A"
            }
            If(($Document.DateFinalised).Length -eq 7) {
                $DateFinalised = [DateTime]::ParseExact("0$($Document.DateFinalised)","MMddyyyy",$null).ToString("dd/MM/yyyy")
            } Elseif(($Document.DateFinalised).Length -eq 8) {
                $DateFinalised = [DateTime]::ParseExact("$($Document.DateFinalised)","MMddyyyy",$null).ToString("dd/MM/yyyy")
            } Else {
                $DateFinalised = "N/A"
            }
            If(($Document.DateRecievedDept).Length -eq 7) {
                $DateRecieved = [DateTime]::ParseExact("0$($Document.DateRecievedDept)","MMddyyyy",$null).ToString("dd/MM/yyyy")
            } Elseif(($Document.DateRecievedDept).Length -eq 8) {
                $DateRecieved = [DateTime]::ParseExact("$($Document.DateRecievedDept)","MMddyyyy",$null).ToString("dd/MM/yyyy")
            } Else {
                $DateRecieved = "N/A"
            }
            $Abstract = Set-eDOCSProfileEscapeCharacters "Description goes here: $($Document.NewName)`r`nExtension of document type stuff: $($Document.Extension)" 

            $properties = @{
                "DOCNAME"        = $($Document.NewName);
                "TYPE_ID"        = "DEFAULT";
                "FILE_EXTENSION" = $($Document.Extension);
                "APP_ID"         = $appID;
                "AUTHOR_ID"      = $user;
                "TYPIST_ID"      = $user;
                "PD_FILEPT_NO"   = $($Document.PartNo);
                "ABSTRACT"       = $Abstract
            }
            $params = @{
                "taskID"     = $taskID++
                "dmDST"      = $($login.dst);
                "library"    = $Server;
                "logPath"    = $Log;
                "form"       = "KMC_ELEC_DOC";
                "filepath"   = $($Document.FullName);
                "properties" = $properties;
            }
            Import-eDOCSDocument @params

        } Else {
            "FAILED: No AppID for $($Document.Fullname)" | Out-File $Log -Append
        }
    }
    Add-eDOCSScriptLogEntry -ErrMsg "Import Completed" -Status Info -Severity INFO -Type Script -Action "Import" -Resource "$csvloc"
}
Import-BulkeDocs -Server EDOCS -Log '\\Path\To\Logs.txt'

<#
Using this code for the profiles Description
$Abstract = Set-eDOCSProfileEscapeCharacters "DocID: $($Document.DocID)`r`nCross Reference: $($Document.CrossReferences)`r`nSubject: $($Document.Subject)`r`nTopic: $($Document.Topic)`r`nBranch: $($Document.Branch)`r`nAccountable Area: $($Document.AccountableArea)`r`nType: $($Document.Type)`r`nSub Type: $($Document.SubType)`r`nAddress: $($Document.Address)`r`nAuthor Type: $($Document.AuthorType)`r`nAuthor: $($Document.Author)`r`nResponse Author: $($Document.ResponseAuthor)`r`nResponse Bch: $($Document.ResponseBch)`r`nDate Of Request: $($DateOfRequest)`r`nDate Recieved Dept: $($DateRecieved)`r`nDate Finalised: $($DateFinalised)`r`nAdditional Comments: $($Document.AdditionalComments)" 

Would result in this on the actual profile.
###########
DocID: MC2058
Cross Reference: ME/08/3003
Subject: Letter from Some Dude concerning overhead powerline
Topic: Energy Initiatives ()
Branch: Industry and Client Services
Accountable Area: Energy Sector Monitoring
Type: Ministerial correspondence
Sub Type: For reply under Senior Policy Advisor's signature
Address: ADDRESS EXAMPLE
Author Type: Member of the Public
Author: FAKE NAME1
Response Author: FAKE NAME2
Response Bch: Industry and Client Services
Date Of Request: 03/09/2008
Date Recieved Dept: 05/09/2008
Date Finalised: 14/10/2008
Additional Comments: N/A
###########

#### CSV Example - Only 1 row ####
COLUMN HEADER      : ROW VALUE
FullName           : \\Document\Location\MC2058-0.pdf
Name               : MC2058-0.pdf
NewName            : MC2058-0
Extension          : .pdf
DocID              : MC2058
CrossReferences    : ME/08/3003
Subject            : Letter from Some Dude concerning overhead powerline
Topic              : Energy Initiatives ()
Branch             : Industry and Client Services
AccountableArea    : Energy Sector Monitoring
Type               : Ministerial correspondence
SubType            : For reply under Senior Policy Advisor's signature
Address            : ADDRESS EXAMPLE
AuthorType         : Member of the Public
Author             : FAKE NAME1
ResponseAuthor     : FAKE NAME2
ResponseBch        : Industry and Client Services
DateOfRequest      : 9032008
DateFinalised      : 10142008
DateRecievedDept   : 9052008
AdditionalComments : N/A
PartNo             : 053/0005742-003(E)
ImportNo           : 78433
#>

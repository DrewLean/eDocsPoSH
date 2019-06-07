# This library is designed to make working with the eDOCS DM API easier.
# For more information about the eDOCS DM API please see:
# eDOCS DM 5.3.1 DM API Overview Guide
# eDOCS DM 5.3.1 DM API Reference Guide
# eDOCS DM 5.3.1 DM Extensions API Programmers Guide
# 
# 
# Functions in this library will only work on computers that have eDOCS DM
# installed.
# 
# With the exception of adding a log entry or logging on to eDOCS all functions
# in this library require a security token which can be obtained by logging
# into eDOCS using the Get-eDOCSLogon function.
# 
# Most of the functions in this library frequently check the ErrNumber property
# of DM API objects. This is because the eDOCS DM API is old school and tends
# not to throw errors rather what you wanted to happen doesn't and the object
# that you used to try to make it happen has its ErrNumber set to a value other
# than 0 and ErrDescription set to a sometimes useful error message.
# 
# The results of most of the eDOCS API function calls are redirected to $null
# because they will return a value typically 0 which gets passed down the
# pipeline as part of the return value of the function that the API call is
# made within.

function Add-eDOCSScriptLogEntry {
    <#
    .SYNOPSIS
    Uploading a document to eDOCS with the eDOCS DM API is a very labour
    intensive task with no bulit in error checking this function ensures that
    manually created log entries are time stamped and have a consistent layout.
    
    .DESCRIPTION
    Uploading a document to eDOCS with the eDOCS DM API is a very labour
    intensive task with no built in error checking this function ensures that
    manually created log entries are time stamped and have a consistent layout.
    
    .PARAMETER msg
    [LEGACY] String representing the log message to be added to the log.
    
    .PARAMETER logPath
    [LEGACY] Path to the log file.

    .PARAMETER ErrMsg
    String representing the log message to be added to the log.

    .PARAMETER Status
    'Failed', 'Succeeded', 'Info', depending on the outcome of the script.

    .PARAMETER Severity
    'LOW', 'MEDIUM', 'HIGH', depending on the impact of the script.
    
    .EXAMPLE
    While transitioning, scripts will function as normal. Once I transition the scripts, new line will be as below.
    Add-eDOCSScriptLogEntry -ErrMsg "New message type" -Status Succeeded -Severity LOW -Type Script
    Preffered type of message would be as per below.

    [Action] [item] [to/from] [location] - [optional code/task]
    -ErrMsg "Importing $file into $Server - $taskID"
    -ErrMsg "Logging into $user on $library"

    The Status will either succeed, fail or info. No need to include this in your error message.
    
    [LEGACY]
    $taskId = 5
    $docnumber = Import-eDOCSDocument @parms # lets assume $docnumber comes back as 12345
    Add-eDOCSScriptLogEntry -msg "$taskId; $docnumber; Document Uploaded successfully" -path $logPath

    Would add the following line to the text file referenced by $logPath
    28/06/2018 8:50:44 AM - 5; 12345; Document Uploaded successfully
    
    .NOTES
    #>

    Param(
        [parameter(Mandatory = $false, Position = 0)]
        [String]$msg,
        [parameter(Mandatory = $false, Position = 1)]
        [String]$logPath,
        [parameter(Mandatory = $false, Position = 2)]
        [String]$ErrMsg,
        [parameter(Mandatory = $false, Position = 3)]
        [ValidateSet('Failed', 'Succeeded', 'Info')]
        [String]$Status,
        [Parameter(Mandatory = $false, Position = 4)]
        [ValidateSet('LOW', 'MEDIUM', 'HIGH', 'INFO')]
        [string]$Severity,
        [Parameter(Mandatory = $false, Position = 5)]
        [ValidateSet('Script', 'SeeUnity', 'Server')]
        [string]$Type,
        [parameter(Mandatory = $false, Position = 2)]
        [String]$Action,
        [parameter(Mandatory = $false, Position = 2)]
        [String]$Resource
    )
 
# SPLUNK command to be implemented at a later time. This is a short term solution while all logs are changed over and import the SplunkUtils module.

    # Calling "$($(Get-PSCallStack)[1].InvocationInfo.PSCommandPath)" in the module will give the scripts location that called the module.
    
    If ($Severity -eq "") {
        $Severity = 'INFO'
    }
    if ($ErrMsg -eq "") {
        $ErrMsg     = "N/A"
        $Type       = "N/A"
        $Status     = "N/A"
    }
    $SplunkLog = @{
        Status      = $Status
        LegacyError = "$($msg)"
        Message     = "$($ErrMsg)"
        Severity    = $Severity
        Process     = "$((Get-PSCallStack)[0].InvocationInfo.MyCommand.Source)"
        Type        = $Type
        User        = $env:USERNAME
        Action      = $Action
        Resource    = $Resource
    }

    if($logPath -eq "" -or $logPath -eq $null) {
        $logPath = "C:\temp\Logs.txt"
    }

    Add-Content -LiteralPath $logPath -Value ([DateTime]::Now.ToString() + ' - ' + $msg) -Encoding utf8
}


function Get-eDOCSLogon {
    <#
    .SYNOPSIS
    Logs a user onto eDOCS and generates a security token that is required for
    all other actions 
    
    .DESCRIPTION
    All actions with the eDOCS API require passing a security token to eDOCS.
    This function performs a logon to eDOCS and returns a hashtable indicating
    if the login succeeded and if so the security token associated with the
    login.
    
    .PARAMETER dmUser
    Username of the account to login to eDOCS DM with.
    
    .PARAMETER dmPass
    eDOCS DM password for the specified account.
    
    .PARAMETER dmLibrary
    eDOCS DM library to be logged into typically EDOCS for prod or EDOCS_TEST
    for test.
    
    .PARAMETER logPath
    Path to the log file you are using for your script, any login failures will
    be written to this file.

    .PARAMETER dmServer
    Enabled the ability to have a specific server handle the logon call. 
    Unsure how well this works, but it was a native option within the API so why not.    
    .EXAMPLE
    Get-eDOCSLogon -dmUser "currentuser" -dmPass "secrets" -dmLibrary "EDOCS_TEST" -logPath "H:\My\Script\Log.txt"

    Would return 
    @{
        "success" = $true;
        "dst" = "12345abcdefgh";
    }

    If the logon succeeded or 
    @{
        "success" = $false;
        "dst" = -1;
    }

    If the logon failed (it would also write an entry to H:\My\Script\Log.txt)
    with the eDOCS server failure reason.
    
    .NOTES
    You might want to refactor this function to return the PCDClient.PCDLogin
    object rather than just a reference to the DST token.
    #>

    Param(
        [parameter(Mandatory = $true, Position = 0)]
        [String]
        $dmUser,
        [parameter(Mandatory = $true, Position = 1)]
        [String]
        $dmPass,
        [parameter(Mandatory = $true, Position = 2)]
        [String]
        $dmLibrary,
        [parameter(Mandatory = $true, Position = 3)]
        [String]
        $logPath,
        [parameter()]
        [ValidateSet('EDOCSDMPROD1', 'EDOCSDMPROD2', 'EDOCSDMPROD3', 'EDOCSDMPROD4', 'EDOCSDMPROD5', 'EDOCSDMPROD6', 'EDOCSWEBPROD', 'EDOCSDMTEST1')]
        $dmServer
    )

    # Login to the given eDOCS library. More details about this process available on 
    # page 55 of the eDOCS DM 5.3.1 DM API Reference Guide.
    $login = New-Object -ComObject PCDClient.PCDLogin
    $login.AddLogin(0, $dmLibrary, $dmUser, $dmPass)
    If($dmLibrary -eq "EDOCS") {
        If(!$dmServer) {
            $dmServer = "EDOCSDMPROD1"
        }
    } Else {
        If(!$dmServer) {
            $dmServer = "EDOCSDMTEST1"
        }
    }
    $login.SetServerName("$dmServer")
    $login.Execute()
    $dmDST = $login.GetDST()
    $prodserver = $login.GetServerName()

    if ($dmDST.Length -eq 0) {
        $msg = 'Failed to login to eDOCS ' + $dmUser + '; Error: ' + $login.ErrDescription
        Add-eDOCSScriptLogEntry -msg $msg -logPath $logPath -ErrMsg "Log into $dmUser on $dmLibrary server $dmServer" -Severity INFO -Status Failed -Type Script -Action Login -Resource $dmLibrary
        @{
            success = $false;
            dst     = 0;
            Server  = $prodserver;
        }
    } else {
        Add-eDOCSScriptLogEntry -ErrMsg "Log into $dmUser on $dmLibrary server $dmServer" -Severity INFO -Status Succeeded -Type Script -Action Login -Resource $dmLibrary
        @{
            success = $true;
            dst     = $dmDST;
            server  = $prodserver;
        }
    }
}

function Query-EdocsSQL {
    <#
    .SYNOPSIS
    Logs a user onto eDOCS and generates a security token that is required for
    all other actions 
    
    .DESCRIPTION
    All actions with the eDOCS API require passing a security token to eDOCS.
    This function performs a logon to eDOCS and returns a hashtable indicating
    if the login succeeded and if so the security token associated with the
    login.

    .PARAMETER Server
    Select the MSSQL server your tables reside.

    .PARAMETER Database
    Select Database on the MSSQL Server.

    .PARAMETER Query
    Construct your SQL Query and pass to this param.
        
    .EXAMPLE
    $Query = "SELECT TOP(100) * FROM docsadm.Profile"
    Query-EdocsSQL -Server SQLVMPD001\DB01 -Database EDOCS -Query $Query
    
    .NOTES

    #>

    param (
        [Parameter(Mandatory=$true)]
        [ValidateSet('SQLVMPD001\DB01','SQLVMTD001\DB01')]
        [string]$Server,
        [Parameter(Mandatory=$true)]
        [ValidateSet('EDOCS','EDOCS_TEST')]
        [string]$Database,
        [Parameter(Mandatory=$true)]
        [string]$Query
    )

    $connection = new-object System.Data.SqlClient.SQLConnection("Data Source=$server;Integrated Security=SSPI;Initial Catalog=$database");
    $cmd = new-object System.Data.SqlClient.SqlCommand($Query, $connection);

    $connection.Open();
    $reader = $cmd.ExecuteReader()

    $results = @()
    while ($reader.Read()) {
        $row = @{}
        for ($i = 0; $i -lt $reader.FieldCount; $i++) {
            $row[$reader.GetName($i)] = $reader.GetValue($i)
        }
        $results += new-object psobject -property $row            
    }
    $connection.Close();
    $results
}

function Import-eDOCSDocument {
    <#
    .SYNOPSIS
    Import a given file (and associated metadata) into eDOCS.

    .DESCRIPTION
    Given a file on a Windows file system and eDOCS document metadata import
    the file into eDOCS as a document with the metadata set as specified.

    .PARAMETER taskID
    Unique ID used to keep track of multiple documents being updated as a
    single batch.

    .PARAMETER dmDST
    eDOCS security token.

    .PARAMETER library
    eDOCS library to import the document into.

    .PARAMETER logPath
    Path to the log file being used for this operation.

    .PARAMETER form
    eDOCS profile form being used to perform the import.

    .PARAMETER filepath
    Full path to the file being imported as an eDOCS document.

    .PARAMETER properties
    eDOCS metadata values to be associated with the document. The following are
    mandatory:
    DOCNAME - title of the document max length is 240 characters.
    TYPE_ID - document type, typically this will be DEFAULT other options
        include REPORT, BRIEF etc.
    APP_ID - eDOCS application that the document is to be associated with.
    AUTHOR_ID - user name to set as the document author must be in the eDOCS
        people table.
    TYPIST_ID - user name to set as the document typist (i.e. registered by)
        must be in the eDOCS people table.
    PD_FILEPT_NO - the eDOCS file part that the document is to be saved into
        e.g. TEST/0000001-001(E)
    
    The following is optional but HIGHLY recommended, just pretend I listed it
    under mandatory:
    FILE_EXTENSION - eDOCS won't automatically use the file extension for the
        document you are importing instead you should explicitly provide it. If
        you don't it will use the default file extension for the eDOCS
        application you set via the APP_ID property. This may be the wrong
        extension for your file (e.g. docx when you wanted .doc) and will cause
        your file to be unopenable until the extension is fixed.

    The following are properties you may use more often than not:
    ABSTRACT - this is the profile form description field.
    KMC_OLD_DOCNO - useful if you are importing a document from another system.
    PIF_LETTER_DATE - the date written field.
    
    The following are properties for the email forms:
    PD_ORIGINATOR - sender
    PD_ADDRESSEE - to
    PD_EMAIL_CC - cc
    EMAIL_RECEIVED - email received date
    EMAIL_SENT - email sent date

    .EXAMPLE
    $properties = @{
        "DOCNAME"        = "My super great document";
        "TYPE_ID"        = "DEFAULT";
        "FILE_EXTENSION" = ".pdf";
        "APP_ID"         = "PDF";
        "AUTHOR_ID"      = "dm_imports";
        "TYPIST_ID"      = "dm_imports";
        "PD_FILEPT_NO"   = "TEST/0000345-001(E)";
        "ABSTRACT"       = "This document rocks!";
    }

    $params = @{
        "taskID"     = $taskId++;
        "dmDST"      = $login.dst;
        "library"    = "EDOCS_TEST";
        "logPath"    = $log;
        "form"       = "KMC_ELEC_DOC";
        "filepath"   = "C:\My\Docs\123.pdf";
        "properties" = $properties;
    }

    $docnumber = Import-eDOCSDocument @params

    Would import document C:\My\Docs\123.pdf into file part TEST/0000345-001(E)
    in library EDOCS_TEST with name My super great document. It would then
    assign the document number of the imported document to variable $docnumber
    for future use within the script.

    .NOTES
    At the time of writing the APP_ID for paper documents is set to null but
    the API seems to take exception to setting properties to null. This means
    I haven't yet worked out how to create paper document profiles
    programmatically.
    #>

    Param(
        [parameter(Mandatory = $true, Position = 0)]
        [String]
        $taskID,
        [parameter(Mandatory = $true, Position = 1)]
        [String]
        $dmDST,
        [parameter(Mandatory = $true, Position = 2)]
        [String]
        $library,
        [parameter(Mandatory = $true, Position = 3)]
        [String]
        $logPath,
        [parameter(Mandatory = $true, Position = 4)]
        [String]
        $form,
        [parameter(Mandatory = $true, Position = 5)]
        [String]
        $filepath,
        [parameter(Mandatory = $true, Position = 6)]
        $properties
    )

    # Create profile more information about this process can be found on
    # pages 41 and 84 of the eDOCS DM 5.3.1 DM API Reference Guide.
    $doc = New-Object -ComObject PCDClient.PCDDocObject.1
    $doc.SetDST($dmDST) > $null
    $doc.SetObjectType($form) > $null
    $doc.SetProperty("%TARGET_LIBRARY", $library) > $null

    # Check that each of the properties is a string before trying to set it.
    # Once I accidently set a property to an an array value and it everything
    # crashed terribly. Strictly speaking SetProperty can take non-string
    # values but I can't think of any that we would use. This code may cause
    # problems down the track.
    ForEach ($prop in $properties.GetEnumerator()) {
        if ($prop.Value) {
            if ($prop.Value.GetType().Name -eq "String") {
                $doc.SetProperty($prop.Name, $prop.Value) > $null
            } else {
                Add-eDOCSScriptLogEntry "$taskID; Failure; -1; -1; Failed to set property $($prop.Name); Value not a string." $logPath -ErrMsg "Failed to set property Task# $taskID : $($prop.Name)" -Status Failed -Severity MEDIUM -Type Script -Resource $library -Action Import
            }
        }
    }

    # Setting %VERIFY_ONLY to %YES will just check if the document properties
    # are valid without creating the document.
    $doc.SetProperty("%VERIFY_ONLY", "%YES") > $null
    $doc.Create() > $null

    # If the properties are valid woohoo, lets turn %VERIFY_ONLY off so we can
    # actually create the profile.
    if ($doc.ErrNumber -eq 0) {
        $doc.SetProperty("%VERIFY_ONLY", "%NO") > $null
        $doc.Create() > $null
    }
    # Verification failed, abort!
    else {
        Add-eDOCSScriptLogEntry "$taskID; Failure; -1; -1; Failed to verify profile $filepath; Error: $($doc.ErrDescription)" $logPath -ErrMsg "Profile Not Verified for Task# $taskID : $filepath" -Status Failed -Severity MEDIUM -Type Script -Action Import -Resource $library
        return
    }

    # Yes we already verified and it said it was okay but we still need to
    # check that the actual profile creation worked.
    if ($doc.ErrNumber -ne 0) {
        # Release com objects used during import process so far.
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) > $null
        Add-eDOCSScriptLogEntry "$taskID; Failure; -1; -1; Failed to create profile $filepath; Error: $($doc.ErrDescription)" $logPath -ErrMsg "Profile Not Created for Task# $taskID : $filepath" -Status Failed -Severity MEDIUM -Type Script -Action Import -Resource $library
        return
    } else {
        # Okay profile created, lets upload the actual document now.
        # Get key for components table, this is generated automatically by
        # eDOCS when a profile is created and is required to write the document
        # to the eDOCS doc store.
        $docNumber = $doc.GetReturnProperty("%OBJECT_IDENTIFIER")
        $docVersion = $doc.GetReturnProperty("%VERSION_ID")

        # The PCDPutDoc object will be used to generate a pointer required to
        # copy the actual document into the doc store.
        $putDoc = New-Object -ComObject PCDClient.PCDPutDoc.1
        $putDoc.SetDST($dmDST) > $null
        $putDoc.AddSearchCriteria("%TARGET_LIBRARY", $library) > $null
        $putDoc.AddSearchCriteria("%DOCUMENT_NUMBER", $docNumber) > $null
        $putDoc.AddSearchCriteria("%VERSION_ID", $docVersion) > $null
        $putDoc.Execute() > $null

        if ($putDoc.ErrNumber -ne 0) {
            # Release com objects used during import process so far.
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) > $null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($putDoc) > $null
            Add-eDOCSScriptLogEntry "$taskID; Failure; $docNumber; $docVersion; Failed to find profile $docNumber; Error: $($putDoc.ErrDescription)" $logPath -ErrMsg "Import Task # $taskID into $library - Profile Not Found for Task# $taskID : $docNumber" -Status Failed -Severity MEDIUM -Type Script -Action Import -Resource $library
            return
        } else {
            # Write file to eDOCS server by reading it in as a byte array and writing chunks using
            # the putStream.Write method.
            $putDoc.NextRow() > $null
            $putStream = New-Object -ComObject PCDClient.PCDPutStream.1
            $putStream = $putDoc.GetpropertyValue("%CONTENT")

            # Number of bytes to write with each write action.
            $buffer = 1024 * 64
            
            # Open the document we are writing to the doc store.
            $fileO = New-Object IO.FileStream($filepath, [IO.FileMode]::Open)
            
            # Set variables used to track the write process.
            $fileSize = $fileO.Length
            $totalBytesWritten = 0
            $readResult = -1

            # $readResult will be set to 0 once all bytes have been read.
            while ($readResult -ne 0) {
                # Work out how many bytes we want to read with the next read
                # action. It will either be the buffer, the total number of
                # unread bytes (total remaining is less than the buffer) or
                # 0 (read as been completed).
                $bytesRemaining = $fileO.Length - $fileO.Position

                if ($bytesRemaining -gt $buffer) {
                    $toRead = $buffer
                } else {
                    $toRead = $bytesRemaining
                }

                if ($toRead -gt 0) {
                    # Perform the next read action.
                    $chunk = New-Object byte[] $toRead
                    $readResult = $fileO.Read($chunk, 0, $toRead)

                    # Write the results of the read action to the doc store.
                    $putStream.Write($chunk, $toRead) > $null

                    # Keep track of how many bytes have been successfully
                    # written so we can validate at the end.
                    $totalBytesWritten += $putStream.BytesWritten
                } else {
                    $readResult = 0
                }
            }

            # End the write process and close the document.
            $putStream.SetComplete() > $null
            $fileO.Close()

            # Upload process failed some how.
            if ($putStream.ErrNumber -ne 0) {
                # Release com objects used during import process so far.
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) > $null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($putDoc) > $null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($putStream) > $null
                Add-eDOCSScriptLogEntry "$taskID; Failure; $docNumber; $docVersion; Failed to write file to eDOCS $docNumber; Error: $($putStream.ErrDescription)" $logPath -ErrMsg "Upload Failed on Task# $TaskID : $docNumber" -Status Failed -Severity MEDIUM -Type Script -Action Import -Resource $library
                return
            }

            # eDOCS didn't report an error but the file written to eDOCS is a
            # different size to the original.
            if ($totalBytesWritten -ne $fileSize) {
                Add-eDOCSScriptLogEntry "$taskID; Failure; $docNumber; $docVersion; Upload failed part way through $docNumber; Error: $totalBytesWritten of $fileSize bytes written" $logPath -ErrMsg "Size Mismatch on Task# $TaskID : $docNumber" -Status Failed -Severity MEDIUM -Type Script -Action Import -Resource $library
                return
            }

            # Unlock document, documents uploaded by the eDOCS DM API will remain locked until manually unlocked.
            $doc = New-Object -ComObject PCDClient.PCDDocObject.1
            $doc.SetDST($dmDST) > $null
            $doc.SetObjectType($form) > $null
            $doc.SetProperty("%TARGET_LIBRARY", $library) > $null
            $doc.SetProperty("%OBJECT_IDENTIFIER", $docNumber) > $null
            $doc.SetProperty("%VERSION_ID", $docVersion) > $null
            $doc.SetProperty("%STATUS", "%UNLOCK") > $null
            $doc.Update() > $null

            # Of all the things to fail unlocking the document wasn't expected.
            if ($doc.ErrNumber -ne 0) {
                # Release com objects used during import process so far.
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) > $null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($putDoc) > $null
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($putStream) > $null
                Add-eDOCSScriptLogEntry "$taskID; Failure; $docNumber; $docVersion; Failed to unlock document $docNumber; Error: $($doc.ErrDescription)" $logPath -ErrMsg "Unlock Failed on Task# $TaskID : $docNumber" -Status Failed -Severity MEDIUM -Type Script -Action Import -Resource $library
                return
            }

            # upload completed, write to log and return the number of the uploaded document
            Add-eDOCSScriptLogEntry "$taskID; Success; $docNumber; $docVersion; Document Uploaded $docNumber, $filepath" $logPath -ErrMsg "Imported DocNum: $docNumber" -Status Succeeded -Severity LOW -Type Script -Action Import -Resource $library
            
            # Release com objects used during import process so far.
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) > $null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($putDoc) > $null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($putStream) > $null

            # Return the document number as the function output.
            $docNumber
        }
    }
}

function Set-eDOCSProfileEscapeCharacters {
    <#
    .SYNOPSIS
    The eDOCS document profile description/abstract field translates at least
    \\, \r, \n and \t into line breaks and tabs this function adds an extra \
    to escape these characters.
    
    .DESCRIPTION
    The eDOCS document profile description/abstract field translates at least
    \\, \r, \n and \t into line breaks and tabs this function adds an extra \
    to escape these characters.

    When writing to the eDOCS document profile description/abstract field you
    should call this function over the data you are writing.

    If you want to write a line break to the description/abstract field then
    add a line break via "`r`n" or System.String.TextBuilder.AppendLine() into
    the string you write to this field.

    There may be other special characters that I haven't run into yet, to my
    knowledge these aren't documented anywhere.
    
    .PARAMETER string
    The string you need to escape before writing to the eDOCS profile
    description field.
    
    .EXAMPLE
    Set-eDOCSProfileEscapeCharacters "\\cbdfile1\nature\really\is\cool"
    Would return \\\\cbdfile1\\nature\\really\is\cool
    
    .NOTES
    If you want to write a line break to the description/abstract field then
    add a line break via "`r`n" or System.String.TextBuilder.AppendLine() into
    the string you write to this field.

    There may be other special characters that I haven't run into yet, to my
    knowledge these aren't documented anywhere. 
    #>

    Param(
        [parameter(Mandatory = $true, Position = 0)]
        [String]
        $string
    )

    $string = $string -replace '\\\\', '\\\\'
    $string = $string -replace '\\n', '\\n'
    $string = $string -replace '\\r', '\\r'
    $string = $string -replace '\\t', '\\t'
    $string
}

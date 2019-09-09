Param(
    [Parameter(Mandatory=$true)]
    [string]$Domain   
)

# Function to select the input file.
Function Get-FileName($initialDirectory, $ltitle){   
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
    Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.Title = $ltitle
    $OpenFileDialog.filter = "All files (*.*)| *.*"
    $OpenFileDialog.ShowHelp = $true
    $OpenFileDialog.ShowDialog() | Out-Null
    Return $OpenFileDialog.filename
}
# End Function Get-FileName


#Function to create log file
Function Create-LogFile(){
    Param ($TimeStamp)
    $LogFile= "Change-Attribute-Log"+"-"+$TimeStamp+".txt"
    New-Item $LogFile -ItemType "file"
    Add-Content $LogFile -Value "Log File Created at Time $Timestamp" 
    Add-Content $LogFile -Value "-------------------------------------------------------------------------------------------------------------"
    return $LogFile;
}
# End of function Create-LogFile

#Function to update log file
Function Log-Write(){
    Param ([string]$logstring,$Logfile)
    Add-Content $Logfile -Value (Get-date -UFormat "%d:%m:%Y:%R" )
    Add-content $Logfile -value $logstring
    Add-Content $LogFile -Value "-------------------------------------------------------------------------------------------------------------"
}
#End of function Log-Write

Function Call-Main(){
    #Creating the log file.
    Write-Host "Creating the log file for script execution" -ForegroundColor Yellow -BackgroundColor Black
    $TimeStamp = Get-Date -Format o | ForEach-Object { $_ -replace ":", "." }
    $Logfile=(Create-LogFile -TimeStamp $TimeStamp).Name

    Start-Sleep -Seconds 3

    #Checking the ActiveRolesManagementShell (Quest) Module
    Write-Host "`n Importing the ActiveRolesManagementShell Module in the powershell session" -ForegroundColor Yellow -BackgroundColor Black
    $Error.Clear()
    Import-Module ActiveRolesManagementShell -ErrorAction SilentlyContinue
    if($Error){
        Write-Host "`n ActiveRolesManagementShell module not installed." -ForegroundColor DarkRed -BackgroundColor White
        Write-Host "`n Exiting the script" -ForegroundColor DarkRed -BackgroundColor White
        Log-Write -Logfile $Logfile -logstring "ActiveRolesManagementShell module not installed.Exiting the Script"
        return;
    }
    Else{
        Write-Host "`n ActiveRolesManagementShell Powershell module imported" -ForegroundColor Yellow -BackgroundColor Black
        Log-Write -Logfile $Logfile -logstring "ActiveRolesManagementShell Powershell module imported"
    }

    #Checking domain availability
    $Error.Clear()
    $DomainDN=(Get-QADRootDSE).DefaultNamingContextDN
    if($Error){
        Write-Host "`n Active Directory domain not found.Please login to a domain joined machine with a domain account.Exiting the script" -ForegroundColor DarkRed -BackgroundColor White
        Log-Write -Logfile $Logfile -logstring "Active Directory domain not found.Please login to a domain joined machine with a domain account.Exiting the script"
        return;
    }

    Start-Sleep -Seconds 3

    #Connecting to available domain controller
    $Error.Clear()
    $ServiceConnectionOutput=(Connect-QADService -servicename $Domain).Domain.Name
    if($Error){
        Write-Host "`n Not able to connect to a domain controller.Exiting the script" -ForegroundColor DarkRed -BackgroundColor White
        Log-Write -logstring "$Error.Exception" -Logfile $Logfile
        Log-Write -logstring "Exiting the script" -Logfile $Logfile
        return;
    }
    else{
        Write-Host "`n Connected to the domain $ServiceConnectionOutput"
        Log-Write -logstring "Connected to the domain $ServiceConnectionOutput" -Logfile $Logfile
    }

    #Selecting the Input File.
    Write-Host "`n Please select the input file.Please ensure the header (Display Name,Alias,Organizational Unit) is kept intact." -ForegroundColor Yellow -BackgroundColor Black
    $FilePath = Get-FileName -$ltitle "Select Input File"
    Write-Host "$FilePath file has been selected" -ForegroundColor Yellow -BackgroundColor Black
    Log-Write -logstring "$FilePath file has been selected" -Logfile $Logfile
    
    Start-Sleep -Seconds 3

    #Processing the users
    Write-Host "`n Process to Alter the Display Name for the users in the selected file will commence" -ForegroundColor Yellow -BackgroundColor Black
    Log-Write -logstring "Process to Alter the Display Name for the users in the selected file will commence." -Logfile $Logfile
    $ChangeList=Import-Csv $FilePath
    $ChangeSummaryReport=@()
    foreach($user in $ChangeList){
        $Error.clear()
        Start-Sleep -Seconds 3
        $DisplayName=$User.'Display Name'
        $Alias=$User.Alias
        Write-Host "`n Processing User $DisplayName for DisplayName change.Removing the ',' from displayname" -ForegroundColor DarkGreen -BackgroundColor White
        Log-Write "Processing User $DisplayName for DisplayName change.Removing the ',' from displayname" -Logfile $Logfile
        $DistinguishedName=(Get-QADUser -DisplayName $DisplayName).DN
        $NewDisplayName=$DisplayName.replace(",","")
        Set-QADUser -Identity $DistinguishedName -displayname $NewDisplayName
        $NewDisplayName=(Get-QADuser -identity $DistinguishedName).DisplayName
        if($Error){
            Write-Host "`n Change of DisplayName attribute failed for user $DisplayName.Please check summary report for details"
            Log-Write -logstring "Change of DisplayName attribute failed for user $DisplayName.Please check summary report for details" -Logfile $Logfile
            $UserReport = New-Object PsObject
            $UserReport | Add-Member OldDisplayName $DisplayName
            $UserReport | Add-Member NewDisplayName "NA"
            $UserReport | Add-Member Alias $Alias
            $UserReport | Add-Member ChangeStatus "Failed"
            $UserReport | Add-Member Error "$Error.Exception"
            $ChangeSummaryReport=$ChangeSummaryReport+$UserReport
        }
        Else{
            Write-Host "`n Change of DisplayName attribute completed successfully for user $DisplayName.Please check summary report for details"
            Log-Write -logstring "Change of DisplayName attribute completed successfully for user $DisplayName.Please check summary report for details" -Logfile $Logfile
            $UserReport = New-Object PsObject
            $UserReport | Add-Member OldDisplayName $DisplayName
            $UserReport | Add-Member NewDisplayName $NewDisplayName
            $UserReport | Add-Member Alias $Alias
            $UserReport | Add-Member ChangeStatus "Success"
            $UserReport | Add-Member Error "NA"
            $ChangeSummaryReport=$ChangeSummaryReport+$UserReport
        }
    }

    Start-Sleep -Seconds 3

    #Creating the ChangeSummaryReport
    $ChangeSummaryFileName="Change-Summary-Report-"+$TimeStamp+".csv"
    $ChangeSummaryReport | Export-csv $ChangeSummaryFileName -NoClobber -NoTypeInformation
    Write-Host "`n Created the Change summary report file $ChangeSummaryFileName" -ForegroundColor Yellow -BackgroundColor Black
    Log-Write -logstring "Created the Change summary report file $ChangeSummaryFileName" -Logfile $Logfile
}

Call-Main;

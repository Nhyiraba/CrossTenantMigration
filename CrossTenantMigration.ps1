
#Instruction for each sectiion

$ActionSelection = @"
Select the action you want to perform by entering the corresponding number
______________________________________________________________________________________

CHOICE SELECTION 
=========================================

    1 -- Create one  mailuser on Target/destination tanant"
    2 -- Crete in bulk multiple mailusers by CSV file"
    3 -- Stand ExchangeGUID and LegacyExchangDN"

"@

$BulkMUInst = @"
================ Bulk mail user creation =======================

Provide csv file for bulk. The CSV file must contant 3 columns

Example: 
    Proivde CSV file with 3 columns DisplayName, EmailAddress, Password

    DiaplayName,EmailAddress, Password
    Daniel Alex,dalex@check.com,PassWord!@#
    Daniel Mykel,dykel@check.com,PassWord!@# 

"@

$ExtractBulkorOneInst = @"

================ Bulk User Mailbox Infomation Retrival =======================

Select the corresponding option for data source

    1 -- Enter email address or display name of the mailboxes seperated by comma on single line.
         Example : 
            dlex@hoperoom.com, ernesto@hoperoom.com or Daniel Alex, Ernest Alex

    2 -- Select a CSV file that contain list of users, the should have no header
         Example :
            Daniel Mykel,
            dykel@check.com
            gylex@checj.com
            Atta Amam
            .................. nth

==============================================================================

"@


function Get-CSVFile {
    #get csv file
    [void] [System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    if ($initialDirectory) { $OpenFileDialog.initialDirectory = "." }
    $OpenFileDialog.filter = "CSv files (*.csv)|*.csv"
    $OpenFileDialog.Title = "Select CSV file"
    [void] $OpenFileDialog.ShowDialog()
    return $OpenFileDialog.FileName
}

function CrossT2TMigration {
  
    <#
    .Synopsis
        This script will enable user to automate the of the processess in volved in tenant to tenant mailbox migration.

    .Description
        add later

    .Example
        later   
#>

    [CmdletBinding(DefaultParameterSetName = "Default")]
    param (
        # MailUser principal name
        [Parameter(
            Mandatory = $false, ValueFromPipeline, ValueFromPipelineByPropertyName,
            ParameterSetName = "MailUserPrincipalName"
        )]
        [array]$CreateTargetMailUser,

        [Parameter(Mandatory = $false, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        #[ValidateSet("1", "2","3")] # bulk mailuser creation or single
        [string]$ChoseSelection

    )


    # Choice seelect tion
    Write-Host $ActionSelection

    $ChoseSelection = Read-Host "CHOICE SELECTION  "
    
    switch ($ChoseSelection) {

        "1" {
            CreateOneTragetMailUser 
        }
        "2" { 
            Write-Host 
            

        }
    
        Default { "sdflshlfks" }
    }
}


#Creating mail user on source tenant
function CreateOneTragetMailUser {
   
    [CmdletBinding(DefaultParameterSetName = "Default")]
    param (
        # MailUser password
        [Parameter()]
        [string[]] $MailUserPwd
    )

    Write-Host "`nEnter Detail Seperated by comma (,) - DisplayName, MailUserUPN, Passowrd`n"
    Write-host "Example : Daniel Alex, dlex@hoperoom.com, Passw0rd!@# `n" -ForegroundColor Yellow
    $MailUserDetail = Read-Host "Entert the mail user detail "
    $MailUserDetail = $MailUserDetail.Split(",").Trim().Trim("'").Trim('"') #split and remove all white spaces from the imput

    Write-Host $MailUserDetail
    if ($MailUserDetail[1].ToLower().Contains("@") -eq $true) {
        #create  mail user with entered data
        New-MailUser -Name $MailUserDetail[0]  -MicrosoftOnlineServicesID  $MailUserDetail[1] -DisplayName $MailUserDetail[0] `
            -Password (ConvertTo-SecureString -String $MailUserDetail[2] -AsPlainText -Force)
    }
    else {
        $pvd = $MailUserDetail[1] #email check
        Write-Host "Invalid import, email/UPN must contain the @ character. You entered $pvd " -ForegroundColor Red
    }
        
}


#Creating mail user on source tenant

function CreateBulkTargetMailUser {

    
    [CmdletBinding(DefaultParameterSetName = "Default")]
    param (
        [Parameter(Mandatory = $false)]
        [ValidateScript({
                if (Test-Path $_) { $true }
                else { throw "Path $_ is not valid" }
            })]
        [string]$BulkUserData
    )

    if (-not($BulkUserData)) { $BulkUserData = Get-CSVFile }

    $LoadBulkUserData = Import-Csv -Path $BulkUserData
    
    #creating users
    Write-Host "`nMailuser object creation started`n--------------------------------------" -ForegroundColor Green
    $LoadBulkUserData | ForEach-Object { `
            New-MailUser -Name $_.DiaplayName -MicrosoftOnlineServicesID $_.EmailAddress  -DisplayName $_.DiaplayName `
            -Password (ConvertTo-SecureString -String $_.Password -AsPlainText -Force)
        Write-Host "`tDone: " $_.DiaplayName
    }
    Write-Host "--------------------------------------`nMailuser object creation completed" -ForegroundColor Green
        
}

<#
    This function is designed to retrive user propertions for a given mailbox
    The mailbox or mailboxes can be supplied inline data seperated by comma or 
    by csv file with single but not header. The column can be mix user "Display Name" or email address

    Restults:
        This will return a table formated file and can be exported as csv file.
#>
function ExtractExchGUIandX500 {

    [CmdletBinding(DefaultParameterSetName = "Default")]
    param (
        # User mailbox email address or full display name
        [Parameter(Mandatory = $false)]
        [array]
        $BulkOrOneMailbox
    )

    #for retrieving the mailbox information
    function RetrievMailboxInfo {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [array]
            $UserMailbox
        )
        
        # get mailbox information  and return array object
        $UserMailbox | ForEach-Object {
            Get-MailBox -Identity $_ | Select-Object DisplayName, PrimarySmtpAddress, ExchangeGuid, LegacyExchangeDN

            $MailBoxInfo = [PSCustomObject]@{
                DisplayName      = $_.DisplayName
                ExchangeGuid     = $_.ExchangeGuid
                EmailAddress     = $_.PrimarySmtpAddress
                LegacyExchangeDN = $_.LegacyExchangeDN
            }
        }
    }


    #requesting user to input mailbox email address or displayname if not specified
    # this provides flexibility to specify the paramter when calling the function independently
    if (-not($BulkOrOneMailbox)) {

        Write-Host $ExtractBulkorOneInst

        $DataSource = Read-Host " CHOICE SELECTION (1 or 2) "

        Write-Host "`n"
        
        switch ($DataSource) {
            "1" { 
                $BulkOrOne = Read-Host "Enter the Mailbox email address"
                $BulkOrOneMailbox = $BulkOrOne.split(",").Trim().Trim("'").Trim('"') #split and remove all white spaces from the imput

                if ($BulkOrOneMailbox.Length -ge 1) {
                    #get single mailbox information
                    $MailboxInfoResults = RetrievMailboxInfo -UserMailbox $BulkOrOneMailbox
                    return $MailboxInfoResults 
                }
                else {
                    Write-Host "You have provided invalid email address or display name"
                }
            }
            "2" {
                # this is single column CSV data without any column name, and it can be a mix of email addressess and display name
                Write-Host "Retrieving the DisplayName, PrimarySmtpAddress, ExchangeGUID and LegacyExchangeDN of the mailbox "
                $getCsvData = Get-CSVFile #get file
                $readCsvData = Get-Content -Path $getCsvData
                    
                $MailboxInfoResults = RetrievMailboxInfo -UserMailbox $readCsvData
                return $MailboxInfoResults
            
            }
            Default { Write-Host "Invalid selected choice" }
        }
        
    }
    else {
    
        $BulkOrOne = $BulkOrOneMailbox.split(",").Trim() #split and remove all white spaces from the imput

        if ($BulkOrOne.Length -ge 1) {
            #get single mailbox information
            $MailboxInfoResults = RetrievMailboxInfo -UserMailbox $BulkOrOne
            return $MailboxInfoResults 
        }
        else {
            Write-Host "You have provided invalid email address or display name"
        }
    }

}

function ExchGuidAndLegacyExchDNTranferSourceToTarget {
    param ()

    <#
        Its recommneded to use the same Display Name from the source tenant aids the transfer of object properties

        Automated process
        ===========================
            Get source properties using 
                $userpros = ExtractExchGUIandX500 function
            Connect to target tenant
                Connect-exchangeOnline #target tenant admin
            Use the properties from $userpros to stamp it on the target mail users
                $userpros | ForEach-Object { Set-MailUser -identity $_.DisplayName -ExchangeGUID [GUID]$_.ExchangeGuid -EmailAddresses@{add="x500"$_.LegacyExchangDN}}
            Use the display name from the retrieved data

        Manual process
        ==========================
        Another way is to manually map the object from the destination to source properties
        
        To get source properties use
            ExtractExchGUIandX500, 
            export the results as csv
            Open the csv file
            Create new column and the mail user emails address from the target tenant.
    #>

    


    
}
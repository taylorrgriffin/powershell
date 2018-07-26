$mailboxToExport = "Taylor.Testing@oregonstate.edu"
$destinationMailbox = "reinscha-test@oregonstate.edu" 

function SelectMailboxFolder ($emailaddress) {
    ((Get-MailboxFolderStatistics $emailaddress).identity |Out-GridView -OutputMode Single -Title "Select Folder to export").tostring()
}

function mailboxExists([string]$mailbox) {
    $checkMailbox = (Get-Mailbox $mailbox -erroraction SilentlyContinue);
    if ($checkMailbox -eq $null) {
        return $false
    }
    else {
        return $true
    }
}

function startExport([string]$mailboxToExport,$tempPST,$sourceFolder) {
    New-MailboxExportRequest -Name $mailboxToExport -Mailbox $mailboxToExport -FilePath $tempPST -SourceRootFolder $sourceFolder -ExcludeDumpster
}

function startImport([string]$destinationMailbox,$tempPST) {
    New-MailboxImportRequest -Name $destinationMailbox -Mailbox $destinationMailbox -TargetRootFolder "Calendar" -FilePath $tempPST -ExcludeDumpster
}

function waitExport([string]$mailboxToExport) {
    do {
        $checkExport = (Get-MailboxExportRequest | where-object {$_.Name -like $mailboxToExport.toString()}).status
        sleep -Seconds 3
        $exportFinished = $true
        for ($j=0; $j -lt $checkExport.length;$j++) {
            if (-Not ($checkExport[$j].toString() -eq "Completed")) {
                $exportFinished = $false
                "Export is not yet complete, sleeping for three seconds."
            }
        }
    }
    until($exportFinished -eq $true)
    "Export is finished."
    # check that all content was exported
}

function waitImport([string]$destinationMailbox) {
    do {
        $checkImport = (Get-MailboxImportRequest | where-object {$_.Name -like $destinationMailbox.toString()}).status
        sleep -Seconds 3
        $importFinished = $true
        for ($j=0; $j -lt $checkImport.length;$j++) {
            if (-Not ($checkImport[$j].toString() -eq "Completed")) {
                $importFinished = $false
                "Import is not yet complete, sleeping for three seconds."
            }
        }
    }
    until($importFinished -eq $true)
    "Import is finished."
    # check that all content was imported
}

function export_import() {
    if (mailboxExists($mailboxToExport)) {
        "Mailbox exists, proceeding with the export."
        $folderToExport = SelectMailboxfolder -emailaddress $mailboxToExport
        if ($folderToExport -eq $null) {
            "Oops! Looks like there's no folders in this mailbox."
        }
        else {
            $sourceFolder = $folderToExport.Replace($mailboxToExport+"\","")
            if (-Not($sourceFolder -Match "Calendar")) {
                "Oops! You must select a subfolder of Calendar or the root Calendar folder itself to import from."
            }
            else {
                # REMEMBER TO CHANGE THIS LATER OKKKKK
                $tempPST = "\\iscs-export\Export\griftayl"+$mailboxToExport.toString()+".PST"
                "Starting export..."
                startExport $mailboxToExport $tempPST $sourceFolder
                waitExport $mailboxToExport
                "Starting import..."
                startImport $destinationMailbox $tempPST
                waitImport $destinationMailbox
            }
        }
    }
    else {
        "Mailbox doesn't exist, cancelling import."
    }
}

export_import
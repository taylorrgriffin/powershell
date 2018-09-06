function CalendarExport
{
    PARAM
    (
        [Parameter(parametersetname="Export Type")][validateset("Just Import","Import and Export")][Parameter(Mandatory=$true)]$Type,
        [Parameter(Mandatory=$true)]$ExportMailbox,
        $TargetMailbox
    )

    #cleans up old exports
    function CleanExports ($name) {
        $export = (Get-MailboxExportRequest | Where-Object {$_.name -like "$name"})
        if ($export -eq $null) {
            #no export found
        }
        else {
            "Deleting existing export request."
            Get-MailboxExportRequest | Where-Object {$_.name -like "$name"} | Remove-MailboxExportRequest
        }
    }

    #cleans up old imports
    function CleanImports ($name) {
        $import = (Get-MailboxImportRequest | Where-Object {$_.name -like "$name"})
        if ($import -eq $null) {
            #no export found
        }
        else {
            "Deleting existing import request."
            Get-MailboxImportRequest | Where-Object {$_.name -like "$name"} | Remove-MailboxImportRequest
        }
    }

    # opens up the gui and allows user to select a folder to export
    function SelectMailboxFolder ($emailaddress) {
        ((Get-MailboxFolderStatistics $emailaddress).identity |Out-GridView -OutputMode Single -Title "Select Folder to export").tostring()
    }

    # checks if a mailbox exists, returns true if exists, else false
    function mailboxExists([string]$mailbox) {
        $checkMailbox = (Get-Mailbox $mailbox -erroraction SilentlyContinue);
        if ($checkMailbox -eq $null) {
            return $false
        }
        else {
            return $true
        }
    }

    # begins the export
    function startExport([string]$mailboxToExport,$tempPST,$sourceFolder) {
        CleanExports $mailboxToExport
        New-MailboxExportRequest -Name $mailboxToExport -Mailbox $mailboxToExport -FilePath $tempPST -SourceRootFolder $sourceFolder -ExcludeDumpster
    }

    # begins the import
    function startImport([string]$destinationMailbox,$tempPST) {
        CleanImports $destinationMailbox
        New-MailboxImportRequest -Name $destinationMailbox -Mailbox $destinationMailbox -TargetRootFolder "Calendar" -FilePath $tempPST -ExcludeDumpster
    }

    # begins the import (w/ source folder arg for just_import)
    function startImportAlt([string]$destinationMailbox,$PST,$sourceFolder) {
        New-MailboxImportRequest -Name $destinationMailbox -Mailbox $destinationMailbox -TargetRootFolder "Calendar" -FilePath $PST -Include -ExcludeDumpster
    }

    # blocks until the export is complete
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
        # check that all content was exported - I'll do this later if I have time
    }

    # blocks until import is complete
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
        # check that all content was imported - I'll do this later if I have time
    }

    # preforms an export of a selected mailbox and imports it into the newly created room calendar
    function export_import() {
        if (mailboxExists($mailboxToExport)) {
            "Mailbox exists, proceeding with the export."
            $folderToExport = SelectMailboxfolder -emailaddress $mailboxToExport
            if ($folderToExport -eq $null) {
                "Oops! Looks like there's no folders in this mailbox."
            }
            else {
                # replace backslashes with forwardslashses
                $folderToExport = $folderToExport.Replace("\","/")
                # add mailbox name and fix backslashes
                $sourceFolder = $folderToExport.Replace($mailboxToExport+"/","")
                if (-Not($sourceFolder -Match "Calendar")) {
                    "Oops! You must select a subfolder of Calendar or the root Calendar folder itself to import from."
                }
                else {
                    # might wanna change this later
                    $tempPST = "\\iscs-export\Export\griftayl\"+$mailboxToExport.toString()+".PST"
                    "Starting export..."
                    startExport $mailboxToExport $tempPST $sourceFolder
                    waitExport $mailboxToExport
                    "Starting import..."
                    startImport $destinationMailbox $tempPST
                    waitImport $destinationMailbox
                    "Cleaning requests"
                    CleanExports $mailboxToExport
                    CleanImports $destinationMailbox
                }
            }
        }
        else {
            "Mailbox doesn't exist, cancelling export."
        }
    }

    # does an import of an existing pst (for when the source mailbox no longer exists, but has been previously exported)
    function just_import() {
        $PST = Read-Host "Enter filepath of PST to import from"
        # I should probably do some error checking here to make sure the pst path is valid- I'll do it later
        "PST filepath is valid, proceeding with import."
        $sourceFolder = Read-Host "Enter folder within Calendar to import from (leave blank for root Calendar)"
        if ($sourceFolder -eq $null) {
            $sourceFolder = "Calendar"
        }
        else {
            $sourceFolder = "Calendar/"+$sourceFolder
        }
        "Starting import..."
        startImportAlt $destinationMailbox $sourcePST $sourceFolder
        waitImport $destinationMailbox
        "Cleaning import request"
        CleanImports $destinationMailbox
    }

    switch ($PSBoundParameters.Values)
    {
        'Just Import' {just_import}
        'Import And Export' {export_import}  #find a way to check for target mailbox parameter?
    }
}
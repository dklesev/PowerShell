# Add Params Here

param(
[string]$Sender = "***@mail.com",
[string]$FilePath = "C:\temp\VEEAM\",
[string]$SearchFolder = "VEEAM"
)

if(!(Test-Path $FilePath)){
       New-Item -ItemType directory -Path $FilePath
    }

# Today's Date
$tdate = Get-Date -format d

Function MoveMail ($items, $folder){
    if($items.Items.Count){
        $FirstItem = $items.Items.GetFirst()
        if($FirstItem.SenderEmailAddress -match $Sender){
            $FileName = $FirstItem.Subject.replace('[', '').replace(']','')
            $FileName
            [void]$FirstItem.SaveAs($FilePath + $FileName + ".htm", $olSaveType::olHTML)
        }
        [void]$FirstItem.Move($folder)
        MoveMail $items $folder
    }
}

Function MergeFiles ($PathToFolder){
    $ReportsFolder = $PathToFolder + "\Reports"
    if(!(Test-Path $ReportsFolder)){
       New-Item -ItemType directory -Path $ReportsFolder
    }
    $Files = $PathToFolder + '*?.htm' | Get-Item -filter $_
    $Final = $ReportsFolder + '\report_'+ $tdate +'.html'
    if($Files){
        Get-Content $Files | Add-Content $Final
    }
}

Function Clean {
    $ReportsFolder = $FilePath + "Reports"
    
    $dItems = $FilePath + '*?' | Get-Item -filter $_
    
    $dItems | foreach{
        $dIt = $FilePath + $_.Name
        If (Test-Path $dIt){
            if(!($dIt -eq $ReportsFolder)){
                Remove-Item $dIt -Force -Recurse
            }
        }
    }
}

### Script ###

Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null

$olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type] 
$olSaveType = "Microsoft.Office.Interop.Outlook.OlSaveAsType" -as [type]

$outlook = new-object -comobject outlook.application

$namespace = $outlook.GetNameSpace("MAPI")

$inbox = $namespace.GetDefaultFolder($olFolders::olFolderInBox)

$items = $InBox.Folders.Item($SearchFolder)

$MoveToFolder = $InBox.Folders.Item($SearchFolder).Folders.Item("VEEAM_read")

MoveMail $items $MoveToFolder

MergeFiles $FilePath

Clean

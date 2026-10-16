$sourceItems = @()
$diffItems = @()
$duplicates = @()
$sourceItemNames = @()
$diffItemNames = @()
$newestVersions =@()

Function Compare-ObjectProperties {
    Param(
        [PSObject]$ReferenceObject,
        [PSObject]$DifferenceObject 
    )
    $objprops = $ReferenceObject | Get-Member -MemberType Property,NoteProperty | % Name
    $objprops += $DifferenceObject | Get-Member -MemberType Property,NoteProperty | % Name
    $objprops = $objprops | Sort | Select -Unique
    $diffs = @()
    foreach ($objprop in $objprops) {
        $diff = Compare-Object $ReferenceObject $DifferenceObject -Property $objprop
        if ($diff) {            
            $diffprops = @{
                PropertyName=$objprop
                RefValue=($diff | ? {$_.SideIndicator -eq '<='} | % $($objprop))
                DiffValue=($diff | ? {$_.SideIndicator -eq '=>'} | % $($objprop))
            }
            $diffs += New-Object PSObject -Property $diffprops
        }        
    }
    if ($diffs) {return ($diffs | Select PropertyName,RefValue,DiffValue)}     
}

$sourcePath = "\\srvrdc42\FileStorage\Documentation\Projects\Servers\AppAssure Verification\PowerShell"
$diffPath = "C:\PSscript\RR Verification"

$sourceList = Get-ChildItem -path $sourcePath -recurse
$diffList = Get-ChildItem -path $diffPath -recurse 
<#
Compare-ObjectProperties $sourceList $diffList | Sort-Object -Property CreationTime,LastWriteTime


$pathLastWriteTime = Get-ChildItem -path 'C:\PSscript\RR Verification' | Select-Object -ExpandProperty LastWriteTime
$pathCreationTime = Get-ChildItem -path 'C:\PSscript\RR Verification' | Select-Object -ExpandProperty CreationTime
$pathFullName = Get-ChildItem -path 'C:\PSscript\RR Verification' | Select-Object -ExpandProperty FullName

Write-Host $pathLastWriteTime $pathCreationTime $pathFullName
#>

$differences = Compare-Object -ReferenceObject $sourceList -DifferenceObject $diffList -Property LastWriteTime,PSChildName,Length,FullName,Directory
#$differences

ForEach ($item in $differences) {
    if ($item.Directory -like $sourcePath) {
        $sourceItems += $item
        $sourceItemNames += $item.PSChildName
    } else {
        $diffItems += $item
        $diffItemNames += $item.PSChildName
    }
    #$itemName = if ($itemFullName.StartsWith($sourcePath,1)) {$itemFullName.Replace($sourcePath,"")} else {$itemFullName.Replace($diffPath,"")}
    #Write-Host $itemFullName
    #Write-Host $itemName
    #$itemName = $item.PSChildName
    #$itemDir = $item.Directory
    #Write-Host $itemDir "\" $itemName.Trim()
}

<#
Write-Host "Source Item Names:"
$sourceItemNames
Write-Host "Differential Item Names:"
$diffItemNames
#>
ForEach ($dItem in $diffItems) {
    if ($dItem.PSChildName -in $sourceItemNames) {
        $dItemName = $dItem.PSChildName
        Write-Host "There is a match!" $dItemName
        $match = $sourceItems.Where{$_.PSChildName -like $dItemName}
        if ($dItem.LastWriteTime -gt $match.LastWriteTime) {
            Write-Host "The differential version is newer!"
            $newestVersions += $dItem
        } else {
            Write-Host "The source version is newer!"
            $newestVersions += $match
        }
    }
}

Write-Host "Here are the source items:"
$sourceItems
Write-Host "And here are the differential items:"
$diffItems

 <#
                $itemFullPath = $item.InputObject.FullName
                $itemTimeStamp = $item.InputObject.LastWriteTime
                $itemIndicator = $item.SideIndicator
                Write-Host "SRC" $itemIndicator "DLT  | " $itemFullPath "-" $itemTimeStamp
                #>
## Adapted Write-Log Function ##


# Function to generate logging
$log_directory_name = "C:\Logs\"
$log_file_name = "GUI_Disable.log"
$archive_directory = $log_directory_name + "Archive\"
If (!(Test-Path -Path $log_directory_name)) {
    New-Item -Path $log_directory_name -ItemType Directory
}
If (!(Test-Path -Path ($log_directory_name + $log_file_name))) {
    $log_file = New-Item -Path $log_directory_name -ItemType File -Name $log_file_name
    $log_path_name = $log_directory_name + $log_file_name
} Else {
    $log_file = $log_file_name
    $log_path_name = $log_directory_name + $log_file
}
Function Write-Log {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$False)]
    [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
    [String]$Level = "INFO",
    [Parameter(Mandatory=$True)]
    [String]$Message,
    [Parameter(Mandatory=$False)]
    [String]$Log_Path = $log_path_name
    )
    $date = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $line = "$date | $level | $message"
    If(((Get-Item -Path $Log_Path).Length)/1MB -ge 1) {
        $date_file_format = (Get-Date).ToString("MMddyyyy_HHmmss")
        $log_file_split = $log_file_name.Split(".")
        $log_file_short_name = $log_file_split[0]
        $log_file_extension = $log_file_split[1]
        $log_archive_path = $archive_directory + $log_file_short_name + "_" + $date_file_format + "." + $log_file_extension
        If (!(Test-Path $archive_directory)) {
            New-Item -Path $archive_directory -ItemType Directory | Out-Null
        }
        Move-Item -Path $Log_Path -Destination $log_archive_path
        Write-Verbose $line
        Add-Content $Log_Path -Value $line
    }
    Else {
        Write-Verbose $line
        Add-Content $Log_Path -Value $line
    }
}

# Uncomment the lines below to perform a test #

Do {
    Write-Log -Message "Testing logging function. Yay for logs!"
    $current_size = (Get-Item -Path $log_path_name).Length/1MB
    Write-Log -Message "Testing archival function. The current file size is $current_size."
    $archive_check = Test-Path $archive_directory
    $exit_condition = If ($archive_check -eq $false) {((Get-ChildItem -Path $log_directory_name | Measure-Object).count -gt 2)} Else {((Get-ChildItem -Path $archive_directory | Measure-Object).count -gt 2)}
} Until ($exit_condition)

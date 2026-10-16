# Set the variables 
$WebURL = “http://sewellautomotive.sharepoint.com/sites/serversonly” 
$DocLibName = “Shared Documents/General” 
#$FilePath = “C:\Docs\stuff\Secret Sauce.docx” 

# Get a variable that points to the folder 
$Web = Get-SPWeb $WebURL 
$List = $Web.GetFolder($DocLibName) 
$Files = $List.Files

$Files

# Get just the name of the file from the whole path 
#$FileName = $FilePath.Substring($FilePath.LastIndexOf("\")+1)

# Load the file into a variable 
#$File= Get-ChildItem $FilePath

# Upload it to SharePoint 
#$Files.Add($DocLibName +"/" + $FileName,$File.OpenRead(),$false) 
#$web.Dispose()
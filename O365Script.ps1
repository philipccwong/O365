#Set-ExecutionPolicy RemoteSigned (as admin if not set)

function connectToO365{

Add-Type -Path "C:\SharePoint_PowerShell\O365\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\SharePoint_PowerShell\O365\Microsoft.SharePoint.Client.Runtime.dll"

 
    $adminUrl = Read-Host "Enter the Admin URL of 0365 (eg. https://<Tenant Name>-admin.sharepoint.com)"
    $userName = Read-Host "Enter the username of 0365 (eg. admin@<tenantName>.onmicrosoft.com)"
    $password = Read-Host "Please enter the password for $($userName)" -AsSecureString
    $credentials = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $userName, $password
    $SPOCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userName, $password)
 
    try{
        Connect-SPOService -Url $adminUrl -Credential $credentials
        write-host "Info: Connected succesfully to Office 365" -foregroundcolor green }
 
    catch{
        write-host "Error: Could not connect to Office 365" -foregroundcolor red
        Break connectToO365
    }
  
    $filePath = create-outputfile 
    get-siteCollections
    add-content -value "
 
    </body>
 
    </html>
 
    " -path $filePath
 }


 
function create-outputfile(){
  
    $date = get-date -format dMMyyyyhhmm
    $filePath = "$($PSScriptRoot)\O365_Output$($date).html"

    if (!(Test-Path -path $filePath)){
  
    New-Item $filePath -type file | out-null 
    write-host “File created: $($filePath)” -foregroundcolor green
    add-content -value "
 
    <html>
    <body>
    <h1>Sites information Office 365</h1>
    <table border='1' style='font-family: Calibri, sans-serif'>
    <tr>
    <th style='background-color:blue; color:white'>Site Collection</th>
    <th style='background-color:blue; color:white'>Sub-site</th>
    <th style='background-color:blue; color:white'>List/Library</th>
    <th style='background-color:blue; color:white'>Total Items</th>
    <th style='background-color:blue; color:white'>Size</th>
    <th style='background-color:blue; color:white'>Folder Count</th>
    <th style='background-color:blue; color:white'>Content Type Count</th>
    <th style='background-color:blue; color:white'>Document Set Count</th>
    </tr>
 
    " -path $filePath
 
    }
 
    else{
    write-host "Output file already exists, wait 1 minute" -foregroundcolor yellow
    Break create-outputfile
    }
 
    return $filePath
 
}
 
function get-siteCollections{
 
    $siteColl_count = 1
    $subColl_count = 1
    $siteCollections = Get-SPOSite
  
    foreach ($siteCollection in $siteCollections)
    {
 
        $pixelsweb = 0
        $pixelslist = 0
        add-content -value "<tr style='background-color:cyan'><td>$($siteCollection.url)</td><td>Top Level</td><td>Type: $($sitecollection.template)</td><td></td><td></td><td></td><td></td><td></td></tr>" -path $filePath
        write-host "Info: Found $($siteCollection.url)" -foregroundcolor green
        $AllWebs = Get-SPOWebs $siteCollection.url  $siteColl_count $subColl_count
        $siteColl_count = $siteColl_count + 1
    }
 
}
 


function Get-SPOWebs($url, $siteColl_count, $subColl_count){
 
    $add = "WA01SC"+$siteColl_count
    $add1 =  "WA01SC"+$siteColl_count +"SS"+$subColl_count
    $list_count = 1
    $showContentType = 1 
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($url)
    $context.Credentials = $SPOcredentials
    $web = $context.Web
    $context.Load($web)
    $context.Load($web.Webs)
    $context.load($web.lists)
 
    try{
 
        $context.ExecuteQuery()
        foreach($list in $web.lists){

              if($list -ne $null)
              {
                     if($list.BaseTemplate -eq "101")             
                     {

                                $a = "WA01SC"+$siteColl_count +"SS"+$subColl_count + "DL" + $list_count
                                $b = $list.ItemCount
                                
                                if ($showContentType)
                                {
                                
                                $context.Load($list.ContentTypes)
                                #$context.Load($list.Folders)
                                
                                try{
                                $context.ExecuteQuery()
                                }
                                catch
                                {}
                               

                                $contentTypelist = $list.ContentTypes

                                if ($contentTypelist -ne $null) {
                                $e = $contentTypelist.Count
                                }
                                else
                                {
                                $e = 0
                                }
                                
                                }
                                
                                if ($contentTypelist -ne $null) {
                                
                                   $listSize = 0
                                   $DocSetFolderCount = 0
                                   $FolderCount = 0
                                   $DocumentSetCount = 0
                                   $qry = [Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery()
                                   $items = $list.GetItems($qry)
                                   $context.Load($items)
                                   $context.ExecuteQuery()
                                

                                    foreach ($item in $items) 
                                    { 
                                    if (($item["FSObjType"] -eq "1") -and ($item["HTML_x0020_File_x0020_Type"] -eq $null))
                                    {
                                        $FolderCount = $FolderCount + 1
                                    }   
                                    if ($item["HTML_x0020_File_x0020_Type"] -eq "SharePoint.DocumentSet")
                                     {   

                                        $DocumentSetCount = $DocumentSetCount + 1
                                     }
                                     $listSize += ($item["File_x0020_Size"])
                                    }
                                    $c = [Math]::Round(($listSize/1KB),2)     
                                
                                }
                                write-host "WA01, $add, $add1, $a, $b, $c, $FolderCount, $e, $DocumentSetCount"
                                add-content -value "<tr><td><span style='margin-left:$($pixelslist)px'>$($add)</td><td>$($add1)</td><td>$($a)</td><td>$($b)</td><td>$($c)</td><td>$($FolderCount)</td><td>$($e)</td><td>$($DocumentSetCount)</td></tr>" -path $filePath

                                $list_count = $list_count + 1
                            }
               
                        }  

                        
            }
 
            $pixelsweb = $pixelsweb + 15
            $pixelslist = $pixelslist + 15
 
            foreach($web in $web.Webs) {
 
            add-content -value "<tr style='background-color:yellow'><td><span style='margin-left:$($pixelsweb)px'>$($web.url)</td><td>Sub-Site</td><td>Type: $($web.webtemplate)</td><td></td><td></td><td></td><td></td><td></td></tr>" -path $filePath
            write-host "Info: Found $($web.url)" -foregroundcolor green
            $subColl_count = $subColl_count + 1
            Get-SPOWebs $web.url $siteColl_count $subColl_count
 
            }
 
    }
 
    catch{
 
        write-host "Could not find web" -foregroundcolor red
 
    }
 
}
 
 #Main
connectToO365
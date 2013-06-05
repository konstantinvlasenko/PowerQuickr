$PQ_DIR = Split-Path ($MyInvocation.MyCommand.Path)

function Merge-Tokens($template, $tokens)
{
  return [regex]::Replace(
      $template,
      '\$(?<tokenName>\w+)\$',
      {
          param($match)
          $tokenName = $match.Groups['tokenName'].Value
          return $tokens[$tokenName]
      })    
}

function Convert-EntriesToObject {
  BEGIN{ 
    $Object = New-Object PSObject
  }
  PROCESS {
    foreach ($entry in $input) {
      $Object | add-member Noteproperty $entry.title.'#text' (Convert-EntryToHash $entry)
    }
  }
  END{
    $Object
  }
}

function Convert-EntryToHash {
  param(
    [parameter(Mandatory = $true, ValueFromPipeline=$true)]
    [System.Xml.XmlElement] $entry
  )
  @{url = $entry.content.src}
}

function Set-Quickr {
  param(
    [parameter(Mandatory = $true)]
    [string] $server,
    [parameter(Mandatory = $true)]
    [string] $user,
    [parameter(Mandatory = $true)]
    [string] $password
  )
  Set-Variable -Name "PQ_BASE" -Value "http://$server" -Scope "Global"
  Set-Variable -Name "PQ_LIBRARIES" -Value "http://$server/dm/atom/libraries/feed" -Scope "Global"
  Set-Variable -Name "PQ_ADMIN" -Value $user -Scope "Global"
  Set-Variable -Name "PQ_ADMIN_PASSWORD" -Value $password -Scope "Global"
  $encoded = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes("$user`:$password"))
  $PQ_HEADERS = @{Authorization = "Basic $encoded"}
  Set-Variable -Name "PQ_HEADERS" -Value @{Authorization = "Basic $encoded"} -Scope "Global"
}

function New-QuickrPlace {
  <#
  .SYNOPSIS
  Creates Quickr place
  .DESCRIPTION
  Creates Quickr place (http://wiki.urspringer.de/doku.php/qfaq/tutorials_howtos_etc/how_to_create_a_quickr_place_automatically)
  .EXAMPLE
  New-QuickrPlace TestPlace "John Doe" passw0rd
  .PARAMETER place
  The name of the Place.
  .PARAMETER owner
  The Place's owner.
  .PARAMETER owner_password
  The Place's owner password.
  #>
  param(
    [parameter(Mandatory = $true)]
    [string] $place,
    [string] $owner = $PQ_ADMIN,
    [string] $owner_password = $PQ_ADMIN_PASSWORD
  )
  
  $url = "$PQ_BASE/LotusQuickr/LotusQuickr/CreateHaiku.nsf?OpenDatabase&PresetFields=h_SetEditCurrentScene;h_CreateManager,h_EditAction;h_Next,h_SetCommand;h_CreateOffice,h_PlaceTypeName;,h_Name;$place,h_UserName;$owner,h_SetPassword;$owner_password,h_EmailAddress;,h_SetReturnUrl;$PQ_BASE/$place,h_OwnerAuth;h_External"
  $res = Invoke-WebRequest -Uri $url -Headers $PQ_HEADERS
  if($res.StatusCode -eq 200) {
    Get-QuickrPlace $place
  }
  else {
    $res
  }
}

function Get-QuickrPlace {
  param(
    [parameter(Mandatory = $true)]
    [string] $place
  )
  $url = "$PQ_BASE/dm/atom/library/%5B@P$place/@RMain.nsf%5D/feed"
  ([xml](Invoke-WebRequest -Uri $url -Headers $PQ_HEADERS).Content).feed.entry | Convert-EntriesToObject
}

function New-QuickrFolder {
  param(
    [parameter(Mandatory = $true)]
    [string] $name,
    [parameter(Mandatory = $true, ValueFromPipeline=$true)]
    [hashtable[]] $parents
  )
  BEGIN {
    $template = [string](Get-Content "$PQ_DIR\xml\create_folder.xml")
  }
  PROCESS {
    foreach($p in $parents) {
      $url = "$PQ_BASE$($p.url)"
      $body = Merge-Tokens $template @{ base = $url; name = $name}
      ([xml](Invoke-WebRequest -Uri $url -Method Post -Body $body -ContentType "application/atom+xml" -Headers $PQ_HEADERS).Content).entry | Convert-EntryToHash
    }
  }
}

function New-QuickrDocument {
  param(
    [parameter(Mandatory = $true, ValueFromPipeline=$true)]
    [hashtable] $parent,
    [parameter(Mandatory = $false)]
    [string] $name = "PowerQuickr.txt",
    [parameter(Mandatory = $false)]
    [Byte[]] $content = [system.Text.Encoding]::UTF8.GetBytes('Hello PowerQuickr!')
    
  )
  $header = @{Slug = $name}
  $url = "$PQ_BASE$($parent.url)"
  (Invoke-WebRequest -Uri $url -Method Post -Body $content -ContentType "application/atom+xml" -Headers ($PQ_HEADERS + $header)).Content
}

export-modulemember -function Set-Quickr
export-modulemember -function Get-QuickrPlace
export-modulemember -function New-QuickrPlace
export-modulemember -function New-QuickrRootFolder
export-modulemember -function New-QuickrFolder
export-modulemember -function New-QuickrDocument
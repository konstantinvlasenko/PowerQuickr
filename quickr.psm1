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
      $Object | add-member Noteproperty $entry.title.'#text' $entry.content.src
    }
  }
  END{
    $Object
  }
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
    [parameter(Mandatory = $true)]
    [string] $owner,
    [parameter(Mandatory = $true)]
    [string] $owner_password
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

function Get-QuickrLibraries {
  $url = "$PQ_BASE/dm/atom/libraries/feed"
  ([xml](Invoke-WebRequest -Uri $url -Headers $PQ_HEADERS).Content).feed.entry | Convert-EntriesToObject
}

function Get-QuickrPlace {
  param(
    [parameter(Mandatory = $true)]
    [string] $place
  )
  $url = iex "(Get-QuickrLibraries).$place"
  ([xml](Invoke-WebRequest -Uri $url -Headers $PQ_HEADERS).Content).feed.entry | Convert-EntriesToObject
}

function New-QuickrFolder {
  param(
    [parameter(Mandatory = $true)]
    [string] $parentUrl,
    [parameter(Mandatory = $true)]
    [string] $name
  )
  $url = "$PQ_BASE$parentUrl"
  $template = [string](Get-Content "$PQ_DIR\xml\create_folder.xml")
  $body = Merge-Tokens $template @{ base = $url; name = $name}
  ([xml](Invoke-WebRequest -Uri $url -Method Post -Body $body -ContentType "application/atom+xml" -Headers $PQ_HEADERS).Content).entry | Convert-EntriesToObject
}


function New-QuickrDocument {
  param(
    [parameter(Mandatory = $true, ValueFromPipeline=$true)]
    [xml[]] $folders,
    [parameter(Mandatory = $false)]
    [string] $name = "PowerQuickr.txt",
    [parameter(Mandatory = $false)]
    [Byte[]] $content = [system.Text.Encoding]::UTF8.GetBytes('Hello PowerQuickr!')
    
  )
  BEGIN{ 
    $header = @{Slug = $name}
  }
  PROCESS {
    foreach ($f in $folders) {
      $url = "$PQ_BASE$($f.entry.content.src)"
      (Invoke-WebRequest -Uri $url -Method Post -Body $content -ContentType "application/atom+xml" -Headers ($PQ_HEADERS + $header)).Content
    }
  }
}

export-modulemember -function Set-Quickr
export-modulemember -function Get-QuickrLibraries
export-modulemember -function Get-QuickrPlace
export-modulemember -function New-QuickrPlace
export-modulemember -function New-QuickrRootFolder
export-modulemember -function New-QuickrFolder
export-modulemember -function New-QuickrDocument
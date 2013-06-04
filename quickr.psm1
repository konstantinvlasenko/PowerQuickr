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
  Invoke-WebRequest -Uri $url -Headers $PQ_HEADERS
}


function New-QuickrFolder {
  param(
    [parameter(Mandatory = $true)]
    [string] $place,
    [parameter(Mandatory = $true)]
    [string] $name
  )
  $template = [string](Get-Content "$PQ_DIR\xml\create_folder.xml")
  $body = Merge-Tokens $template @{ base = $PQ_BASE; place = $place;  name = $name}
  $url = "$PQ_BASE/dm/atom/library/%5B@P$place/@RMain.nsf%5D/feed"
  Invoke-WebRequest -Uri $url -Method Post -Body $body -ContentType "application/atom+xml" -Headers $PQ_HEADERS
}

export-modulemember -function Set-Quickr
export-modulemember -function New-QuickrPlace
export-modulemember -function New-QuickrFolder
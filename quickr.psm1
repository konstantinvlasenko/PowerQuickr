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

function New-QuickrFolder {
  param(
    [parameter(Mandatory = $true)]
    [string] $place,
    [parameter(Mandatory = $true)]
    [string] $name
  )
  $PQ_BASE | out-default
  if($PQ_BASE -eq $null) {
    "Call Set-Quickr first!"
    return
  }
  $template = [string](Get-Content "$PQ_DIR\xml\create_folder.xml")
  $body = Merge-Tokens $template @{ base = $PQ_BASE; place = $place;  name = $name}
  $url = "$PQ_BASE/dm/atom/library/%5B@P$place/@RMain.nsf%5D/feed"
  Invoke-WebRequest -Uri $url -Method Post -Body $body -ContentType "application/atom+xml" -Headers $PQ_HEADERS
}

export-modulemember -function Set-Quickr
export-modulemember -function New-QuickrFolder
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
  Set-Variable -Name "PQ_HEADERS" -Value @{Authorization = "Basic $encoded"; "Content-Language" = "en"} -Scope "Global"
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

function Login-QuickrPlace {
  param(
    [parameter(Mandatory = $true)]
    [string] $place
  )
  $url = "$PQ_BASE/LotusQuickr/$place/Main.nsf`?Login"
  Invoke-WebRequest $url -Method post -ContentType 'application/x-www-form-urlencoded' -Body "Username=$PQ_ADMIN&password=$PQ_ADMIN_PASSWORD" -SessionVariable PQ | OUT-NULL
  $PQ.Cookies
  
}

function Get-QuickrPlace {
  param(
    [parameter(Mandatory = $true)]
    [string] $place
  )
  
  $url = "$PQ_BASE/dm/services/ContentService?wsdl"
  $global:proxy = New-WebServiceProxy -uri $url
  $global:proxy.url = "$PQ_BASE/dm/services/DocumentService"
  $global:proxy.CookieContainer = Login-QuickrPlace $place
   
  
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
      $url = "$PQ_BASE/$($p.url)"
	  $url | out-default
      $body = Merge-Tokens $template @{ base = $url; name = $name}
      ([xml](Invoke-WebRequest -Uri $url -Method Post -Body $body -ContentType "application/atom+xml" -Headers $PQ_HEADERS).Content).entry | Convert-EntryToHash
    }
  }
}

function Lock-QuickrDocument {
  param(
    [parameter(Mandatory = $true)]
    [string] $place,
	[parameter(Mandatory = $true)]
    $page
  )
  $url = "$PQ_BASE/dm/services/ContentService?wsdl"
  $proxy = New-WebServiceProxy -uri $url
  $proxy.url = "$PQ_BASE/dm/services/DocumentService"
  $proxy.CookieContainer = Login-QuickrPlace $place
  $id = $page.id.split(':')[-1] 
  $proxy.lockDocument($id, $null)
}

function Unlock-QuickrDocument {
  param(
    [parameter(Mandatory = $true)]
    [string] $place,
	[parameter(Mandatory = $true)]
    $page
  )
  $url = "$PQ_BASE/dm/services/ContentService?wsdl"
  $proxy = New-WebServiceProxy -uri $url
  $proxy.url = "$PQ_BASE/dm/services/DocumentService"
  $proxy.CookieContainer = Login-QuickrPlace $place
  $id = $page.id.split(':')[-1] 
  $proxy.unlockDocument($id, $null)
}

function New-QuickrDocument {
  param(
    [parameter(Mandatory = $true, ValueFromPipeline=$true)]
    [hashtable] $parent,
    [string] $name = "PowerQuickr.txt",
    [Byte[]] $content = [system.Text.Encoding]::UTF8.GetBytes('Hello PowerQuickr!')
    
  )
  $header = @{Slug = $name}
  $url = "$PQ_BASE/$($parent.url)"
  (Invoke-WebRequest -Uri $url -Method Post -Body $content -ContentType "text/plain" -Headers ($PQ_HEADERS + $header)).StatusCode
}

function New-QuickrPage {
  param(
    [parameter(Mandatory = $true, ValueFromPipeline=$true)]
    [hashtable] $parent,
    [string] $name,
    [string] $content
    
  )
  $url = "$PQ_BASE/$($parent.url)?doctype=[@D30DF3123AEFAF358052567080016723D]"
  $xml = "
	<?xml version='1.0' encoding='utf-8'?>
	<a:entry xmlns:a='http://www.w3.org/2005/Atom'>
		<title type='text'>$name</title>
		<a:label>$name</a:label>
		<a:category scheme='tag:ibm.com,2006:td/type' term='page' label='page' />
		<a:summary type='text'>$content</a:summary>
		<a:content type='text'>
			Created by PowerQuickr
		</a:content>	
	</a:entry>
	"
  $payload = [system.Text.Encoding]::UTF8.GetBytes($xml) 	
  $entry = ([xml](Invoke-WebRequest -Uri $url -Method Post -Body $payload -ContentType "text/xml" -Headers $PQ_HEADERS)).entry
  
  $id = $entry.id.split(':')[-1] 
  $proxy.getDocument($id,$null,$null,$null).document
}

function Update-Document {
  param(
    [parameter(Mandatory = $true)]
    $document
  )
  $proxy.updateDocument($document).error
}

function New-QuickrDocumentVersion {
  param(
    [parameter(Mandatory = $true)]
    $document,
    [string] $comments
    
  )
  $proxy.createVersion($null,$document.path, $comments)
}

export-modulemember -function Set-Quickr
export-modulemember -function Get-QuickrPlace
export-modulemember -function New-QuickrPlace
export-modulemember -function New-QuickrRootFolder
export-modulemember -function New-QuickrFolder
export-modulemember -function New-QuickrDocument
export-modulemember -function New-QuickrPage
export-modulemember -function New-QuickrDocumentVersion
export-modulemember -function Lock-QuickrDocument
export-modulemember -function Unlock-QuickrDocument
export-modulemember -function Update-Document


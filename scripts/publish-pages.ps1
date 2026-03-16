param(
  [Parameter(Mandatory = $true)]
  [string]$BaseUrl,

  [string]$ProjectRoot = "",

  [string]$OutputDir = ""
)

$ErrorActionPreference = "Stop"

function Join-Url {
  param(
    [Parameter(Mandatory = $true)]
    [string]$Root,

    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  return ([System.Uri]::new(([System.Uri]::new($Root.TrimEnd("/") + "/")), $Path.TrimStart("/")).AbsoluteUri)
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
  $ProjectRoot = (Resolve-Path (Join-Path $scriptRoot "..")).Path
}

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = Join-Path $ProjectRoot "publish/pages"
}

$resolvedProjectRoot = (Resolve-Path $ProjectRoot).Path
$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)
$normalizedBaseUrl = $BaseUrl.TrimEnd("/")

if (-not (Test-Path (Join-Path $resolvedProjectRoot "wwwroot"))) {
  throw "Could not find wwwroot under '$resolvedProjectRoot'."
}

New-Item -ItemType Directory -Force -Path $resolvedOutputDir | Out-Null
Get-ChildItem -Force -Path $resolvedOutputDir | Remove-Item -Recurse -Force

Copy-Item -Path (Join-Path $resolvedProjectRoot "wwwroot\\*") -Destination $resolvedOutputDir -Recurse -Force

$manifestPath = Join-Path $resolvedProjectRoot "manifest.xml"
[xml]$manifest = Get-Content -Path $manifestPath

$namespaceManager = New-Object System.Xml.XmlNamespaceManager($manifest.NameTable)
$namespaceManager.AddNamespace("o", "http://schemas.microsoft.com/office/appforoffice/1.1")
$namespaceManager.AddNamespace("ov", "http://schemas.microsoft.com/office/mailappversionoverrides")
$namespaceManager.AddNamespace("bt", "http://schemas.microsoft.com/office/officeappbasictypes/1.0")
$namespaceManager.AddNamespace("xsi", "http://www.w3.org/2001/XMLSchema-instance")

$urlMappings = @{
  "o:IconUrl" = "assets/icon-80.png"
  "o:HighResolutionIconUrl" = "assets/icon-80.png"
  "o:SupportUrl" = "taskpane.html"
  "o:FormSettings/o:Form[@xsi:type='ItemRead']/o:DesktopSettings/o:SourceLocation" = "taskpane.html"
  "o:FormSettings/o:Form[@xsi:type='ItemEdit']/o:DesktopSettings/o:SourceLocation" = "taskpane.html"
  "ov:Resources/bt:Images/bt:Image[@id='Icon.16']" = "assets/icon-16.png"
  "ov:Resources/bt:Images/bt:Image[@id='Icon.32']" = "assets/icon-32.png"
  "ov:Resources/bt:Images/bt:Image[@id='Icon.80']" = "assets/icon-80.png"
  "ov:Resources/bt:Urls/bt:Url[@id='Taskpane.Url']" = "taskpane.html"
  "ov:Resources/bt:Urls/bt:Url[@id='Commands.Url']" = "commands.html"
}

foreach ($mapping in $urlMappings.GetEnumerator()) {
  $node = $manifest.SelectSingleNode("//$($mapping.Key)", $namespaceManager)
  if ($null -eq $node) {
    throw "Manifest node '$($mapping.Key)' was not found."
  }

  $node.SetAttribute("DefaultValue", (Join-Url -Root $normalizedBaseUrl -Path $mapping.Value))
}

$appDomainsNode = $manifest.SelectSingleNode("//o:AppDomains", $namespaceManager)
if ($null -eq $appDomainsNode) {
  throw "Manifest AppDomains node was not found."
}

$existingDomains = $appDomainsNode.SelectNodes("./o:AppDomain", $namespaceManager)
foreach ($domainNode in @($existingDomains)) {
  [void]$appDomainsNode.RemoveChild($domainNode)
}

$appDomainElement = $manifest.CreateElement("AppDomain", "http://schemas.microsoft.com/office/appforoffice/1.1")
$appDomainElement.InnerText = $normalizedBaseUrl
[void]$appDomainsNode.AppendChild($appDomainElement)

$manifest.Save((Join-Path $resolvedOutputDir "manifest.xml"))
Set-Content -Path (Join-Path $resolvedOutputDir ".nojekyll") -Value "" -NoNewline

Write-Host "Pages package written to $resolvedOutputDir"
Write-Host "Manifest base URL: $normalizedBaseUrl"

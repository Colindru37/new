. .\edge.ps1
$page = (New-Object Uri ((Resolve-Path .\selftest\page.html).Path)).AbsoluteUri
$page
edge open $page
edge url
edge title

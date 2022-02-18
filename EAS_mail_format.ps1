

$final_result = @()
$fullnames = @()
$Result1 = @()
$cn = "LP-16WD533"
for($i=0 ; $i -lt 2 ; $i++)
{
if (test-Connection -ComputerName $cn -Count 2)
{
$Yesterday = Get-ChildItem -Recurse -path "C:\Users\dudekula.ameer\Documents\checking\Yesterday" 
$Yesterday_count = $Yesterday | Measure-Object -Property Length -Sum | Select-Object @{Name="Size(MB)";Expression={("{0:N2}" -f($_.Sum/1mb))}}, Count
$Today = Get-ChildItem -Recurse -path "C:\Users\dudekula.ameer\Documents\checking\Today" 
$Today_count = $Today | Measure-Object -Property Length -Sum | Select-Object @{Name="Size(MB)";Expression={("{0:N2}" -f($_.Sum/1mb))}}, Count

$delta_files = $Today_count.count - $Yesterday_count.count
$delta_size = $Today_count.'Size(MB)' - $Yesterday_count.'Size(MB)' 

$result = New-Object PSObject
$result | Add-Member -MemberType NoteProperty -Name "Yesterday Count" -Value $Yesterday_count.Count
$result | Add-Member -MemberType NoteProperty -Name "Today Count" -Value $Today_count.Count
$result | Add-Member -MemberType NoteProperty -Name "Delta Count files" -Value $delta_files
$result | Add-Member -MemberType NoteProperty -Name "Yesterday Size" -Value $Yesterday_count.'Size(MB)'
$result | Add-Member -MemberType NoteProperty -Name "Today Size" -Value $Today_count.'Size(MB)'
$result | Add-Member -MemberType NoteProperty -Name "Delta of Size" -Value $delta_size
$final_result += $result

$compare_folders = Compare-Object $Yesterday $Today | select -ExpandProperty InputObject
$fullnames = $compare_folders.FullName
if($delta_files -lt 0)
{
$CompareCou = "Decreased"
}
else{
$CompareCou = "Increased"
}
foreach($temp in $fullnames)
{
$result1 += [PSCustomObject] @{
"Computer Name" = "$cn"
"Full Path" = "$temp"
"Increase/Decrease" = "$CompareCou"
}
}
}
else
{
"$computer is not in online"
}
}
$Result1 | Export-Csv "C:\Users\dudekula.ameer\Documents\checking\result.csv" -NoTypeInformation

$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
</style>
"@

$Body = $final_result | ConvertTo-Html -Head $Header

$EmailBody=@"
<br>
Hi Team,
<br>
<br>
Here is the attachment file. PFA
<br>
<br>
$Body

<br>

<br>

Thanks

"@

$Outlook = New-Object -comobject Outlook.Application
$mailitem=$Outlook.CreateItem("olmailitem")
$mailitem.to="polamreddy.sowmya@hcl.com"
$mailitem.subject="EAS Tower-GSK"
$mailitem.HTMLBody = $EmailBody
$mailitem.Attachments.add("C:\Users\dudekula.ameer\Documents\checking\result.csv")
$mailitem.send()
 
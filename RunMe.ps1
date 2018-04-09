$siteCollectionURL = Read-Host "`n`nPlease enter site collection URL"
$jobRoles = Read-Host "`n`nPlease enter the number of Job Roles"
$jobTitles = Read-Host "`n`nPlease enter the number of Job Titles"
$courses = Read-Host "`n`nPlease enter the number of Courses"
$employees = Read-Host "`n`nPlease enter the number of Employees"
$activities = Read-Host "`n`nPlease enter the number of Activities"

sleep 5
.\MassUploader.exe $siteCollectionURL $jobRoles $jobTitles $courses
sleep 5
Write-Host Creating Employees
.\CreateEmployees.exe $siteCollectionURL $employees 
Write-Host Done Employees
sleep 5
$Site = get-SPSite $siteCollectionURL
$rootweb = $Site.rootweb
$EmployeeList = $rootweb.Lists["Employee List"]
$EmployeeListCount = $EmployeeList.Items.Count 
$count = 0
Write-Host Creating Employee TMM Folders
foreach ($Employee in $EmployeeList.Items)
{
	$EmployeeStatus = $Employee["Employee Status"]
	if ($EmployeeStatus -eq "Not Active")
	{
		$Employee["Employee Status"] = "Active"
		$Employee.Update()
		$count++
		if ($count%10 -eq 0){
			WRITE-HOST $count
		}
		sleep 1
	}
}
Write-Host Done Employee TMM Folders
sleep 5
$tmmSite = $siteCollectionURL + "TMM/"
$Web = get-SPWeb $tmmSite
$gaps = $Web.Lists["Training Gaps"].Items.Count
if ($gaps -ge $activities){
	Write-Host Creating Activities
	.\GapsToActivities.exe $siteCollectionURL $activities $gaps
	Write-Host Done Activities
}
elseif ($gaps -lt $activities){
	Write-Host Not enough Gaps to be converted to requested amount of Activities. Please create more Gaps and try again
	Write-Host Closing in 10 seconds...
	sleep 10
	Exit
}
sleep 5
Write-Host List population complete
Write-Host Starting Document Uploader
[string]$l = Get-Location 
$location = $l + "\Documents"
.\LoadEffective.exe $siteCollectionURL $location
Write-Host Done Document Uploading 


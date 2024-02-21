# Initialize HTML report
$report = @()

# Get host information
$cs = Get-WmiObject -Class Win32_ComputerSystem
$os = Get-WmiObject -Class Win32_OperatingSystem
$processors = Get-WmiObject -Class Win32_Processor
$disks = Get-WmiObject -Class Win32_LogicalDisk -Filter "DriveType = 3"
$networks = Get-WmiObject -Class Win32_NetworkAdapterConfiguration -Filter "IPEnabled = TRUE"

# Write host information to report
$report += "<h2>Host Information</h2>"
$report += "<p>Name: $($cs.Name)</p>"
$report += "<p>Manufacturer: $($cs.Manufacturer)</p>"
$report += "<p>OS: $($os.Caption) - $($os.Version)</p>"
$report += "<p>Processors: $(($processors | Measure-Object).Count)</p>"
$report += "<p>Logical Disks: $(($disks | Measure-Object).Count)</p>"
$report += "<p>IPEnabled Network Adapters: $(($networks | Measure-Object).Count)</p>"

# Get VM information without the filter
$vms = Get-WmiObject -Namespace root\virtualization -Class Msvm_ComputerSystem

# Write VM information to report
foreach ($vm in $vms)
{
  $report += "<h2>VM Information: $($vm.ElementName)</h2>"
  $report += "<h3>Status: $($vm.EnabledState)</h3>"
}

# Export report to HTML file
$report | Out-File -FilePath C:\HyperVReport.html

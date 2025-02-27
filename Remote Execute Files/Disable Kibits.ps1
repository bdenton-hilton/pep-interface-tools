Write-Host "Starting The IFC Disable Script"
Write-Host "Getting  The List Of IFC Services "
$IfcServices = Get-service -Name OnQIFC*
Write-Host "Found $($IfcServices.count) IFC Services"
if($IfcServices.count -gt 0)
{
    foreach($Ifcservice in $IfcServices)
    {
       if($Ifcservice.status -eq "Running")
       {
            Write-Host "Getting List Of Process ID's" 
            $ProcessIDs= (Get-WmiObject -Class Win32_Service |Where-Object{$_.Name -like $Ifcservice}).ProcessId
            Write-Host "There are $($ProcessIDs.count) associated with the service $($Ifcservice.Name)"
            if($ProcessIDs.count -gt 0)
            {
                    foreach($processId in $ProcessIDs)
                    {
                        Write-Host "Stopping The Process $($processId)"
                        Stop-Process $processId -Force 
                        start-sleep 5
                    }
            }
            Write-Host "Stopping The Service $($Ifcservice.name)"
            Stop-Service -Name $Ifcservice.Name -Force 
            if( $?)
            {
                Write-Host "$($Ifcservice.Name) Stopped Succesfully So Disabling The Service"
                     start-sleep 5
                Set-Service  -Name $Ifcservice.name  -StartupType Disabled 
                if($?)
                {
                    Write-Host "$($Ifcservice.Name) Disabled Succesfully"
                }
                else 
                {
                    Write-Host "$($Ifcservice.name) Fail To Disable , Please Check Manually"
                }

            }
            else {
                Write-Host "$($Ifcservice.name) Failed while Stopping , Please Check Manually"
            }
   
        }
    }
}
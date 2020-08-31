<#
Powershell script to kill process that is locking a file.

Thanks to:
https://devblogs.microsoft.com/scripting/batchman-writes-a-powershell-script-to-automate-handle/
https://mcpmag.com/articles/2018/07/10/check-for-locked-file-using-powershell.aspx

This script requires the presence of handle.exe in path
https://docs.microsoft.com/en-us/sysinternals/downloads/handle

#>
Function Kill-Process-Locking-File {
    [cmdletbinding()]
    Param (
        [parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [Alias('FullName','PSPath')]
        [string[]]$Path
    )
    Process {
        ForEach ($Item in $Path) {
            #Ensure this is a full path
            $Item = Convert-Path $Item
            #Verify that this is a file and not a directory
            If ([System.IO.File]::Exists($Item)) {
                Try {
                    $FileStream = [System.IO.File]::Open($Item,'Open','Write')
                    $FileStream.Close()
                    $FileStream.Dispose()
                    $IsLocked = $False
                } Catch [System.UnauthorizedAccessException] {
                    $IsLocked = 'AccessDenied'
                } Catch {
                    $IsLocked = $True
		    Write-Output "is locked"
		    $ScreenOutput = C:\Users\jjacob\handle -a -u "$Item"
		    $abc= $ScreenOutput|SELECT-STRING -pattern 'pid: [\w]*' 
                    
                    $pidIndex=$abc[0].matches[0].Index
                    $pidLength=$abc[0].matches[0].Length
                    $processToKill = $abc[0].tostring().substring($pidIndex+4,$pidLength-4)
                    Stop-Process -Id $processToKill 
                }
                
            }
        }
    }
}

Kill-Process-Locking-File -Path 'C:\Users\jjacob\test.docx'

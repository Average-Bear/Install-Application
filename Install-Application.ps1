function Install-Application {

<#     
.SYNOPSIS     
    Copies and installs specifed filepath ($Path). This serves as a template for the following filetypes: .EXE, .MSI, & .MSP 

.DESCRIPTION     
    Copies and installs specifed filepath ($Path). This serves as a template for the following filetypes: .EXE, .MSI, & .MSP

.EXAMPLE    
    Install-Application (Get-Content C:\ComputerList.txt)

.EXAMPLE    
    Install-Application Computer1, Computer2, Computer3   

.NOTES   
    Author: JBear 
    Date: 2/9/2017 
   
    Edited: JBear
    Date: 10/13/2017 
#> 

    param(

        [Parameter(Mandatory=$true,HelpMessage="Enter Computername(s)")]
        [String[]]$Computername,

        [Parameter(ValueFromPipeline=$true,HelpMessage="Enter installer path(s)")]
        [String[]]$Path = $null,

        [Parameter(ValueFromPipeline=$true,HelpMessage="Enter remote destination: C$\Directory")]
        $Destination = "C$\TempApplications"
    )

    if($Path -eq $null) {

        Add-Type -AssemblyName System.Windows.Forms

        $Dialog = New-Object System.Windows.Forms.OpenFileDialog
        $Dialog.InitialDirectory = "\\FileShare01\IT\Applications"
        $Dialog.Title = "Select Installation File(s)"
        $Dialog.Filter = "Installation Files (*.exe,*.msi,*.msp)|*.exe; *.msi; *.msp"        
        $Dialog.Multiselect=$true
        $Result = $Dialog.ShowDialog()

        if($Result -eq 'OK') {

            Try {
      
                $Path = $Dialog.FileNames
            }

            Catch {

                $Path = $null
	            Break
            }
        }

        else {

            #Shows upon cancellation of Save Menu
            Write-Host -ForegroundColor Yellow "Notice: No file(s) selected."
            Break
        }
    }

    #Create function    
    function InstallAsJob {

        #Each item in $Computernam variable        
        foreach($Computer in $Computername) {

            #If $Computer IS NOT null or only whitespace
            if(!([string]::IsNullOrWhiteSpace($Computer))) {

                #Test-Connection to $Computer
                if(Test-Connection -Quiet -Count 1 $Computer) {                                               
                    
                    #Create job on localhost
                    Start-Job { param($Computer, $Path, $Destination)

                        foreach($P in $Path) {
                         
                            #Static Temp location
                            $TempDir = "\\$Computer\$Destination"

                            #Create $TempDir directory
                            if(!(Test-Path $TempDir)) {

                                New-Item -Type Directory $TempDir | Out-Null
                            }
                    
                            #Retrieve Leaf object from $Path
                            $FileName = (Split-Path -Path $P -Leaf)

                            #New Executable Path
                            $Executable = "C:\$(Split-Path -Path $Destination -Leaf)\$FileName"

                            #Copy needed installer files to remote machine
                            Copy-Item -Path $P -Destination $TempDir

                            #Install .EXE
                            if($FileName -like "*.exe") {

                                function InvokeEXE {

                                    Invoke-Command -ComputerName $Computer { param($TempDir, $FileName, $Executable)
                                   
                                        Try {

                                            #Start EXE file
                                            Start-Process $Executable -ArgumentList "/s" -Wait -NoNewWindow   
                                            Write-Output "`n$FileName installation complete on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName installation failed on $env:computername."
                                        }

                                        Try {
                                   
                                            #Remove $TempDir location from remote machine
                                            Remove-Item -Path $Executable -Recurse -Force
                                            Write-Output "`n$FileName source file successfully removed on $env:computername."
                                        }

                                        Catch {
                                        
                                            Write-Output "`n$FileName source file removal failed on $env:computername."    
                                        }  
                                    } -AsJob -JobName "Silent EXE Install" -ArgumentList $TempDir, $FileName, $Executable
                                }

                                InvokeEXE | Receive-Job -Wait
                            }
                               
                            #Install .MSI                                        
                            elseif($FileName -like "*.msi") {

                                function InvokeMSI {

                                    Invoke-Command -ComputerName $Computer { param($TempDir, $FileName, $Executable)

                                        Try {
                                       
                                            #Start MSI file                                    
                                            Start-Process 'msiexec.exe' "/i $Executable /qn" -Wait -ErrorAction Stop
                                            Write-Output "`n$FileName installation complete on $env:computername."
                                        }

                                        Catch {
                                       
                                            Write-Output "`n$FileName installation failed on $env:computername."
                                        }

                                        Try {
                                   
                                            #Remove $TempDir location from remote machine
                                            Remove-Item -Path $Executable -Recurse -Force
                                            Write-Output "`n$FileName source file successfully removed on $env:computername."
                                        }

                                        Catch {
                                       
                                            Write-Output "`n$FileName source file removal failed on $env:computername."    
                                        }                              
                                    } -AsJob -JobName "Silent MSI Install" -ArgumentList $TempDir, $FileName, $Executable                            
                                }

                                InvokeMSI | Receive-Job -Wait
                            }

                            #Install .MSP
                            elseif($FileName -like "*.msp") { 
                                                                     
                                function InvokeMSP {

                                    Invoke-Command -ComputerName $Computer { param($TempDir, $FileName, $Executable)

                                        Try {
                                                                              
                                            #Start MSP file                                    
                                            Start-Process 'msiexec.exe' "/p $Executable /qn" -Wait -ErrorAction Stop
                                            Write-Output "`n$FileName installation complete on $env:computername."
                                        }

                                        Catch {
                                      
                                            Write-Output "`n$FileName installation failed on $env:computername."
                                        }

                                        Try {
                                   
                                            #Remove $TempDir location from remote machine
                                            Remove-Item -Path $Executable -Recurse -Force
                                            Write-Output "`n$FileName source file successfully removed on $env:computername."
                                        }

                                        Catch {                                      

                                            Write-Output "`n$FileName source file removal failed on $env:computername."    
                                        }                             
                                    } -AsJob -JobName "Silent MSP Installer" -ArgumentList $TempDir, $FileName, $Executable
                                }

                                InvokeMSP | Receive-Job -Wait
                            }

                            else {

                                Write-Host "$Destination has an unsupported file extension. Please try again."                        
                            }
                        }                      
                    } -Name "Application Install" -Argumentlist $Computer, $Path, $Destination            
                }
                                           
                else {                                
                  
                    Write-Host "Unable to connect to $Computer."                
                }            
            }        
        }   
    }

    #Call main function
    InstallAsJob
}#End Install-Application

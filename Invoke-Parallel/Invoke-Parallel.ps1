
Function Recycle-AppPoolParallel
{

	<#
	.Synopsis

	.Description

	.Example
	 PS c:\> 'IISServer1','IISServer2','IISServer3' | Recycle-AppPoolParallel -AppPoolNames 'AppPool1','AppPool2'  

	.Link
	 https://groups.google.com/forum/#!topic/microsoft.public.windows.powershell/qmKyifvwBso

	.Notes
	 NAME:      Recycle-AppPoolParallel
	 AUTHOR:    NSCORP\zspd6
	 LASTEDIT:  2/1/2018
	 #Requires -Version 3.0
	#>

	[CmdletBinding()]
	param(
	[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
	[Alias('cn')]
	[string]$ComputerName,
	[Parameter(Position=1, Mandatory=$true)]
	$AppPoolNames,
	[Parameter(Position=2)]
	$ThrottleLimit = 20
	)

    Begin
    {
      $Computers = @()
      if(!(Get-Command Invoke-Parallel)){
        Write-Host "This function requires the Invoke-Parallel Function which can be found at...`nhttps://github.com/RamblingCookieMonster/Invoke-Parallel/blob/master/Invoke-Parallel/Invoke-Parallel.ps1" -ForegroundColor Red
        break
      }#if
    }

	Process{$Computers += $ComputerName}

    End
    {

      #Function payload is defined in the "End" section to collect all pipeline parameters prior to script execution.
      $ElapsedTime = Measure-Command {

        $Computers | Invoke-Parallel -Throttle $ThrottleLimit -ScriptBlock {

          #Used functions must be defigned inside of the Invoke-Parallel ScriptBlock
          Function FriendlyErrorString ($thisError) {
            [string] $Return = $thisError.Exception
            $Return += "`r`n"
            $Return += "At line:" + $thisError.InvocationInfo.ScriptLineNumber
            $Return += " char:" + $thisError.InvocationInfo.OffsetInLine
            $Return += " For: " + $thisError.InvocationInfo.Line
            Return $Return
          }

          $cn = $_
          if(Test-Connection -Quiet -Count 1 -ComputerName $cn){
        
           $Output = foreach($AppPool in $Using:AppPoolNames){
           	
             try{
               ############ PAYLOAD ##########
               #Replacement for Get-WmiObject
               $Result =(Get-CIMInstance `
                 -Namespace "root\MicrosoftIISv2" `
                 -ClassName "IIsApplicationPool" `
                 -ComputerName $cn `
                 | Where {$_.Name -eq "W3SVC/APPPOOLS/$AppPool"}).Recycle()
               
               #region ----- Output as a Custom Object ------
                [pscustomobject][ordered]@{
                  #'Property'=$Value 
                   AppPoolName=$AppPool
                   Result=$Result
                   ComputerName=$cn
                  #-----------------
                }
               #endregion -- End Output as a Custom Object --
               ###############################
             }catch{
                [string]$ErrorString = FriendlyErrorString $Error[0]
                Write-Error $ErrorString
             }#try

	       }#foreach
        
          }else{
            $Output = "Could not reach $cn."
          }#if(Test-Connection..
           
          #region -- Garbage Collection / Variable Cleanup --
            $GC = 'AppPool','cn','Result','ErrorString','Return'
            
            foreach($var in $GC){Remove-Variable -Name $var -ErrorAction SilentlyContinue}
          #endregion ----------------------------------------
        
        }#Invoke-Parallel

      }#Measure-Command

      $Output | Sort-Object
      
      if($($ElapsedTime.Minutes) -like '0'){
        $ElapsedTime = $ElapsedTime.ToString().Substring(6,5) 
        Write-Verbose "Recycle-AppPoolParallel function completed in $ElapsedTime seconds."
      }else{
        $ElapsedTime = $ElapsedTime.ToString().Substring(4,7)
        Write-Verbose "Recycle-AppPoolParallel function completed in $($ElapsedTime.split(':')[0]) minutes & $($ElapsedTime.split(':')[-1]) seconds."
      }#if
      
      #region -- Garbage Collection / Variable Cleanup --
        $GC = 'ComputerName','AppPoolNames','Output','ElapsedTime','ThrottleLimit'
        
        foreach($var in $GC){Remove-Variable -Name $var -ErrorAction SilentlyContinue}
      #endregion ----------------------------------------

    }#End

}# End Function Recycle-AppPoolParallel


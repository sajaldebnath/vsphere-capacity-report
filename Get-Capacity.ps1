#requires -version 3

###########################################################################################
# Title     :   Capactiy Audit Report for VMware Environment
# Filename  :   Get-Capacity.ps1          
# Created by:   Sajal Debnath           
# Date      :   01-09-2015                
# Version   :   1.0        
# Update    :   This is the first version
# E-mail    :   debnathsajal@gmail.com
###########################################################################################

<# 
    .Synopsis 
   Get-Capacity does a capacity audit of a vSphere environment. 
    .DESCRIPTION 
   The Get-Capacity function is designed to audit capacity aspects of a vSphere environment. It will check
   the vCenter, ESXi hosts and all the Datastores of the given environment.
 
   By default the function will create the report in a HTML format in the path from where it was called

    .NOTES 
   Created by: Sajal Debnath 
   Modified: 9/7/2015 10:29:58 PM  
 
   Changelog: 
    *  

 
   To Do: 
    * Create a front end form where users will be able to choose the Clusters on which the Capacity test will be done
    * Create separate functions for calculating capacity at different layers
    * Create seprate function to get the HTML output
    * Create more proper Verbose and Debug output
    * Create more detailed logging 
    * Take input from credential file
 
 
    .EXAMPLE 
    Get-Capacity -vcenter <vcenter> -vcuser <vcenter user> -vcpassword <vcenter password> -consolidation <CPU consolidation ratio>

#> 




Param(
    [Parameter (Mandatory=$true, ValueFromPipeline=$false)]
    [String]$vcenter,
    [Parameter (Mandatory=$true, ValueFromPipeline=$false)]
    [String]$vcuser,
    [Parameter (Mandatory=$true, ValueFromPipeline=$false)]
    [String]$vcpassword,
    [Parameter (Mandatory=$true, ValueFromPipeline=$false)]
    [String]$consolidation

)


#########################################################################

$version = 1.0
# Set the SMTP Server address
# $SMTPSRV = "10.25.114.246"
# Set the Email address to recieve from
# $EmailFrom = "debnathsajal@gmail.com"
# Set the Email address to send the email to
# $EmailTo =  "debnath.sajal@hotmail.com"
# Use the following item to define if the output should be displayed in the local browser once completed

$DisplaytoScreen = "YES"
# Use the following item to define if an email report should be sent once completed
$SendEmail = "YES"
# Silencing the comments in the HTML report
$CommentsH = $false
$CommentsA = $false
$CommentsB = $false

# Use the following area to define the colours of the report
#$Colour1 = "CC0000" # Main Title - currently red
$Colour1 = "F87217" # Main Title - currently Dark Orange1
#$Colour2 = "7BA7C7" # Secondary Title - currently blue
$Colour2 = "82CAFA" # Secondary Title - currently Light Sky blue
$Colour3 = "FFF380" # Secondary Sub Title - Khaki1
$Colour4 = "FFF8C6" # Secondary main Title - Lemon Chiffon
$Colour5 = "58D3F7"

#### Detail Settings ####

$date = Get-Date

## Report Base path definition

$LogFile = '.\Get-Security.log'

#######################################
# Functions definitions
#######################################

# Log Write Function 

function Write-Log 
{ 
    [CmdletBinding()] 
    #[Alias('wl')] 
    [OutputType([int])] 
    Param 
    ( 
        # The string to be written to the log. 
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true, 
                   Position=0)] 
        [ValidateNotNullOrEmpty()] 
        [Alias("LogContent")] 
        [string]$Message, 
 
        # The path to the log file. 
        [Parameter(Mandatory=$false, 
                   ValueFromPipelineByPropertyName=$true, 
                   Position=1)] 
        [Alias('LogPath')] 
        [string]$Path=".\Get-Capacity.log", 
 
        [Parameter(Mandatory=$false, 
                    ValueFromPipelineByPropertyName=$true, 
                    Position=3)] 
        [ValidateSet("Error","Warn","Info")] 
        [string]$Level="Info", 
 
        [Parameter(Mandatory=$false)] 
        [switch]$NoClobber 
    ) 
 
    Process 
    { 
         
        if ((Test-Path $Path) -AND $NoClobber) { 
            Write-Warning "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name." 
            Return 
            } 
 
        # If attempting to write to a log file in a folder/path that doesn't exist 
        # to create the file include path. 
        elseif (!(Test-Path $Path)) { 
            Write-Verbose "Creating $Path." 
            $NewLogFile = New-Item $Path -Force -ItemType File 
            } 
 
        else { 
            # Nothing to see here yet. 
            } 
 
        # Now do the logging and additional output based on $Level 
        switch ($Level) { 
            'Error' { 
                Write-Error $Message 
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") ERROR: $Message" | Out-File -FilePath $Path -Append 
                } 
            'Warn' { 
                Write-Warning $Message 
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") WARNING: $Message" | Out-File -FilePath $Path -Append 
                } 
            'Info' { 
                Write-Verbose $Message 
                Write-Output "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss") INFO: $Message" | Out-File -FilePath $Path -Append 
                } 
            } 
    } 
    End 
    { 
    } 
}


function Get-CustomHTML ($Header){
$Report = @"
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Frameset//EN" "http://www.w3.org/TR/html4/frameset.dtd">
<html><head><title>$($Header)</title>
        <META http-equiv=Content-Type content='text/html; charset=windows-1252'>

        <style type="text/css">

        TABLE       {
                        TABLE-LAYOUT: fixed; 
                        FONT-SIZE: 100%; 
                        WIDTH: 100%
                    }
        *
                    {
                        margin:0
                    }
        .hidden { display: none; }
        .unhidden { display: block; }
        .dspcont    {
    
                        BORDER-RIGHT: #bbbbbb 1px solid;
                        BORDER-TOP: #bbbbbb 1px solid;
                        PADDING-LEFT: 0px;
                        FONT-SIZE: 8pt;
                        MARGIN-BOTTOM: -1px;
                        PADDING-BOTTOM: 5px;
                        MARGIN-LEFT: 0px;
                        BORDER-LEFT: #bbbbbb 1px solid;
                        WIDTH: 95%;
                        COLOR: #000000;
                        MARGIN-RIGHT: 0px;
                        PADDING-TOP: 4px;
                        BORDER-BOTTOM: #bbbbbb 1px solid;
                        FONT-FAMILY: Tahoma;
                        POSITION: relative;
                        BACKGROUND-COLOR: #f9f9f9
                    }
                    
        .filler     {
                        BORDER-RIGHT: medium none; 
                        BORDER-TOP: medium none; 
                        DISPLAY: block; 
                        BACKGROUND: none transparent scroll repeat 0% 0%; 
                        MARGIN-BOTTOM: -1px; 
                        FONT: 100%/8px Tahoma; 
                        MARGIN-LEFT: 43px; 
                        BORDER-LEFT: medium none; 
                        COLOR: #FFFFFF; 
                        MARGIN-RIGHT: 0px; 
                        PADDING-TOP: 4px; 
                        BORDER-BOTTOM: medium none; 
                        POSITION: relative
                    }

        .pageholder {
                        margin: 0px auto;
                    }
                    
        .dsp
                    {
                        BORDER-RIGHT: #bbbbbb 1px solid;
                        PADDING-RIGHT: 0px;
                        BORDER-TOP: #bbbbbb 1px solid;
                        DISPLAY: block;
                        PADDING-LEFT: 0px;
                        FONT-WEIGHT: bold;
                        FONT-SIZE: 8pt;
                        MARGIN-BOTTOM: -1px;
                        MARGIN-LEFT: 0px;
                        BORDER-LEFT: #bbbbbb 1px solid;
                        COLOR: #000000;
                        MARGIN-RIGHT: 0px;
                        PADDING-TOP: 4px;
                        BORDER-BOTTOM: #bbbbbb 1px solid;
                        FONT-FAMILY: Tahoma;
                        POSITION: relative;
                        HEIGHT: 2.25em;
                        WIDTH: 95%;
                        TEXT-INDENT: 10px;
                    }

        .dsphead0   {
                        BACKGROUND-COLOR: #$($Colour1);
                    }
                    
        .dsphead1   {
                        
                        BACKGROUND-COLOR: #$($Colour2);
                    }
        .dsphead2   {
                        
                        BACKGROUND-COLOR: #$($Colour3);
                    }
        .dsphead3   {
                        
                        BACKGROUND-COLOR: #$($Colour4);
                    }     

        .dsphead4   {
                        BACKGROUND-COLOR: #$($Colour5);
                    }
                          
                    
    .dspcomments    {
                        BACKGROUND-COLOR:#FFFFE1;
                        COLOR: #000000;
                        FONT-STYLE: ITALIC;
                        FONT-WEIGHT: normal;
                        FONT-SIZE: 8pt;
                    }

    td              {
                        VERTICAL-ALIGN: TOP; 
                        FONT-FAMILY: Tahoma
                    }
                    
    th              {
                        VERTICAL-ALIGN: TOP; 
                        COLOR: #$($Colour1); 
                        TEXT-ALIGN: left
                    }
                    
    BODY            {
                        margin-left: 4pt;
                        margin-right: 4pt;
                        margin-top: 6pt;
                    } 
    .MainTitle      {
                        font-family:Arial, Helvetica, sans-serif;
                        font-weight:bolder;
                        colour: #FF8040;
                    }
    .SubTitle       {
                        font-family:Arial, Helvetica, sans-serif;
                        font-size:14px;
                        font-weight:bold;
                    }
    .Created        {
                        font-family:Arial, Helvetica, sans-serif;
                        font-size:10px;
                        font-weight:normal;
                        margin-top: 20px;
                        margin-bottom:5px;
                    }
    .links          {   font:Arial, Helvetica, sans-serif;
                        font-size:10px;
                        FONT-STYLE: ITALIC;
                    }
                    
        </style>
        <script type="text/javascript">
         function unhide(divID) {
         var item = document.getElementById(divID);
        if (item) {
        item.className=(item.className=='hidden')?'unhidden':'hidden';
        }
        }
        </script>       
    
    </head>
    <body >
        <div class="MainTitle"><font size="20ptx"><center>$($Header)</center></font></div>
        <hr size="8" color="#$($Colour1)">
        <div class="SubTitle">Infrastructure Capacity Report Version:$($version):: Generated on $($ENV:Computername):&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp</div>
        <br/>
        <hr size="4" color="#$($Colour1)">
        <br/>
        <div class="Created">Report created on $(get-date)</div>

"@
return $Report
}

function Get-CustomHeader0 ($Title){
$Report = @"
        <div class="pageholder">        

        <h1 class="dsp dsphead0">$($Title)</h1>
    
        <div class="filler"></div>

"@
return $Report
}


function Get-CustomHeader1 ($Title){
$Report = @"
        <div class="pageholder">        

        <h1 class="dsp dsphead4">$($Title)</h1>
    
        <div class="filler"></div>

"@
return $Report
}

function Get-CustomHeader ($Title, $cmnt, $Div){
$Report = @"
        <h2 class="dsp dsphead1">$($Title)</h2>
"@
if ($CommentsH) {
    $Report += @"
            <div class="dsp dspcomments">$($cmnt)</div>
"@
}
$Report += @"
        <div class="dspcont">
        <a href="javascript:unhide('$($Div)');"><button>Show/Hide</button></a>
        <div id="$($Div)" class="hidden">
"@
return $Report
}

function Get-CustomHeaderA ($Title, $cmnt, $Div){
$Report = @"
        <h2 class="dsp dsphead2">$($Title)</h2>
"@
if ($CommentsA) {
    $Report += @"
            <div class="dsp dspcomments">$($cmnt)</div>
"@
}
$Report += @"
        <div class="dspcont">
        <a href="javascript:unhide('$($Div)');"><button>Show/Hide</button></a>
        <div id="$($Div)" class="hidden">
"@
return $Report
}

function Get-CustomHeaderB ($Title, $cmnt){
$Report = @"
        <h2 class="dsp dsphead3">$($Title)</h2>
"@
if ($CommentsB) {
    $Report += @"
            <div class="dsp dspcomments">$($cmnt)</div>
"@
}
$Report += @"
        <div class="dspcont">
"@
return $Report
}


function Get-CustomHeaderClose{

    $Report = @"
        </DIV>
        </div>
        <div class="filler"></div>
"@
return $Report
}


function Get-CustomHeaderAClose{

    $Report = @"
        </DIV>
        </div>
        <div class="filler"></div>
"@
return $Report
}

function Get-CustomHeaderBClose{

    $Report = @"
        </DIV>

        <div class="filler"></div>
"@
return $Report
}



function Get-CustomHeader0Close{
    $Report = @"
</DIV>
"@
return $Report
}


function Get-CustomHeader1Close{
    $Report = @"
</DIV>
"@
return $Report
}



function Get-CustomHTMLClose{
    $Report = @"
</div>

</body>
</html>
"@
return $Report
}


function Get-HTMLTable {
    param([array]$Content)
    $HTMLTable = $Content | ConvertTo-Html -Fragment
    return $HTMLTable
}



################################################################################################################################
# Start of script
################################################################################################################################


# Adding Snapin

Try{
    Write-Verbose 'Loading the Required Snapins'
    Write-Log  -LogContent 'Loading the Required Snapins' -Level Info
    
    Add-PSSnapin VMware.VimAutomation.Core -ea SilentlyContinue
   }
    
 Catch{
     Write-Debug 'Could Not Load the Snapin'
     Write-Log  -LogContent 'Could Not Load the Snapin'  -Level Error
     Exit
   
   }

# Connecting to vCenter Server

Try{
     Write-Verbose 'Connecting to vCenter Server'
     Write-Log  -LogContent 'Connecting to vCenter Server' -LogPath $LogFile -Level Info
     [void] (Connect-VIServer -Server $vcenter -User $vcuser -Password $vcpassword)
     $vcfullname = (Get-View ServiceInstance).Content.About.FullName
    }
    
 Catch{
     Write-Debug 'Could Not Connect to vCenter Server'
     Write-Log  -LogContent 'Could Not Connect to vCenter Server' -LogPath $LogFile -Level Error
     Exit
    }



##################################################################
# Find out which version,build and fullname of vCenter Server
##################################################################

$vcversion = (Get-View ServiceInstance).Content.About.Version
$vcbuild = (Get-View ServiceInstance).Content.About.Build

##################################################################
# Start of Main Reporting 
##################################################################


# Data place holders declaration

$alldata = @()
$cludetail = @()
$esxdetails = @()
$datadetails = @()

# Setting the variables for Total Capacity Calculation

$Etotalcpu = 0
$Etotalusedcpu = 0
$Etotalavalcpu = 0
$Etotalmem = 0
$Etotalusedmem = 0
$Etotalavalmem = 0
$Etotalstorage = 0
$Etotalavalstorage = 0
$Etotalusedstorage = 0
$Eallovcpu = 0
$Eallomem = 0
$Eallovm = 0 
$EmemUsage = 0
$EhostUsage = 0

# Getting clusters

$clusters = Get-Cluster | Sort Name

$clusternumber = $clusters.Count



$detail = "" | Select "Physical Core","Allo. vCPU","Used vCPU","Host Usage %","Usage %","vCPU Remaining","Physical Mem","Used Mem","Aval MEM","Used Mem %","Total VM","Allo. MEM","Total Storage-GB","Used Storage-GB","Aval. Storage-GB","Used Storage-%"

# Calculating the vales

foreach ($clu in $clusters){
	$name = $clu.Name

   $clupart = "" | Select Name,"Physical Core","Allo. vCPU","Used vCPU","Host Usage %","Usage %","vCPU Remaining","Physical Mem","Used Mem","Aval MEM","Used Mem %","Total Storage-GB","Used Storage-GB","Aval. Storage-GB","Used Storage-%","Total VM","Allo. MEM"


   # Setting the variables for Cluster Capacity Calculation

    $Ctotalcpu = 0
    $Ctotalavalcpu = 0
    $Ctotalusedcpu = 0
    $Ctotalmem = 0
    $Ctotalavalmem = 0
    $Ctotalusedmem = 0
    $Ctotalstorage = 0
    $Ctotalavalstorage = 0
    $Ctotalusedstorage = 0
    $Callovcpu = 0
    $Callomem = 0
    $Callovm = 0 
    $CpCPUUsage = 0
    $CmemUsage = 0
    $ChostUsage = 0

    # Finding how many hosts are in this cluster

    $hosts =  Get-VMHost -Location $clu  | Sort MemoryUsageGB -Descending 

    $hostnumber = $hosts.Count
    

    Foreach ($host1 in $hosts) { 


        $esxdetail = "" | Select Name,CName,"Physical Core","Allo. vCPU","Used vCPU","Host Usage %","Usage %","vCPU Remaining","Physical Mem","Used Mem","Aval MEM","Used Mem %","Total VM","Allo. MEM"

        $esxdetail.Name = $host1.Name
        $esxdetail.CName = $clu.Name

        # Calculating number of Physical Cores depending on whether Hyperthreading is available and enabled or not

        if ($host1.HyperthreadingActive){
            $esxdetail."Physical Core" += $host1.NumCpu * 2

        }
        else {
            $esxdetail."Physical Core" += $host1.NumCpu 
        }

        $esxdetail."Physical MEM" = [math]::round($host1.MemoryTotalGB,0)


        For ($i = $hostnumber - $Clu.HAFailoverLevel;$i -gt 0;$i-- ){
                $Ctotalcpu += $esxdetail."Physical Core"
                $Ctotalmem += $esxdetail."Physical MEM"
               
        }

        # Calculating server CPU speed in MGHz so that we can use it to calculate equivalent vCPU's
	
        $speed = [math]::round($host1.ExtensionData.Hardware.CpuInfo.Hz /1000 / 1000,0)

        # Calculating usage history of the host for the last 7 days

        $stats = Get-Stat -Entity $host1 -start (get-date).AddDays(-7) -Finish (Get-Date) -MaxSamples 10000 -stat "cpu.usagemhz.average","cpu.usage.average","mem.consumed.average","mem.usage.average"

        $stats | Group-Object -Property Entity | %{
        $cpu = $_.Group | where {$_.MetricId -eq "cpu.usagemhz.average"} | Measure-Object -Property value -Average -Maximum -Minimum
        $acpu = $_.Group | where {$_.MetricId -eq "cpu.usage.average"} | Measure-Object -Property value -Average -Maximum -Minimum
        $mem = $_.Group | where {$_.MetricId -eq "mem.consumed.average"} | Measure-Object -Property value -Average -Maximum -Minimum
        $amem = $_.Group | where {$_.MetricId -eq "mem.usage.average"} | Measure-Object -Property value -Average -Maximum -Minimum
        }

        # $cpu gives averge CPU utilization in MGhz
        # $acpu gives the average CPU usage percentage
        # $mem gives averge memory utilization in KB
        # $amem gives the average memory usage percentage


        # Calculating information for the VM's on the host
        $vms = Get-VM -Location $host1

        # Calculating total allocated vCPU and Memory to the VM's
        $allovcpu = 0
        $allomem = 0

        foreach ($vm in $vms){

            $allovcpu += $vm.NumCpu
            $allomem += $vm.MemoryGB
        }

        # Average CPU usage in terms of vCPU usage per host in the last 7 days. It also contributes to total CPU usage in the cluster in terms of vCPU in last 7 days

        $esxdetail."Used vCPU" = [math]::round($cpu.Maximum / $speed,0) 
        $Ctotalusedcpu += $esxdetail."Used vCPU"


	    $esxdetail."Usage %" = [math]::round($acpu.Maximum,2)

        $esxdetail."Host Usage %" = [math]::round($acpu.Maximum,2)

        $esxdetail."Used MEM" = [math]::round(($mem.Maximum/1024)/1024 ,0) 
        $Ctotalusedmem += $esxdetail."Used MEM"


        $esxdetail."Used Mem %" = [math]::round($amem.Maximum,2) 

        # Calculating total available vCPU in the server considering the CPU consolidation ratio

        $esxdetail."vCPU Remaining" = ($esxdetail."Physical Core" * $consolidation ) - $esxdetail."Used vCPU"

        $esxdetail."Aval MEM" = $esxdetail."Physical MEM" - $esxdetail."Used MEM"

        $esxdetail."Total VM" = $vms.count

        $esxdetail."Allo. vCPU" = $allovcpu

        $esxdetail."Allo. MEM" = $allomem

        # Gathering Host information for this host
        $esxdetails += $esxdetail
        
        # Adding up the information of this host to the total cluster value

        $Callovm += $vms.count
        $Callovcpu += $allovcpu
        $Callomem += $allomem
        $CpCPUUsage += $esxdetail."Usage %"
        $CmemUsage += $esxdetail."Used Mem %"
        $ChostUsage += $esxdetail."Host Usage %"

        }
    
   # Calculate Datastore wise information, not considering the local datastores

    $datastores = Get-Datastore -RelatedObject $clu | Sort Name

    Foreach($datastore in $datastores) { 
        $name = $datastore.Name

        if ( $name.StartsWith("datastore1") ) {
       
        }
        else {	
		

			$datadetail = "" | Select Name,CName,"Total Capacity-GB","Used Capacity-GB","Available Capacity-GB","Used - %","Total VMs"
			
			$datadetail.Name = $datastore.Name
			
			$datadetail.CName = $clu.Name
			
            $datadetail."Available Capacity-GB" = [math]::round($datastore.FreeSpaceGB,0) 

            $datadetail."Total Capacity-GB" = [math]::round($datastore.CapacityGB,0)

            $datadetail."Used Capacity-GB" = ($datadetail."Total Capacity-GB" - $datadetail."Available Capacity-GB")

            $datadetail."Used - %" = [math]::round(($datadetail."Used Capacity-GB" / $datadetail."Total Capacity-GB") * 100,0) 

            $datadetail."Total VMs" = $datastore.ExtensionData.VM.Count

            # Gathering the Data store details for this cluster
            $datadetails += $datadetail

            # Calculating total storage available in this cluster
            $Ctotalavalstorage += [math]::round($datastore.FreeSpaceGB,0) 
            $Ctotalstorage += [math]::round($datastore.CapacityGB,0)
        }
   	}    

    # Calculating total available memory and used storage in this cluster

    $Ctotalavalmem = $Ctotalmem - $Ctotalusedmem
    $Ctotalusedstorage = $Ctotalstorage - $Ctotalavalstorage

    # Calculating values for the cluster 

    $clupart.Name = $clu.Name
    $clupart."Physical Core" = $Ctotalcpu
    $clupart."Allo. vCPU" = $Callovcpu
    $clupart."Used vCPU" = $Ctotalusedcpu
	$clupart."Host Usage %" = [math]::round($ChostUsage / $hostnumber,2)  
		
	$clupart."Usage %" = [math]::round($clupart."Host Usage %",2)

    $clupart."vCPU Remaining" = ($Ctotalcpu * 3) - $Ctotalusedcpu

    $clupart."Physical MEM" = $Ctotalmem
    $clupart."Used MEM" = $Ctotalusedmem
    $clupart."Aval MEM" = $Ctotalavalmem 
 
    $clupart."Used Mem %" = [math]::round($CmemUsage / $hostnumber,2)
    $clupart."Total Storage-GB" = $Ctotalstorage
    $clupart."Used Storage-GB" = $Ctotalusedstorage
    $clupart."Aval. Storage-GB" = $Ctotalavalstorage 
    $clupart."Used Storage-%" = [math]::round(($Ctotalusedstorage / $Ctotalstorage)*100,2)
    $clupart."Total VM" = $Callovm
    $clupart."Allo. MEM" = $Callomem

    $cludetail += $clupart 
    
    # Adding the values of the each cluster to the total infrastructure values so that at the end total infrastructure values can be calculated

    $Etotalcpu += $Ctotalcpu
    $Etotalusedcpu += $Ctotalusedcpu
    $Etotalmem += $Ctotalmem
    $Etotalusedmem += $Ctotalusedmem
    $Etotalavalmem += $Ctotalavalmem
    $Etotalstorage += $Ctotalstorage
    $Etotalavalstorage += $Ctotalavalstorage
    $Eallovcpu += $Callovcpu
    $Eallomem += $Callomem
    $Eallovm +=  $Callovm
    $EpCPUUsage += $clupart."Usage %"
    $EmemUsage += $clupart."Used Mem %"
    $EhostUsage += $clupart."Host Usage %"

}


# Calculating total values for the entire infrastructure
$Etotalavalcpu = $Etotalcpu - $Etotalusedcpu

$Etotalusedstorage = $Etotalstorage - $Etotalavalstorage

$detail."Total Storage-GB" = $Etotalstorage 
$detail."Used Storage-GB" = $Etotalusedstorage
$detail."Aval. Storage-GB" = $Etotalavalstorage
$detail."Used Storage-%" = [math]::round(($Etotalusedstorage / $Etotalstorage)*100,2)

$detail."Physical Core" = $Etotalcpu

$detail."Used vCPU" = $Etotalusedcpu
$detail."vCPU Remaining" = ($Etotalcpu * 3) - $detail."Used vCPU"

$detail."Host Usage %" = [math]::round($EhostUsage / $clusternumber,2)



$detail."Usage %" = [math]::round($detail."Host Usage %",2)



$detail."Allo. vCPU" = $Eallovcpu
$detail."Physical Mem" = $Etotalmem
$detail."Used Mem" = $Etotalusedmem
$detail."Aval MEM" = $Etotalavalmem

$detail."Used Mem %" = [math]::round($EmemUsage / $clusternumber,2)
$detail."Allo. MEM" = $Eallomem

$detail."Total VM" = $Eallovm

$alldata += $detail


######################################################################
#
# Start the HTML Reporting
#
#####################################################################

$MyReport = Get-CustomHTML "Infrastructure Capacity Report"
# Set the HTML Header for the report 

$MyReport += Get-CustomHeader0 ("vCenter Server: " + $vcenter.Name + "  --  " + "vCenter Version: " + $vcversion + " -- " + "Build : " + $vcbuild )


$MyReport += Get-CustomHeader1 "Total Infrastructure Capacity ----" 

# Entire Infrastructure Report 

$MyReport += Get-CustomHeaderA "Entire Infrastructure Capactiy Report ::" "" "head2"

$MyReport += Get-HTMLTable ($alldata | Select "Physical Core","Allo. vCPU","Used vCPU","vCPU Remaining","Host Usage %","Physical Mem","Allo. MEM","Used Mem","Aval MEM","Used Mem %","Total VM" )


$MyReport += Get-CustomHeaderAClose

$MyReport += Get-CustomHeader1Close



# Cluster Wise Report 


$MyReport += Get-CustomHeader1 "Infrastructure Capacity At Cluster Level ----" 

$MyReport += Get-CustomHeaderA "Cluster Wise Capactiy Report ::" "As per design HA consideration was calculated. " "head3"

$MyReport += Get-HTMLTable ( $cludetail | Select Name,"Physical Core","Allo. vCPU","Used vCPU","vCPU Remaining","Host Usage %","Physical Mem","Allo. MEM","Used Mem","Aval MEM","Used Mem %","Total VM")


$MyReport += Get-CustomHeaderAClose

$MyReport += Get-CustomHeader1Close

# Host Wise Report 

$MyReport += Get-CustomHeader1 "Infrastructure Capacity At ESXi Host Level ----" 

$MyReport += Get-CustomHeaderA "Host Wise Capactiy Report ::" "As per best practice for each host 80% capacity was taken into consideration for capacity calculation. 'Considered' cloumn shows whether the host was considered for cluster capacity calculation or not." "head4"

Foreach ($clu in $clusters) {

    $MyReport += Get-CustomHeaderB "Under Cluster :: $clu"  

    $MyReport += Get-HTMLTable ( $esxdetails | Where-Object { $_.CName -eq $clu } | Select Name,"Physical Core","Allo. vCPU","Used vCPU","vCPU Remaining","Host Usage %","Physical Mem","Allo. MEM","Used Mem","Aval MEM","Used Mem %","Total VM") 

    $MyReport += Get-CustomHeaderBClose  

}



$MyReport += Get-CustomHeaderAClose
$MyReport += Get-CustomHeader1Close

# Datastore Wise Report 

$MyReport += Get-CustomHeader1 "Infrastructure Capacity at Datastore Level ----" 

$MyReport += Get-CustomHeaderA "Datastore Wise Capactiy Report ::" "" "head5"

Foreach ($clu in $clusters) {

    $MyReport += Get-CustomHeaderB "Under Cluster :: $clu" 

    $MyReport += Get-HTMLTable ( $datadetails | Where-Object { $_.CName -eq $clu.Name } | Select Name,"Total Capacity-GB","Available Capacity-GB","Used Capacity-GB","Used - %","Total VMs") 

    $MyReport += Get-CustomHeaderBClose  
}


$MyReport += Get-CustomHeaderAClose
$MyReport += Get-CustomHeader1Close

$MyReport += Get-CustomHeader0Close
$MyReport += Get-CustomHTMLClose


#Uncomment the following lines to save the htm file in a central location
if ($DisplayToScreen) {

    $Filename = ".\"+ $vcenter + "_Capacity_Report" + "_" + $Date.Day + "-" + $Date.Month + "-" + $Date.Year +"-" + $Date.Hour + "-" + $Date.Minute + ".html"

    $MyReport | out-file -encoding ASCII -filepath $Filename
    Invoke-Item $Filename
}

######################
# E-mail HTML output #
######################

#if ($SendEmail) {
#    Write-CustomOut "Sending Email"
#    Send-MailMessage $EmailTo $EmailFrom "$VISRV Monitor Report" $SMTPSRV $MyReport
#}

[void] (Disconnect-VIServer -Server $vcenter -Force -Confirm:$false)





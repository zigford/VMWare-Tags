function Test-VIConnection {
    if (-Not $global:DefaultVIServer) {
        throw "Not connected to VCenter. Use Connect-VIServer"
    }
}

function Get-TagOrCreateIfNotPresent {
    [CmdLetBinding(SupportsShouldProcess)]
    Param($Name,$Category,$Server)

    # A Tag is a Categore + Value

    $Tag = Get-Tag -Category $Category -Name $Name `
        -EA SilentlyContinue -Server $Server

    If (-Not $Tag) {
        If ($PScmdlet.ShouldProcess(
                    "Category $Category Name $Name","Create new tag")) {
            $Tag = New-Tag -Category $Category -Name $Name -Server $Server
        } else {
            # During a whatif, let's return a dummy tag to simulate what is
            # happening
            $Tag = Get-Tag -Category $Category -Server $Server |
                   Select-Object -First 1
        }
    }

    return $Tag

}

function Update-VMTags {
<#
.SYNOPSIS
    Updates a group of tag assignments to a vm.
.DESCRIPTION
    Take a hashtable of name/values and apply them
    as category/value to a VM. Creates tags if they currently
    don't exist.
.PARAMETER VM
    Specify the VM object tags apply to
.PARAMETER TagTable
    Hashtable of tags and desired values
.EXAMPLE
    $VM = Get-VM $env:computername
    $TagTable = @{
        IT-Team-Owner = "IT Systems"
        SNC = REQ1234567
        Entered-By = "Jesse Harris"
    }
    Update-VMTags -VM $VM -TagTable $TagTable
.NOTES
    notes
.LINK
    online help
#>
    [CmdLetBinding(SupportsShouldProcess)]
    Param($VM,$TagTable)

    $VIServer = ([uri]$VM.ExtensionData.Client.ServiceUrl).Host
    $TagTable.Keys | ForEach-Object {
        # Test if tag already set
        $Cat = $_
        $TagName = $TagTable[$Cat]
        If (-Not $TagName) {return} # break out if no info
        $Tag = Get-TagOrCreateIfNotPresent `
            -Category $Cat `
            -Name $TagName `
            -Server $VIServer
        Set-TagAssignment -Tag $Tag -VM $VM
    }
}

function Set-TagAssignment {
    <#
    .SYNOPSIS
        Emulate New-TagAssignment, but sort of like fake cardinality, ensure
        only one tag of a given category is set at once
    .DESCRIPTION
        Add a new tagassignment. But if one already exists, remove it first
    .PARAMETER Tag
        A Tag as returned by Get-Tag
    .PARAMETER Entity
        A VM as returned by Get-VM
    .EXAMPLE
        $VM = Get-VM wsp-infadmin21
        $Tag = Get-Tag -Name 'IT' -Category 'Business-Owner'
        Set-TagAssignement -Tag $Tag -Entity $VM
    .NOTES
        notes
    .LINK
        online help
    #>
    [CmdLetBinding(SupportsShouldProcess)]
    Param(
            [Parameter(Mandatory=$True)]$Tag,
            [Parameter(Mandatory=$True)]$VM
         )

    # Get current tag assignments
    $TagAssignments = Get-TagAssignment -Entity $VM -Category $Tag.Category.Name
    ForEach ($TagAssignment in $TagAssignments) {
        # Oddly, Get-TagAssignment sometimes emits a null
        If ($null -eq $TagAssignment) { continue }

        If ($TagAssignment.Tag.Name -ne $Tag.Name) {
            # Tag is set incorrectly
            Write-Verbose "Tag $($Tag.Category.Name) is set to $_. removing"
            Remove-TagAssignment $TagAssignment -Confirm:$False

            # Some fricken bug in powercli caused me to write this
            $TestStillThere = Get-TagAssignment -Entity $VM `
                -Category $TagAssign.Tag.Category.Name | Where-Object {
                    $_.Tag.Name -eq $TagAssignment.Tag.Name }

            If ($TestStillThere) {
                Write-Warning ("$($TagAssignment.Tag.Category.Name) - " +
                                $($TagAssignment.Tag.Name) +
                               " not removed from VM $($VM.Name). Log " +
                                "into the gui and manually remove the tag.")
            }
        } else {
            $AlreadySet = $True
        }
    }
    If (-Not $AlreadySet) {
        $result = New-TagAssignment -Tag $Tag -Entity $VM
        if ($result) {
            Write-Verbose ("Succesfully updated $($Tag.Category.Name) " +
                            "to $($Tag.Name)")
        }
    }
}

function Get-AllVMTagTable {
    $Tags = @{}
    Get-TagAssignment | ForEach-Object {
        If (-Not ($Tags[$_.Entity.Name])) {
            $Tags[$_.Entity.Name] = @{}
        }
        $Tags[$_.Entity.Name][$_.Tag.Category.Name] = $_.Tag.Name
    }
    return $Tags
}

function Import-VMTagsFromXlsx {
    <#
    .SYNOPSIS
        Set Tags on VMs in batch taking input from an xlsx file as produced
        by the VM-ExportTags.ps1 script.
    .DESCRIPTION
        Read an xlsx file with the following columns in the main sheet.

        TagEntity   IT-Team-Owner       Entered-By      SNC... etc
        =========   ============        ==========      ====
        vmname1     Server and Storage  Ben Johnston    REQ1234567
        vmname2     IT Support          Jesse Harris    REQ7654321


        Each entry is converted into a hashtable and then input into
        the Update-VMTags function which does the work.
    .PARAMETER Path
        Path to a valid xlsx file with matching columns
    .EXAMPLE
        PS> Import-VMTagsFromXlsx -Path taglist.xlsx

    .NOTES
        notes
    .LINK
        online help
    #>
    [CmdLetBinding(SupportsShouldProcess)]
    Param(
            [System.IO.FileInfo]$Path
    )

    Import-Module ImportExcel # for importing xlsx
    Test-VIConnection

    [array]$VMsToUpdate = Import-Excel -Path $Path -WorksheetName 'Main' |
    Where-Object { $_.TagEntity -ne $null } -ErrorAction Stop

    Write-Progress -Activity "Updating VMTags" -PercentComplete 0 `
        -Status "Gathering existing tags in vcenter..."
     
    # Get all the tags and check if they have changed to speed up processing
    # This way, we will skip VM's whose tags are not changes
    $Tags = Get-AllVMTagTable

    Write-Progress -Activity "Updating VMTags" -PercentComplete 0 `
        -Status "Getting a list of VMs in vcenter..."

    # Get all the VMs in one lump sum. Many times faster than getting
    # them individually
    $AllVMs = Get-VM

    ForEach ($VMToUpdate in $VMsToUpdate) {

        $VM = $AllVMs | Where-Object { $_.Name -eq $VMToUpdate.TagEntity }
        If (-Not $VM) {
            Write-Warning "VM $($VMToUpdate.TagEntity) was not found. Skipping"
            continue
        }
        $HT = @{}
        $VMToUpdate.PSObject.Properties | Where-Object {
            $_.MemberType -eq 'NoteProperty' -and $_.Name -ne 'TagEntity'
        } | ForEach-Object {
            $HT[$_.Name] = $_.Value
        }

        # Compare hashtables to see if there is a diff on the xlsx side
        $Changed = $False
        $HT.Keys | ForEach-Object {
            If (-Not $Tags[$VM.Name]) { $Changed = $True; return }
            If ($Tags[$VM.Name][$_] -ne $HT[$_]) {
                $Changed = $True
            }
        }
        If ($Changed -eq $False ) {
            Write-Verbose "No change for $($VMToUpdate.TagEntity)"
            Write-Progress -Activity "Updating VMTags" -PercentComplete `
                (($VMsToUpdate.IndexOf($VMToUpdate)*100)/$VMsToUpdate.Count) `
                -Status "Skipping $($VMToUpdate.TagEntity). No tag changes"
            continue # Skip to next one, no change
        }

        Write-Progress -Activity "Updating VMTags" -PercentComplete `
            (($VMsToUpdate.IndexOf($VMToUpdate)*100)/$VMsToUpdate.Count) `
            -Status "Updating $($VMToUpdate.TagEntity)"
        Write-Verbose "Updating Tags of $($VMToUpdate.TagEntity)"
        Update-VMTags -TagTable $HT -VM $VM
    }
    Write-Progress -Activity "Updating VMTags" -Complete
}

function Get-VMDatacenterName {
    <#
    .SYNOPSIS
        Get the datacenter name that a vm is a memberof
    .DESCRIPTION
        Using the PowerCLI module and Get-View, walk back through the
        objects parents and return the first datacenter
    .PARAMETER VM
        Must be a VM object as returned by Get-VM
    .EXAMPLE
        Get-VMDatacenterName -VM (Get-VM wsp-infadmin21)
    .NOTES
        Thanks Tony
    .LINK
        online help
    #>
    Param($VM)
    $VMwareObjectView = Get-View -VIObject $VM
    Do
    {
        $VMwareObjectView = Get-View -Id $VMwareObjectView.Parent
    }
    Until ($VMwareObjectView.MoRef.Type -eq 'DataCenter')
    return $VMwareObjectView.Name
}

function Get-VMCluster {
    <#
    .SYNOPSIS
        Return the cluster that a VM is a member of
    .DESCRIPTION
        Full Description
    .PARAMETER VM
        Must be a VM object as returned by Get-VM
    .EXAMPLE
        Get-VMCluser -VM (Get-VM wsp-infadmin21)
    .NOTES
        notes
    .LINK
        online help
    #>
    Param($VM)

    # Get the Datacenter the VM belongs to. Need to do this because there are multiple USC clusters
    $DatacenterName = Get-VMDatacenterName -VM $VM

    # Get the cluster the VM is a member of
    $VmClusterName = $VM.VMHost.Parent.Name

    # Get the cluster object
    $DataCentre = Get-Datacenter -Name $DatacenterName
    $Cluster = $DataCentre | Get-Cluster -name $VmClusterName -EA SilentlyContinue
    
    If ($Cluster) {
        return $Cluster
    } else {
        # Some VM's are not in a cluster.
        return $null
    } 
}

function Export-VMTagsToXlsx {
    <#
    .SYNOPSIS
        Export a list of VM's and their Tags to a spreadsheet.
    .DESCRIPTION
        Create a spreadsheet with two worksheets. Worksheet 1 contains VM's
        and their tags as defined by Required Tags. This page shows all the
        current Tag assignments to the VMs. A second worksheet is populated with
        lists of available tags and data validation is setup to make it easy
        to assign tags in the spreadsheet.
    .PARAMETER Path
        Specify the path to save the Xlsx file. Supply the full path and filename
    .PARAMETER VMs
        Specify a list of VM's to export tag assignments of. If not supplyied, all
        VMs are queried.
    .EXAMPLE
        $VMs = Get-DataCenter 'Sippy Downs' | Get-VM
        Export-VMTagsToXlsx -Path MyReport.Xlsx -VMs $VMs
    .NOTES
        See Also: Import-VMTagsFromXlsx
    .LINK
        online help
    #>
    [CmdletBinding()]
    Param([Parameter(Mandatory=$True)]$Path,
          $VMs)

    Test-VIConnection
    Import-Module ImportExcel # for importing xlsx

    # Get all VMs in VMware cluster and add them to spreadsheet
    If (-Not $VMs) {
       $VMs = Get-VM
    }

    # Gather all tag info at once. (much quicker)
    $Tags = Get-AllVMTagTable

    $Categories = $Tags.Keys | % { $Tags[$_].Keys } | Sort-Object -Unique

    $ExcelPackage = $VMs | ForEach-Object {
        $VMName     = $_.Name
        $TagTable   = [ordered]@{TagEntity = $VMName}
        $Assigned = @{}
        $Categories | ForEach-Object {
            if (-Not $Tags[$VMName]) { return }
            $TagTable[$_] = $Tags[$VMName][$_]
        }

        New-Object -TypeName PSObject -Property $TagTable
     } | Export-Excel -Path $Path -AutoSize -AutoFilter `
            -WorksheetName Main -PassThru

    # Create Data validation sheet

    $ColumnNu = 1

    $Tags.Keys | %{$Tags[$_].Keys}|Sort-Object -Unique | ForEach-Object {
        $Column = [System.Collections.ArrayList]@()
        # Add the heading which is the tag category
        $Column.Add($_) | Out-Null

        # Add data of existing tags to the column
        $Category = $_
        $Tags.Keys | %{$Tags[$_][$Category]}|Sort-Object -Unique | ForEach-Object {
            $Column.Add($_) | Out-Null
        }

        $ExcelPackage = $Column | Export-Excel -ExcelPackage $ExcelPackage `
            -StartColumn $ColumnNu -WorksheetName DataValidation -PassThru

        $DataValidationParams = @{
            Worksheet           = $ExcelPackage.Main
            ShowErrorMessage    = $true
            ErrorStyle          = 'stop'
            ErrorTitle          = 'Invalid Data'
        }

        $DataValidationRules = @{
            Range               = ('{0}2:{1}1001' -f `
                                ([char]($ColumnNu+65)).ToString(),
                                ([char]($ColumnNu+65)).ToString())
            ValidationType      = 'List'
            Formula             = ('DataValidation!${0}$2:${1}$1000' -f `
                                ([char]($ColumnNu+64)).ToString(),
                                ([char]($ColumnNu+64)).ToString())
            ErrorBody           = ("You must select an item from the list.`r`n" +
                                "You can add to the list on the DataValidation page")
        }

        Add-ExcelDataValidationRule @DataValidationParams @DataValidationRules
        $ColumnNu++
    }
    Close-ExcelPackage -ExcelPackage $ExcelPackage
}

Export-ModuleMember -Function Import-VMTagsFromXlsx
Export-ModuleMember -Function Export-VMTagsToXlsx

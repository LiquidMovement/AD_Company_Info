param(
    $Deptartments = @("Accounting", "Cashiers", "Compliance", "Customer Service", "Default Depts", "Exec Admin", "Exec Team", "Facilities","HECM", "HR", "Insurance", "Investor Reporting", "Information Technology", "Origination", "Shipping", "Supervisors", "Taxes_Payoffs", "Treasury"),

    $acctNum = @(XXXX),
    $acctName = @("Accounting"),

    $cashNum = @(XXXX),
    $cashName = @("Cashiers"),

    $compNum = @(),
    $compName = @(),

    $CSNum = @(XXXX),
    $CSName = @("Customer Service"),

    $DefaultNum = @(XXXX, XXXX, XXXX, XXXX, XXXX),
    $DefaultName = @("Claims", "Default Counseling", "Foreclosure", "Loss Mitigation", "Property Preservation"),

    $ExecAdminNum = @(),
    $ExecAdminName = @(),

    $ExecTeamNum = @(),
    $ExecTeamName = @(),

    $FacilNum = @(),
    $FacilName = @(),

    $HECMNum = @(),
    $HECMName = @(),

    $HRNum = @(),
    $HRName = @(),

    $InsNum = @(XXXX),
    $InsName = @("Insurance"),

    $InvestNum = @(),
    $InvestName = @(),

    $ITNum = @(XXXX,
    $ITName = @("HelpDesk"),

    $OrigNum = @(XXXX, XXXX),
    $OrigName = @("Local Loan Officers", "National Loan Officers"),

    $ShipNum = @(),
    $ShipName = @(),

    $SupNum = @(),
    $SupName = @(),

    $TaxesNum = @(XXXX, XXXX),
    $TaxesName = @("Payoffs", "Taxes"),

    $TresNum = @(),
    $TresName = @()
)


$excel = New-Object -ComObject Excel.Application
$excel.visible = $true


$workbook = $excel.Workbooks.Add()
$cells = $workbook.Worksheets.Item(1)

$2 = $workbook.Worksheets.Add()


$counter = 1
$counter2 = 1
$numCounter = 0

function LoopUsers{

    param(

        $UserDept

    )

    foreach($user in $UserDept){
        if($user -eq "S_938g" -or $user -eq "ITCalendar"){
            #do nothing
        }
        else{
            $data = Get-ADUser -Identity $user -Properties telephoneNumber, DisplayName, department, emailAddress | Select-Object -Property telephoneNumber, DisplayName, department, emailAddress

            $cells.Cells.Item($global:counter, 1) = $data.displayname
            $cells.Cells.Item($global:counter, 2) = $data.telephoneNumber
            $cells.Cells.Item($global:counter, 3) = $data.emailAddress
            $global:counter = $global:counter + 1
        }
    }


}


function TitleCell{

    foreach($Dept in $Deptartments){
        
        $global:counter = $global:counter +1

        $cells.Cells.Item($global:counter,1) = $Dept
        $cells.Cells.Item($global:counter,1).Font.Size = 18
        $cells.Cells.Item($global:counter,1).Font.Bold=$True
        $cells.Cells.Item($global:counter,1).Font.Name = "Cambria"
        $cells.Cells.Item($global:counter,1).Font.ThemeFont = 1
        $cells.Cells.Item($global:counter,1).Font.ThemeColor = 4
        $cells.Cells.Item($global:counter,1).Font.ColorIndex = 55
        $cells.Cells.Item($global:counter,1).Font.Color = 8210719

        $global:counter = $global:counter +1

        $cells.Cells.Item($global:counter,1) = $null
        #$global:counter = $global:counter + 1

        $cells.Cells.Item($global:counter,1) = "User"
        $cells.Cells.Item($global:counter,1).Font.Bold=$True

        $cells.Cells.Item($global:counter,2) = "Extension"
        $cells.Cells.Item($global:counter,2).Font.Bold=$True

        $cells.Cells.Item($global:counter,3) = "Email Address"
        $cells.Cells.Item($global:counter,3).Font.Bold=$True

        $global:counter = $global:counter + 2

        if($Dept -eq "Information Technology"){
            $userDept = Get-ADUser -LDAPFilter "(name=*)" -SearchBase "OU=$Dept,OU=staff,DC=corp,DC=jbnutter,DC=com" -SearchScope OneLevel | Select-Object -ExpandProperty SamAccountName
        }
        else{
            $UserDept = Get-ADUser -LDAPFilter "(name=*)" -SearchBase "OU=$Dept,OU=Dept,OU=staff,DC=corp,DC=jbnutter,DC=com" -SearchScope OneLevel | Select-Object -ExpandProperty SamAccountName
        }
        LoopUsers -UserDept $UserDept

    }

}


function loopSharedLines{

    param(

        $Exts,
        $ExtName

    )

    $i = 0
    $len = $exts.length
    for($i = 0; $i -lt $len; $i++){

        $2.Cells.Item($global:counter2, 1) = $ExtName[$i]
        $2.Cells.Item($global:counter2, 2) = $Exts[$i]
        $global:counter2 = $global:counter2 + 1

    }

}

function pullNum{

    param(

        $Dept
    
    )

    if($Dept -eq "Accounting"){
        $global:Exts = $acctNum
        $global:ExtName = $acctName
    }
    elseif($Dept -eq "Cashiers"){
        $global:Exts = $cashNum
        $global:ExtName = $cashName
    }
    elseif($Dept -eq "Compliance"){
        $global:Exts = $compNum
        $global:ExtName = $compName
    }
    elseif($Dept -eq "Customer Service"){
        $global:Exts = $CSNum
        $global:ExtName = $CSName
    }
    elseif($Dept -eq "Default Depts"){
        $global:Exts = $DefaultNum
        $global:ExtName = $DefaultName
    }
    elseif($Dept -eq "Exec Admin"){
        $global:Exts = $ExecAdminNum
        $global:ExtName = $ExecAdminName
    }
    elseif($Dept -eq "Exec Team"){
        $global:Exts = $ExecTeamNum
        $global:ExtName = $ExecTeamName
    }
    elseif($Dept -eq "Facilities"){
        $global:Exts = $FacilNum
        $global:ExtName = $FacilName
    }
    elseif($Dept -eq "HECM"){
        $global:Exts = $HECMNum
        $global:ExtName = $HECMName
    }
    elseif($Dept -eq "HR"){
        $global:Exts = $HRNum
        $global:ExtName = $HRName
    }
    elseif($Dept -eq "Insurance"){
        $global:Exts = $InsNum
        $global:ExtName = $InsName
    }
    elseif($Dept -eq "Investor Reporting"){
        $global:Exts = $InvestNum
        $global:ExtName = $InvestName
    }
    elseif($Dept -eq "Information Technology"){
        $global:Exts = $ITNum
        $global:ExtName = $ITName
    }
    elseif($Dept -eq "Origination"){
        $global:Exts = $OrigNum
        $global:ExtName = $OrigName
    }
    elseif($Dept -eq "Shipping"){
        $global:Exts = $ShipNum
        $global:ExtName = $ShipName
    }
    elseif($Dept -eq "Supervisors"){
        $global:Exts = $SupNum
        $global:ExtName = $SupName
    }
    elseif($Dept -eq "Taxes_Payoffs"){
        $global:Exts = $TaxesNum
        $global:ExtName = $TaxesName
    }
    else{
        $global:Exts = $null
        $global:ExtName = $null
    }

    loopSharedLines -Exts $global:Exts -ExtName $global:ExtName

}

function Sheet2{

    foreach($Dept in $Deptartments){
        
        $global:counter2 = $global:counter2 +1

        $2.Cells.Item($global:counter2,1) = $Dept
        $2.Cells.Item($global:counter2,1).Font.Size = 18
        $2.Cells.Item($global:counter2,1).Font.Bold=$True
        $2.Cells.Item($global:counter2,1).Font.Name = "Cambria"
        $2.Cells.Item($global:counter2,1).Font.ThemeFont = 1
        $2.Cells.Item($global:counter2,1).Font.ThemeColor = 4
        $2.Cells.Item($global:counter2,1).Font.ColorIndex = 55
        $2.Cells.Item($global:counter2,1).Font.Color = 8210719

        $global:counter2 = $global:counter2 +1

        $2.Cells.Item($global:counter2,1) = $null
        #$global:counter = $global:counter + 1

        $2.Cells.Item($global:counter2,1) = "Shared Line"
        $2.Cells.Item($global:counter2,1).Font.Bold=$True

        $2.Cells.Item($global:counter2,2) = "Extension"
        $2.Cells.Item($global:counter2,2).Font.Bold=$True

        $global:counter2 = $global:counter2 + 2

        pullNum -Dept $Dept

    }


}


TitleCell

$cellformat = $cells.UsedRange
$cellformat.EntireColumn.AutoFit()

Sheet2

$2format = $2.UsedRange
$2format.EntireColumn.AutoFit()



# format current date to mm-dd-yyyy
$path = "C:\temp"

# format file name to current directory with name of server-disk-utilization_mm-dd-yyyy.xlsx
 $filename = $path + "\JBNutter_Extensions.xlsx"

# if file exists remove it
if(Test-Path $filename){Remove-Item $filename -force}


$workbook.SaveAs($filename)
# $workbook.Close()
# $excel.Quit()

# cleanly kill processes
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
Remove-Variable excel
# This script parses through folders of excel spreadsheets and generates the total hours of leave used by the entire organization seperated by leave type.

# LEAVE VARIABLES
# =============================
# =============================
$employee_total_annual              = ""
$employee_total_sick                = ""
$employee_total_holiday             = ""
$employee_total_admin               = ""
$employee_total_comp                = ""
$employee_total_new_comp            = ""
$employee_total_leave_without_pay   = ""
$employee_total_personal            = ""
$employee_total_military            = ""


# LEAVE RECORD FOLDERS
# =============================
# =============================
$leave_record_source_folder        = "C:\Leave Records\2016"


# LEAVE RECORD TYPE PREP FOLDERS
# ==============================
$prep_folders = New-Object string[] 6
$prep_folders[0] = "Corr"
$prep_folders[1] = "Crit"
$prep_folders[2] = "FTRH"
$prep_folders[3] = "Non Crit"
$prep_folders[4] = "Op"
$prep_folders[5] = "Sworn"


# EXCEL OBJECT
# =============================
# =============================
$excel = New-Object -ComObject excel.application
$excel.Visible = $false
$excel.DisplayAlerts = $false



# Loop over all of the Excel Spreadsheets in the source folder
# ============================================================
Get-ChildItem $leave_record_source_folder -File -Recurse -Filter *.xls | 

ForEach-Object {

    if ($prep_folders -contains $_.Directory.Name) {

    $employee_total_annual              = ""
    $employee_total_sick                = ""
    $employee_total_holiday             = ""
    $employee_total_admin               = ""
    $employee_total_comp                = ""
    $employee_total_new_comp            = ""
    $employee_total_leave_without_pay   = ""
    $employee_total_personal            = ""
    $employee_total_military            = ""

    # OPEN EMPLOYEE LEAVE RECORD AND MONTH TABS
    # =========================================
    $wb = $excel.Workbooks.Open($_.FullName)
    $ws = $wb.Sheets.Item(1)
    $jan = $wb.Sheets.Item(5)
    $feb = $wb.Sheets.Item(6)
    $mar = $wb.Sheets.Item(7)
    $apr = $wb.Sheets.Item(8)
    $may = $wb.Sheets.Item(9)
    $jun = $wb.Sheets.Item(10)
    $jul = $wb.Sheets.Item(11)
    $aug = $wb.Sheets.Item(12)
    $sep = $wb.Sheets.Item(13)
    $oct = $wb.Sheets.Item(14)
    $nov = $wb.Sheets.Item(15)
    $dec = $wb.Sheets.Item(16)
    
    # Record Employee Attributes
    $employee = $ws.Cells.Item(5,3).text


    # TOTALS
    # =========================================
    $employee_total_annual =  ($jan.Cells.Item(6,14).text/1) `
                            + ($feb.Cells.Item(6,14).text/1) `
                            + ($mar.Cells.Item(6,14).text/1) `
                            + ($apr.Cells.Item(6,14).text/1) `
                            + ($may.Cells.Item(6,14).text/1) `
                            + ($jun.Cells.Item(6,14).text/1) `
                            + ($jul.Cells.Item(6,14).text/1) `
                            + ($aug.Cells.Item(6,14).text/1) `
                            + ($sep.Cells.Item(6,14).text/1) `
                            + ($oct.Cells.Item(6,14).text/1) `
                            + ($nov.Cells.Item(6,14).text/1) `
                            + ($dec.Cells.Item(6,14).text/1)

    $employee_total_sick =    ($jan.Cells.Item(7,14).text/1) `
                            + ($feb.Cells.Item(7,14).text/1) `
                            + ($mar.Cells.Item(7,14).text/1) `
                            + ($apr.Cells.Item(7,14).text/1) `
                            + ($may.Cells.Item(7,14).text/1) `
                            + ($jun.Cells.Item(7,14).text/1) `
                            + ($jul.Cells.Item(7,14).text/1) `
                            + ($aug.Cells.Item(7,14).text/1) `
                            + ($sep.Cells.Item(7,14).text/1) `
                            + ($oct.Cells.Item(7,14).text/1) `
                            + ($nov.Cells.Item(7,14).text/1) `
                            + ($dec.Cells.Item(7,14).text/1)

    $employee_total_holiday =     ($jan.Cells.Item(8,14).text/1) `
                                + ($feb.Cells.Item(8,14).text/1) `
                                + ($mar.Cells.Item(8,14).text/1) `
                                + ($apr.Cells.Item(8,14).text/1) `
                                + ($may.Cells.Item(8,14).text/1) `
                                + ($jun.Cells.Item(8,14).text/1) `
                                + ($jul.Cells.Item(8,14).text/1) `
                                + ($aug.Cells.Item(8,14).text/1) `
                                + ($sep.Cells.Item(8,14).text/1) `
                                + ($oct.Cells.Item(8,14).text/1) `
                                + ($nov.Cells.Item(8,14).text/1) `
                                + ($dec.Cells.Item(8,14).text/1)

    $employee_total_admin =   ($jan.Cells.Item(9,14).text/1) `
                            + ($feb.Cells.Item(9,14).text/1) `
                            + ($mar.Cells.Item(9,14).text/1) `
                            + ($apr.Cells.Item(9,14).text/1) `
                            + ($may.Cells.Item(9,14).text/1) `
                            + ($jun.Cells.Item(9,14).text/1) `
                            + ($jul.Cells.Item(9,14).text/1) `
                            + ($aug.Cells.Item(9,14).text/1) `
                            + ($sep.Cells.Item(9,14).text/1) `
                            + ($oct.Cells.Item(9,14).text/1) `
                            + ($nov.Cells.Item(9,14).text/1) `
                            + ($dec.Cells.Item(9,14).text/1)

    $employee_total_comp =    ($jan.Cells.Item(10,14).text/1) `
                            + ($feb.Cells.Item(10,14).text/1) `
                            + ($mar.Cells.Item(10,14).text/1) `
                            + ($apr.Cells.Item(10,14).text/1) `
                            + ($may.Cells.Item(10,14).text/1) `
                            + ($jun.Cells.Item(10,14).text/1) `
                            + ($jul.Cells.Item(10,14).text/1) `
                            + ($aug.Cells.Item(10,14).text/1) `
                            + ($sep.Cells.Item(10,14).text/1) `
                            + ($oct.Cells.Item(10,14).text/1) `
                            + ($nov.Cells.Item(10,14).text/1) `
                            + ($dec.Cells.Item(10,14).text/1)

    if ((($_.Directory.Name -eq "Sworn") -eq $TRUE) -or (($_.Directory.Name -eq "Corr") -eq $TRUE) -or (($_.Directory.Name -eq "Crit") -eq $TRUE)) {


        $employee_total_new_comp =    ($jan.Cells.Item(11,14).text/1) `
                                    + ($feb.Cells.Item(11,14).text/1) `
                                    + ($mar.Cells.Item(11,14).text/1) `
                                    + ($apr.Cells.Item(11,14).text/1) `
                                    + ($may.Cells.Item(11,14).text/1) `
                                    + ($jun.Cells.Item(11,14).text/1) `
                                    + ($jul.Cells.Item(11,14).text/1) `
                                    + ($aug.Cells.Item(11,14).text/1) `
                                    + ($sep.Cells.Item(11,14).text/1) `
                                    + ($oct.Cells.Item(11,14).text/1) `
                                    + ($nov.Cells.Item(11,14).text/1) `
                                    + ($dec.Cells.Item(11,14).text/1)


        $employee_total_leave_without_pay =   ($jan.Cells.Item(12,14).text/1) `
                                            + ($feb.Cells.Item(12,14).text/1) `
                                            + ($mar.Cells.Item(12,14).text/1) `
                                            + ($apr.Cells.Item(12,14).text/1) `
                                            + ($may.Cells.Item(12,14).text/1) `
                                            + ($jun.Cells.Item(12,14).text/1) `
                                            + ($jul.Cells.Item(12,14).text/1) `
                                            + ($aug.Cells.Item(12,14).text/1) `
                                            + ($sep.Cells.Item(12,14).text/1) `
                                            + ($oct.Cells.Item(12,14).text/1) `
                                            + ($nov.Cells.Item(12,14).text/1) `
                                            + ($dec.Cells.Item(12,14).text/1)

        $employee_total_personal =    ($jan.Cells.Item(13,14).text/1) `
                                    + ($feb.Cells.Item(13,14).text/1) `
                                    + ($mar.Cells.Item(13,14).text/1) `
                                    + ($apr.Cells.Item(13,14).text/1) `
                                    + ($may.Cells.Item(13,14).text/1) `
                                    + ($jun.Cells.Item(13,14).text/1) `
                                    + ($jul.Cells.Item(13,14).text/1) `
                                    + ($aug.Cells.Item(13,14).text/1) `
                                    + ($sep.Cells.Item(13,14).text/1) `
                                    + ($oct.Cells.Item(13,14).text/1) `
                                    + ($nov.Cells.Item(13,14).text/1) `
                                    + ($dec.Cells.Item(13,14).text/1)

        $employee_total_military =    ($jan.Cells.Item(14,14).text/1) `
                                    + ($feb.Cells.Item(14,14).text/1) `
                                    + ($mar.Cells.Item(14,14).text/1) `
                                    + ($apr.Cells.Item(14,14).text/1) `
                                    + ($may.Cells.Item(14,14).text/1) `
                                    + ($jun.Cells.Item(14,14).text/1) `
                                    + ($jul.Cells.Item(14,14).text/1) `
                                    + ($aug.Cells.Item(14,14).text/1) `
                                    + ($sep.Cells.Item(14,14).text/1) `
                                    + ($oct.Cells.Item(14,14).text/1) `
                                    + ($nov.Cells.Item(14,14).text/1) `
                                    + ($dec.Cells.Item(14,14).text/1)


    } else {
        


        $employee_total_leave_without_pay =   ($jan.Cells.Item(11,14).text/1) `
                                            + ($feb.Cells.Item(11,14).text/1) `
                                            + ($mar.Cells.Item(11,14).text/1) `
                                            + ($apr.Cells.Item(11,14).text/1) `
                                            + ($may.Cells.Item(11,14).text/1) `
                                            + ($jun.Cells.Item(11,14).text/1) `
                                            + ($jul.Cells.Item(11,14).text/1) `
                                            + ($aug.Cells.Item(11,14).text/1) `
                                            + ($sep.Cells.Item(11,14).text/1) `
                                            + ($oct.Cells.Item(11,14).text/1) `
                                            + ($nov.Cells.Item(11,14).text/1) `
                                            + ($dec.Cells.Item(11,14).text/1)


        $employee_total_military =    ($jan.Cells.Item(12,14).text/1) `
                                    + ($feb.Cells.Item(12,14).text/1) `
                                    + ($mar.Cells.Item(12,14).text/1) `
                                    + ($apr.Cells.Item(12,14).text/1) `
                                    + ($may.Cells.Item(12,14).text/1) `
                                    + ($jun.Cells.Item(12,14).text/1) `
                                    + ($jul.Cells.Item(12,14).text/1) `
                                    + ($aug.Cells.Item(12,14).text/1) `
                                    + ($sep.Cells.Item(12,14).text/1) `
                                    + ($oct.Cells.Item(12,14).text/1) `
                                    + ($nov.Cells.Item(12,14).text/1) `
                                    + ($dec.Cells.Item(12,14).text/1)

    }


    $excel.Workbooks.Close()


    $total_leave_annual             = ($total_leave_annual/1) + ($employee_total_annual/1)
    $total_leave_sick               = ($total_leave_sick/1) + ($employee_total_sick/1)
    $total_leave_holiday            = ($total_leave_holiday/1) + ($employee_total_holiday/1)
    $total_leave_admin              = ($total_leave_admin/1) + ($employee_total_admin/1)
    $total_leave_comp               = ($total_leave_comp/1) + ($employee_total_comp/1)
    $total_leave_new_comp           = ($total_leave_new_comp/1) + ($employee_total_new_comp/1)
    $total_leave_leave_without_pay  = ($total_leave_leave_without_pay/1) + ($employee_total_leave_without_pay/1)
    $total_leave_military           = ($total_leave_military/1) + ($employee_total_military/1)
    $total_leave_personal           = ($total_leave_personal/1) + ($employee_total_personal/1)

    Write-Host ""
    Write-Host "Annual = $total_leave_annual"
    Write-Host "Sick = $total_leave_sick"
    Write-Host "Holiday = $total_leave_holiday"
    Write-Host "Administrative = $total_leave_admin"
    Write-Host "Comp = $total_leave_comp"
    Write-Host "New Comp = $total_leave_new_comp"
    Write-Host "Leave Without Pay = $total_leave_leave_without_pay"
    Write-Host "Military = $total_leave_military"
    Write-Host "Personal = $total_leave_personal"
    Write-Host ""

    # CLOSE
    # =====
    $excel.Workbooks.Close()
    }

}

$excel.Quit()
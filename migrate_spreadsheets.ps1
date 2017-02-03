# EMPLOYEE VARIABLES
# =============================
# =============================
$year = "2017"
$employee = ""
$date_hired = ""
$carry_over_leave_annual   = ""
$carry_over_leave_sick     = ""
$carry_over_leave_comp     = ""
$carry_over_leave_holiday  = ""
$carry_over_leave_personal = ""


# LEAVE RECORD TEMPLATES
# =============================
# =============================
$template_location      = "C:\Leave Records\Templates\"
#$template_location      = "\\file_server\group_shared\HR\Excel Leave Records\Leave Record Templates\"
$template_corrections   = "PROPOSED 2017 CORRECTIONS MASTER.xls"
$template_critical      = "PROPOSED 2017 CRITICAL MASTER.xls"
$template_ftrh          = "PROPOSED 2017 FT-REDUCED HOURS MASTER.xls"
$template_non_critical  = "PROPOSED 2017 NON-CRITICAL MASTER.xls"
$template_operational   = "PROPOSED 2017 OPERATIONAL MASTER.xls"
$template_sworn         = "PROPOSED 2017 SWORN MASTER.xls"

# Write-Host ""
# Write-Host "Leave Record Template Location: $template_location"

# LEAVE RECORD FOLDERS
# =============================
# =============================
$leave_record_source_folder        = "C:\Leave Records\2016"
$leave_record_destination_folder   = "C:\Leave Records\2017"
#$leave_record_source_folder   = "\\file_server\group_shared\HR\Excel Leave Records\2016"
#$leave_record_destination_folder   = "\\file_server\group_shared\HR\Excel Leave Records\2017"

# Write-Host "Leave Records Location: $leave_record_source_folder"
# Write-Host "Leave Records Destination: $leave_record_destination_folder"
# Write-Host ""


# LEAVE RECORD TYPE PREP FOLDERS
# ==============================
$prep_folders = New-Object string[] 6
$prep_folders[0] = "Corr"
$prep_folders[1] = "Crit"
$prep_folders[2] = "FTRH"
$prep_folders[3] = "Non Crit"
$prep_folders[4] = "Op"
$prep_folders[5] = "Sworn"


# DUPLICATE LEAVE RECORD FOLDER STRUCTURE
# ========================================================================
# ========================================================================

# Remove existing destination folder
Remove-Item $leave_record_destination_folder -Recurse

# Create new destination folder
New-Item -ItemType Directory -Path $leave_record_destination_folder

# Copy source folder subfolders to destination folder
Get-ChildItem $leave_record_source_folder -Attributes D -Recurse | ForEach-Object {

    # Copy Level 1 Subfolders
    if($_.Parent.FullName -eq $leave_record_source_folder) {
        # Copy non-prep folders
        if ($prep_folders -cnotcontains $_.Name) {
            New-Item -ItemType Directory -Path ($leave_record_destination_folder + "\$_")
        }
    }

    # Copy Level 2 Subfolders
    if($_.Parent.Parent.FullName -eq $leave_record_source_folder) {
        # Copy non-prep folders
        if ($prep_folders -cnotcontains $_.Name) {
            New-Item -ItemType Directory -Path ($leave_record_destination_folder + "\" + $_.Parent.Name + "\$_")
        }
    }

    # Copy Level 3 Subfolders
    if($_.Parent.Parent.Parent.FullName -eq $leave_record_source_folder) {
        # Copy non-prep folders
        if ($prep_folders -cnotcontains $_.Name) {
            New-Item -ItemType Directory -Path ($leave_record_destination_folder + "\" + $_.Parent.Parent.Name + "\" + $_.Parent.Name + "\$_")
        }
    }
}

# ========================================================================
# ========================================================================



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

        # Zero Out Variables
        # ==================
        $carry_over_leave_annual   = ""
        $carry_over_leave_sick     = ""
        $carry_over_leave_comp     = ""
        $carry_over_leave_holiday  = ""
        $carry_over_leave_personal = ""

        # Calculate the destination based on the location
        # ===============================================
        $location = $_.FullName.Replace($_.Directory.FullName,$_.Directory.Parent.FullName)
        $destination = $location.Replace($leave_record_source_folder,$leave_record_destination_folder)

        # Choose the appropriate Leave Record Template
        # ============================================
        switch ($_.Directory.Name) {

            "Corr"      { $leave_template = $template_location + $template_corrections }
            "Crit"      { $leave_template = $template_location + $template_critical }
            "FTRH"      { $leave_template = $template_location + $template_ftrh }
            "Non Crit"  { $leave_template = $template_location + $template_non_critical }
            "Op"        { $leave_template = $template_location + $template_operational }
            "Sworn"     { $leave_template = $template_location + $template_sworn }
            Default     { $leave_template = "" }
        
        }

        # We have the location of the file
        # We have the destination of the file
        # We know what template to use
        # Let's begin
        # ====================================

        # OPEN EMPLOYEE LEAVE RECORD
        # ===============================    
        $wb = $excel.Workbooks.Open($_.FullName)
        $ws = $wb.Sheets.Item(1)
        $wt = $wb.Sheets.Item(3)

        # Record Employee Attributes
        $employee                   = $ws.Cells.Item(5,3).text
        $date_hired                 = $ws.Cells.Item(6,3).text
        $carry_over_leave_annual    = $ws.Cells.Item(17,3).text
        $carry_over_leave_sick      = $ws.Cells.Item(18,3).text
        $carry_over_leave_comp      = $ws.Cells.Item(19,3).text
        $carry_over_leave_holiday   = $ws.Cells.Item(20,3).text
        
        if ((($_.Directory.Name -eq "Sworn") -eq $TRUE) -or (($_.Directory.Name -eq "Corr") -eq $TRUE)) {
            $carry_over_leave_personal = $wt.Cells.Item(15,19).text
        }

        $excel.Workbooks.Close()

        Write-Host ""
        Write-Host "Employee: $employee"
        Write-Host "Date Hired: $date_hired" 
        Write-Host "Annual: $carry_over_leave_annual"
        Write-Host "Sick: $carry_over_leave_sick"
        Write-Host "Comp: $carry_over_leave_comp"
        Write-Host "Holiday: $carry_over_leave_holiday"
        Write-Host "Personal: $carry_over_leave_personal"
        Write-Host ""

        # OPEN 2017 LEAVE RECORD TEMPLATE
        # ===============================
        $wb = $excel.Workbooks.Open($leave_template)
        $ws = $wb.Sheets.Item(1)

        # UPDATE CELLS
        # ============
        $ws.Cells.Item(4,3) = $year
        $ws.Cells.Item(5,3) = $employee
        $ws.Cells.Item(6,3) = $date_hired
        $ws.Cells.Item(10,3) = $carry_over_leave_annual
        $ws.Cells.Item(11,3) = $carry_over_leave_sick
        $ws.Cells.Item(12,3) = $carry_over_leave_comp
        $ws.Cells.Item(13,3) = $carry_over_leave_holiday

        if ((($_.Directory.Name -eq "Sworn") -eq $TRUE) -or (($_.Directory.Name -eq "Corr") -eq $TRUE)) {
            $ws.Cells.Item(14,3) = $carry_over_leave_personal
        }

        # SAVE SPREADSHEET
        # ================
        $ws = $wb.Sheets.Item(4)
        $ws.activate()
        $excel.ActiveWorkbook.SaveAs($destination)        

        # CLOSE
        # =====
        $excel.Workbooks.Close()

    }

}

$excel.Quit()
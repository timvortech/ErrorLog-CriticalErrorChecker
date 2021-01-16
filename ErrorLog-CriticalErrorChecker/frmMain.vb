Imports Microsoft.Office.Interop
Imports System.Net.Mail
Imports System.IO

Public Module GlobalFilenameVariables

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'LIVE FILES'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public UserLogged As String = "C:\Users\" & Environment.UserName
    Public strPath As String = UserLogged & "\Lehan Drugs\HME Tactical Team - Error Form Follow Up and Reports"
    Public backupLoc As String = UserLogged & "\Lehan Drugs\HME Tactical Team - Error Form Follow Up and Reports\ErrorReportAutomation\Backups\ErrorLog_CriticalErrorChecker_BackupCopies"
    Public Datestamp As String = Replace(Now.ToString("yyyy-MM-dd HH_mm_ss"), "/", "-")
    Public completedIDsTextFileLoc As String = strPath & "\ErrorReportAutomation\CompletedErrorIDs.txt"
    Public reportTrackingTextFileLoc As String = strPath & "\ErrorReportAutomation\ReportTracking.txt"
    'Public EXLreportsLoc As String = UserLogged & "\Lehan Drugs\HME Tactical Team - Error Form Follow Up and Reports\ErrorReportAutomation\Testing\ReportTextLoc3\"
    Public EXLreportsLoc As String = UserLogged & "\Lehan Drugs\HME Tactical Team - Error Form Follow Up and Reports\Error Coaching Reports\Excel Reports\"
    Public Path_Datestamp As String = (backupLoc & "\" & Datestamp)
    Public excel_ERR_workbook_filename As String = Path_Datestamp & "\Oops I Did It again (2.0).xlsx"
    Public excel_TMP_workbook_filename As String = Path_Datestamp & "\ErrorReportFormTemplate_CriticalError.xlsm"
    '

End Module
Public Class frmMain
    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load








        '

        'SendErrorEmail("Test Error Email")'TESTING

        'Creat Backup Copy to process off of
        'On Error Resume Next

        My.Computer.FileSystem.CreateDirectory(Path_Datestamp)
        'My.Computer.FileSystem.CreateDirectory(Path_Datestamp.Replace("Excel Reports", "Published PDF Reports"))
        'My.Computer.FileSystem.CreateDirectory(EXLreportsLoc)
        My.Computer.FileSystem.CopyFile(strPath & "\Oops I Did It again (2.0).xlsx", Path_Datestamp & "\Oops I Did It again (2.0).xlsx")
        My.Computer.FileSystem.CopyFile(strPath & "\ErrorReportAutomation\ErrorReportFormTemplate_CriticalError.xlsm", Path_Datestamp & "\ErrorReportFormTemplate_CriticalError.xlsm")
        'My.Computer.FileSystem.CopyDirectory(strPath, Path_Datestamp, False)


        ValidateFile(excel_ERR_workbook_filename)       ' Validates error excel file
        ValidateFile(excel_TMP_workbook_filename)       ' Validates error reporting template excel file
        OpenExcelWorkbook_ErrorLog(excel_ERR_workbook_filename)  ' Opens the excel workbook that has the error log 
        'OpenExcelWorkbook_ErrorLog(excel_TMP_workbook_filename)  ' Opens the excel workbook that is the excel reporting template

        Me.Close()
        Me.Dispose()
    End Sub


    ' Opens the excel workbook and exports certain parts to a pdf
    Private Sub OpenExcelWorkbook_ErrorLog(ByVal xl_wrkbk_filename As String)
        ' DECLARE THE VARIABLES
        Dim xl_error_log_wrksheet_name As String      ' The new referrals worksheet's name
        Dim xl_support_wrksheet_name As String            ' The support worksheet's name
        Dim xl_reportfile_wrksheet_name As String
        Dim xl_app As Excel.Application                   ' The excel application object
        Dim xl_wrkbk As Excel._Workbook                   ' The excel workbook object
        Dim xl_error_log_wrksheet As Excel._Worksheet ' The excel new_referrals worksheet object
        Dim xl_support_wrksheet As Excel._Worksheet       ' The excel support worksheet object
        Dim xl_reportfile_wrksheet As Excel._Worksheet
        Dim output_exl_folder As String                   ' The output folder where the pdf will be exported to
        Dim names(0) As String                            ' Contains the physican's names from the support worksheet
        Dim group_names(0) As String                      ' Contains the group names from the support worksheet, with which the physicians will be grouped together
        Dim row_count As Integer = 1                      ' Contains the row count



        'output_exl_folder = "C:\Users\denni\Vortech\Coding - LehanDrugs\Reports\"
        'output_exl_folder = "C:\Users\Administrator\Lehan Drugs\Lehan's HME - Order Processing - Order Processing\Reports\"
        output_exl_folder = EXLreportsLoc
        'output_exl_folder = "C:\Users\BTservice\Lehan Drugs\Lehan's HME - Order Processing - Order Processing\Reports\"

        'output_exl_folder = CreateNewDir(output_exl_folder) ' Creates a new dir in the folder entered

        ' If (Not System.IO.Directory.Exists(output_exl_folder)) Then
        output_exl_folder = CreateNewDir(output_exl_folder)
            'End If


            ' Opens an excel application
            xl_error_log_wrksheet_name = "ErrorReporter"
        xl_support_wrksheet_name = "CriticalErrorSupport"
        xl_app = New Excel.Application()


        ' Opens the excel workbook
        Try

            xl_wrkbk = xl_app.Workbooks.Open(xl_wrkbk_filename)
        Catch ex As Exception
            xl_app.Quit()       ' Exits the excel application

            ReleaseObject(xl_app)       ' Releases the excel applications from use.

            'Windows.Forms.MessageBox.Show("An error has occured while trying to open the excel workbook. Updates have not been saved.", "Error")
            SendErrorEmail("An error has occured while trying to open the excel workbook. Updates have not been saved.")
            Me.Close()
            Me.Dispose()
            Exit Sub
        End Try



        ' Opens the excel worksheets
        Try
            Threading.Thread.Sleep(1000)
            xl_error_log_wrksheet = xl_wrkbk.Worksheets(xl_error_log_wrksheet_name)
            Threading.Thread.Sleep(1000)
            xl_support_wrksheet = xl_wrkbk.Worksheets(xl_support_wrksheet_name)
            Threading.Thread.Sleep(1000)
        Catch ex As Exception
            xl_wrkbk.Close()    ' Closes the excel workbook.
            xl_app.Quit()       ' Exits the excel application

            ReleaseObject(xl_app)       ' Releases the excel applications from use.
            ReleaseObject(xl_wrkbk)     ' Releases the excel workbook from use.

            'Windows.Forms.MessageBox.Show("An error has occured while trying to open a worksheet. Updates have not been saved.", "Error")
            SendErrorEmail("An error has occured while trying to open a worksheet. Updates have not been saved.")
            Me.Close()
            Me.Dispose()
            Exit Sub
        End Try

        ' Counts the number of non-blank rows in the support worksheet.
        ' Exits the loop once it reaches a blank physician cell
        For i As Integer = 2 To xl_support_wrksheet.UsedRange.Rows.Count()
            If Not xl_support_wrksheet.Cells(i, 1).value.ToString = "" Then
                row_count += 1
            Else
                Exit For
            End If
        Next i

        'For i = 1 To xl_wrkbk.Names.Item("Support_NameList").rows.count
        '    MsgBox xl_wrkbk.Names.Item("Support_NameList").RefersToRange

        'Next

        ' Loops throught the rows and columns in the Support worksheet.
        ' Separates the two columns into two different arrays, one
        ' containing the Physician's names and the other their group names.

        'Read in completed ID text file
        Dim textFileLoc As String = completedIDsTextFileLoc
        Dim completedIDs() As String                   ' Contains the row count
        completedIDs = IO.File.ReadAllLines(textFileLoc)

        xl_app.Calculation = Excel.XlCalculation.xlCalculationManual
        For col_i As Integer = 6 To 6
            For row_i As Integer = 2 To completedIDs.Length() + 1 'xl_support_wrksheet.Rows.Count()
                'write in values
                xl_support_wrksheet.Cells(row_i, col_i).Value = completedIDs(row_i - 2)
            Next row_i
        Next col_i
        xl_app.Calculation = Excel.XlCalculation.xlCalculationAutomatic

        xl_app.Calculate()



        If xl_support_wrksheet.Range("CriticalErrorCounter_Open").Value = 0 Then
            'exit if there is nothing to process
            GoTo SKIP_TEMPLATE

        End If




        For col_i As Integer = 1 To 1
            For row_i As Integer = 2 To row_count 'xl_support_wrksheet.Rows.Count()

                ReDim Preserve names(names.Length())    ' Increases the length by 1
                names(names.Length - 1) = xl_support_wrksheet.Cells(row_i, col_i).Value


            Next row_i
        Next col_i


        ProcessGroupFiles(names, output_exl_folder, xl_wrkbk, xl_error_log_wrksheet, xl_support_wrksheet, xl_app) ' NEED TO ADD WHEN UPGRADED TO THIS

        Try
        Catch ex As Exception
            xl_wrkbk.Close()    ' Closes the excel workbook.
            xl_app.Quit()       ' Exits the excel application

            ReleaseObject(xl_app)       ' Releases the excel applications from use.
            ReleaseObject(xl_wrkbk)     ' Releases the excel workbook from use.
            ReleaseObject(xl_error_log_wrksheet)  ' Releases the excel worksheet from use.
            ReleaseObject(xl_support_wrksheet)        ' Releases the excel worksheet from use.

            'Windows.Forms.MessageBox.Show("An error has occured while creating PDFs to be exported. Updates have not been saved.", "Error")
            SendErrorEmail("An error has occured while creating files. Updates have not been saved.")
            Me.Close()
            Me.Dispose()
            Exit Sub
        End Try


SKIP_TEMPLATE:
        Threading.Thread.Sleep(1000)    ' Fixes a syncing error ["Failed to merge Office file."]

        xl_wrkbk.Close()    ' Closes the excel workbook.
        xl_app.Quit()       ' Exits the excel application

        ReleaseObject(xl_app)       ' Releases the excel applications from use.
        ReleaseObject(xl_wrkbk)     ' Releases the excel workbook from use.
        ReleaseObject(xl_error_log_wrksheet)    ' Releases the excel worksheet from use.
        ReleaseObject(xl_support_wrksheet)          ' Releases the excel worksheet from use.
    End Sub

    Private Sub OpenExcelWorkbook_ReportTemplate(ByVal criticalErrorID As String, ByVal xl_wrkbk_filename As String, ByVal xl_error_log_wrkbook As Excel._Workbook, ByVal xl_error_log_wrksheet As Excel._Worksheet, ByVal newFileName As String, ByVal xl_app As Excel._Application, ByVal reportType As String)
        ' DECLARE THE VARIABLES
        Dim xl_reportfile_wrksheet_name As String
        Dim xl_support_wrksheet_name As String
        'Dim xl_app As Excel.Application                   ' The excel application object
        Dim xl_wrkbk As Excel._Workbook                   ' The excel workbook object
        Dim xl_wrkbk_ErrorLog As Excel._Workbook                   ' The excel workbook object
        Dim xl_reportfile_wrksheet As Excel._Worksheet
        'Dim xl_support_wrksheet As Excel._Worksheet
        Dim output_exl_folder As String                   ' The output folder where the pdf will be exported to
        Dim names(0) As String                            ' Contains the physican's names from the support worksheet
        Dim group_names(0) As String                      ' Contains the group names from the support worksheet, with which the physicians will be grouped together
        Dim row_count As Integer = 1                      ' Contains the row count


        'output_exl_folder = "C:\Users\denni\Vortech\Coding - LehanDrugs\Reports\"
        'output_exl_folder = "C:\Users\Administrator\Lehan Drugs\Lehan's HME - Order Processing - Order Processing\Reports\"
        output_exl_folder = EXLreportsLoc
        'output_exl_folder = "C:\Users\BTservice\Lehan Drugs\Lehan's HME - Order Processing - Order Processing\Reports\"

        'output_exl_folder = CreateNewDir(output_exl_folder) ' Creates a new dir in the folder entered

        ' Opens an excel application
        xl_reportfile_wrksheet_name = "ErrorReport"
        'xl_support_wrksheet_name = "CriticalErrorSupport"

        'xl_app = New Excel.Application()


        ' Opens the excel workbook
        Try

            xl_wrkbk = xl_app.Workbooks.Open(xl_wrkbk_filename)
        Catch ex As Exception
            xl_app.Quit()       ' Exits the excel application

            ReleaseObject(xl_app)       ' Releases the excel applications from use.

            'Windows.Forms.MessageBox.Show("An error has occured while trying to open the excel workbook. Updates have not been saved.", "Error")
            SendErrorEmail("An error has occured while trying to open the excel workbook. Updates have not been saved.")
            Me.Close()
            Me.Dispose()
            Exit Sub
        End Try



        ' Opens the excel worksheets
        Try

            'Threading.Thread.Sleep(1000)
            xl_reportfile_wrksheet = xl_wrkbk.Worksheets(xl_reportfile_wrksheet_name)
            'Threading.Thread.Sleep(1000)
            'xl_support_wrksheet = xl_wrkbk.Worksheets(xl_support_wrksheet_name)
            'Threading.Thread.Sleep(1000)
        Catch ex As Exception
            xl_wrkbk.Close()    ' Closes the excel workbook.
            xl_app.Quit()       ' Exits the excel application

            ReleaseObject(xl_app)       ' Releases the excel applications from use.
            ReleaseObject(xl_wrkbk)     ' Releases the excel workbook from use.

            'Windows.Forms.MessageBox.Show("An error has occured while trying to open a worksheet. Updates have not been saved.", "Error")
            SendErrorEmail("An error has occured while trying to open a worksheet. Updates have not been saved.")
            Me.Close()
            Me.Dispose()
            Exit Sub
        End Try

        'NEW CODE

        Clipboard.Clear()


        xl_error_log_wrkbook.Activate()
        xl_error_log_wrksheet.Range("Source_ReportColumnRange").Copy() 'names.Item("Source_ReportColumnRange").RefersToRange.Copy(xl_error_log_wrksheet.Names.Item("Source_ReportColumnRange").RefersToRange)
        xl_wrkbk.Activate()
        ' xl_reportfile_wrksheet.Select()

        xl_reportfile_wrksheet.Range("Dest_PasteLoc").PasteSpecial(Excel.XlPasteType.xlPasteFormats) 'Names.Item("Dest_PasteLoc").RefersToRange.PasteSpecial(Excel.XlPasteType.xlPasteFormats)

        xl_reportfile_wrksheet.Range("Dest_PasteLoc").PasteSpecial(Excel.XlPasteType.xlPasteValues) '.Names.Item("Dest_PasteLoc").RefersToRange.PasteSpecial(Excel.XlPasteType.xlPasteValues)

        xl_reportfile_wrksheet.Range("HideRangeCheck").Rows.AutoFit()

        'xl_reportfile_wrksheet.Range("TeamFooter_ResizeNameGroup").Rows.RowHeight = 72
        xl_reportfile_wrksheet.Range("HideFullRange").Rows.Hidden = True

        If reportType = "Team" Then
            xl_reportfile_wrksheet.Range("HideFooter_TeamReport").Rows.Hidden = True
            Threading.Thread.Sleep(5000)    ' Fixes a syncing error ["Failed To merge Office file."]
        End If



        ' Loops through rows of the specified named range and hides and blank rows
        'xl_app.Calculation = Excel.XlCalculation.xlCalculationManual


        'For row_i As Integer = xl_reportfile_wrksheet.Range("HideRangeCheck").Row To xl_reportfile_wrksheet.Range("HideRangeCheck").Rows.Count() + xl_reportfile_wrksheet.Range("HideRangeCheck").Row - 1
        '    If xl_reportfile_wrksheet.Cells(row_i, 18).Value = Nothing Then
        '        xl_reportfile_wrksheet.Rows(row_i).Hidden = True
        '        xl_reportfile_wrksheet.Rows(row_i & ":" & )
        '    End If
        'Next row_i
        'xl_app.Calculation = Excel.XlCalculation.xlCalculationAutomatic

        Clipboard.Clear()


        xl_app.DisplayAlerts = False
        Threading.Thread.Sleep(1000)    ' Fixes a syncing error ["Failed To merge Office file."]
        xl_wrkbk.SaveCopyAs(newFileName)



SKIP_TEMPLATE:
        Threading.Thread.Sleep(1000)    ' Fixes a syncing error ["Failed To merge Office file."]

        xl_wrkbk.Close()    ' Closes the excel workbook.
        xl_app.DisplayAlerts = True
        'xl_app.Quit()       ' Exits the excel application

        'ReleaseObject(xl_app)       ' Releases the excel applications from use.
        ReleaseObject(xl_wrkbk)     ' Releases the excel workbook from use.
        ReleaseObject(xl_reportfile_wrksheet)    ' Releases the excel worksheet from use.
    End Sub

    ' Releases the object from use.
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


    ' Validates the filename entered.
    Private Function ValidateFile(ByVal filename As String)
        If File.Exists(filename) = True Then
            Return True
        Else
            Return False
        End If
    End Function


    ' Finds each of the group of doctors with the same group name,
    ' gets the worksheet ready to be exported, and then exports each group to pdfs.
    'Private Sub ProcessIndividualFiles(ByVal docnames() As String, ByVal output_folder As String, ByVal excel_workbook As Excel._Workbook, ByVal excel_error_log_wrksheet As Excel._Worksheet)
    '    Dim cur_group(0) As String          ' Contains all of the docs from the current group
    '    Dim grps_done(0) As String          ' Contains all of the group names which have already been completed
    '    Dim cur_count As Integer = 0        ' Count of indexes in the cur_group array
    '    Dim grps_done_count As Integer = 0  ' Count of indexes in the grps_done array
    '    Dim full_output_filename As String  ' Contains the full filename for the pdf file
    '    Dim columns() As String = {"B", "C", "G", "I", "K", "L", "N", "M", "O", "P", "Q"}   ' Array of columns to be hidden, and then shown after the worksheet is exported to a pdf


    '    'New LOOP added for Error Log Reporting
    '    For i_name_1 As Integer = 0 To docnames.Length() - 1
    '        If docnames(i_name_1) <> "" Then
    '            excel_workbook.Names.Item("Error_RawName").RefersToRange.Value = docnames(i_name_1)





    '        End If
    '    Next i_name_1
    '    Exit Sub



    '    'excel_new_referrals_wrksheet.EnableAutoFilter = True    ' Enables the auto-filter

    '    ' Loops through the group names, matches non-blank group names, and adds those doctor's names to the cur_group array
    '    '''        For i_grp_name_1 As Integer = 0 To grp_names.Length() - 1
    '    '''            If grp_names(i_grp_name_1) <> "" Then   ' The current doctor has a group

    '    '''                ' Checks if the current group has already been completed.
    '    '''                ' If so, it skips to the next group.
    '    '''                For i As Integer = 0 To grps_done.Length() - 1
    '    '''                    If grp_names(i_grp_name_1) = grps_done(i) Then
    '    '''                        GoTo skip1
    '    '''                    End If
    '    '''                Next i

    '    '''                grps_done(grps_done_count) = grp_names(i_grp_name_1) ' Adds the current doctor's group to the grps_done array; this group will not be used again
    '    '''                ReDim Preserve grps_done(grps_done.Length())         ' Increases the grps_done array size by one
    '    '''                grps_done_count += 1

    '    '''                ' Loop through the same group names array to find other doctors matching the current doc
    '    '''                For i_grp_name_2 As Integer = 0 To grp_names.Length() - 1
    '    '''                    If grp_names(i_grp_name_1) = grp_names(i_grp_name_2) Then
    '    '''                        cur_group(cur_count) = doc_names(i_grp_name_2)       ' Adds the doctor matching the current doc (from the outside loop), to the cur_group array containing all of the matching docs
    '    '''                        ReDim Preserve cur_group(cur_group.Length())         ' Increases the cur_group size by one
    '    '''                        cur_count += 1
    '    '''                    End If
    '    '''                Next i_grp_name_2

    '    '''                ReDim Preserve cur_group(cur_count - 1)    ' Removes the empty element in the last index of the array

    '    '''            Else    ' The current doctor does not have a group
    '    '''                cur_group(0) = doc_names(i_grp_name_1)
    '    '''            End If

    '    '''            'excel_new_referrals_wrksheet.EnableAutoFilter = True    ' Enables the auto-filter

    '    '''            '' Gets the worksheet ready to be exported to a PDF
    '    '''            'GetWorksheetReadyForExport(cur_group, excel_new_referrals_wrksheet, columns)

    '    '''            ' Gets the full filename for the new pdf file yet-to-be-created
    '    '''            full_output_filename = CreateFilename(output_folder, cur_group, grp_names(i_grp_name_1))


    '    '''            'Exports the worksheet to a PDF (TF: will also now fax as needed by the program)
    '    '''            ExportToEXL(full_output_filename, excel_workbook, excel_new_referrals_wrksheet, cur_group, grp_names(i_grp_name_1))


    '    '''            full_output_filename = output_folder    ' Resets the output filename

    '    '''skip1:
    '    '''            ReDim cur_group(0)  ' Resets the array containing the current group of doctors
    '    '''            cur_group(0) = ""   ' Cleans the first element
    '    '''            cur_count = 0       ' Resets the count
    '    '''        Next i_grp_name_1

    '    'excel_new_referrals_wrksheet.AutoFilter.ShowAllData()    ' Clears all the filters
    '    'ShowColumns(excel_new_referrals_wrksheet, columns)       ' Shows all hidden columns
    '    'excel_new_referrals_wrksheet.EnableAutoFilter = False    ' Disables the auto-filter
    '    'ApplyDefaultFilters(excel_workbook, excel_new_referrals_wrksheet)   ' Applies the default filters which were just cleared - 8/8/2020 TF commented out, due to error thrown

    'End Sub


    Private Sub ProcessGroupFiles(ByVal emp_names() As String, ByVal output_folder As String, ByVal excel_workbook As Excel._Workbook, ByVal excel_error_log_wrksheet As Excel._Worksheet, ByVal excel_support_wrksheet As Excel._Worksheet, ByVal xl_app As Excel._Application)
        Dim cur_group(0) As String          ' Contains all of the docs from the current group
        Dim grps_done(0) As String          ' Contains all of the group names which have already been completed
        Dim cur_count As Integer = 0        ' Count of indexes in the cur_group array
        Dim grps_done_count As Integer = 0  ' Count of indexes in the grps_done array
        Dim full_output_filename As String  ' Contains the full filename for the pdf file
        Dim columns() As String = {"B", "C", "G", "I", "K", "L", "N", "M", "O", "P", "Q"}   ' Array of columns to be hidden, and then shown after the worksheet is exported to a pdf
        Dim leadDirName As String
        Dim selected_name As String
        Dim selected_id As String
        Dim finalFileName As String
        Dim reportType_Val As String



        For i_emp_name_1 As Integer = 0 To emp_names.Length() - 1

            If emp_names(i_emp_name_1) <> "" Then
                selected_id = emp_names(i_emp_name_1).Split("_"c)(0)
                selected_name = emp_names(i_emp_name_1).Split("_"c)(1)

                'leadDirName = output_folder
                'finalFileName = output_folder & excel_error_log_wrksheet.Range("CriticalErrorLeadDirectory_ErrorReport").Value & "\Critical_Error_Reports\"
                finalFileName = output_folder & "Critical_Error_Reports\"

                If (Not System.IO.Directory.Exists(finalFileName)) Then
                    System.IO.Directory.CreateDirectory(finalFileName)
                End If

                If (Not System.IO.Directory.Exists(finalFileName.Replace("Excel Reports", "Published PDF Reports"))) Then
                    System.IO.Directory.CreateDirectory(finalFileName.Replace("Excel Reports", "Published PDF Reports"))
                End If


                If (Not System.IO.Directory.Exists(finalFileName.Replace("Excel Reports", "Published PDF Reports") & "UnCrt\")) Then
                    Directory.CreateDirectory(finalFileName.Replace("Excel Reports", "Published PDF Reports") & "UnCrt\")    ' Creates a new directory with the current date as the folder name
                End If


                'Create Individual File
                reportType_Val = "CriticalError"
                excel_error_log_wrksheet.Range("SeatGroup_ErrorReport").Value = ""
                excel_error_log_wrksheet.Range("CriticalErrorLookupVal_ErrorReport").Value = emp_names(i_emp_name_1)
                excel_error_log_wrksheet.Range("Error_ReportType").Value = reportType_Val
                xl_app.Calculate()
                finalFileName = finalFileName & excel_error_log_wrksheet.Range("Error_FileNameAddition").Value & " - "
                finalFileName = finalFileName & excel_error_log_wrksheet.Range("Error_Name").Value & " - Critical Error Report - "
                finalFileName = finalFileName & Format(Now(), "hh_mm_ss_tt") & ".xlsm"
                'finalFileName = finalFileName & excel_error_log_wrksheet.Range("Error_Name").Value
                'finalFileName = finalFileName & excel_error_log_wrksheet.Range("Error_FileNameAddition").Value & ".xlsm"
                OpenExcelWorkbook_ReportTemplate(selected_id, excel_TMP_workbook_filename, excel_workbook, excel_error_log_wrksheet, finalFileName, xl_app, reportType_Val)

                'Dim inputString As String = "10000"
                'My.Computer.FileSystem.WriteAllText(completedIDsTextFileLoc, selected_id, True)

                Using writer As New StreamWriter(completedIDsTextFileLoc, True)
                    writer.WriteLine(selected_id)
                End Using

                Using writer As New StreamWriter(reportTrackingTextFileLoc, True)
                    writer.WriteLine(Now() & vbTab & "CriticalError" & vbTab & selected_id & vbTab & "1" & vbTab & "0" & vbTab & "0")
                End Using



            End If



        Next i_emp_name_1




        ''New LOOP added for Error Log Reporting
        'For i_grp_name_1 As Integer = 0 To grp_names.Length() - 1
        '    If grp_names(i_grp_name_1) <> "" Then
        '        'excel_workbook.Names.Item("Error_RawName").RefersToRange.Value = grp_names(i_grp_name_1)
        '        leadDirName = output_folder & grp_names(i_grp_name_1)
        '        My.Computer.FileSystem.CreateDirectory(leadDirName)
        '        My.Computer.FileSystem.CreateDirectory(leadDirName.Replace("Excel Reports", "Published PDF Reports"))

        '        'Create Team File
        '        reportType_Val = "Team"
        '        finalFileName = leadDirName & "\"
        '        excel_error_log_wrksheet.Range("Error_RawName").Value = grp_names(i_grp_name_1)
        '        excel_error_log_wrksheet.Range("Error_ReportType").Value = reportType_Val
        '        xl_app.Calculate()
        '        finalFileName = finalFileName & "_" & excel_error_log_wrksheet.Range("Error_Name").Value & " - Team Report - "
        '        finalFileName = finalFileName & excel_error_log_wrksheet.Range("Error_FileNameAddition").Value & ".xlsm"



        '        OpenExcelWorkbook_ReportTemplate(excel_TMP_workbook_filename, excel_workbook, excel_error_log_wrksheet, finalFileName, xl_app, reportType_Val)


        '        For i_emp_name_1 As Integer = 0 To emp_names.Length() - 1

        '            If emp_names(i_emp_name_1) <> "" Then
        '                selected_id = emp_names(i_emp_name_1).Split("_"c)(0)
        '                selected_name = emp_names(i_emp_name_1).Split("_"c)(1)
        '                finalFileName = leadDirName & "\"


        '                If selected_id = grp_names(i_grp_name_1) Then


        '                    'Create Individual File
        '                    reportType_Val = "Individual"
        '                    excel_error_log_wrksheet.Range("Error_RawName").Value = selected_name
        '                    excel_error_log_wrksheet.Range("Error_ReportType").Value = reportType_Val
        '                    xl_app.Calculate()
        '                    finalFileName = finalFileName & excel_error_log_wrksheet.Range("Error_Name").Value
        '                    finalFileName = finalFileName & excel_error_log_wrksheet.Range("Error_FileNameAddition").Value & ".xlsm"
        '                    OpenExcelWorkbook_ReportTemplate(excel_TMP_workbook_filename, excel_workbook, excel_error_log_wrksheet, finalFileName, xl_app, reportType_Val)

        '                End If



        '            End If



        '        Next i_emp_name_1




        '    End If
        'Next i_grp_name_1
        'Exit Sub


        'excel_new_referrals_wrksheet.EnableAutoFilter = True    ' Enables the auto-filter

        ' Loops through the group names, matches non-blank group names, and adds those doctor's names to the cur_group array
        ''''        For i_grp_name_1 As Integer = 0 To grp_names.Length() - 1
        ''''            If grp_names(i_grp_name_1) <> "" Then   ' The current doctor has a group

        ''''                ' Checks if the current group has already been completed.
        ''''                ' If so, it skips to the next group.
        ''''                For i As Integer = 0 To grps_done.Length() - 1
        ''''                    If grp_names(i_grp_name_1) = grps_done(i) Then
        ''''                        GoTo skip1
        ''''                    End If
        ''''                Next i

        ''''                grps_done(grps_done_count) = grp_names(i_grp_name_1) ' Adds the current doctor's group to the grps_done array; this group will not be used again
        ''''                ReDim Preserve grps_done(grps_done.Length())         ' Increases the grps_done array size by one
        ''''                grps_done_count += 1

        ''''                ' Loop through the same group names array to find other doctors matching the current doc
        ''''                For i_grp_name_2 As Integer = 0 To grp_names.Length() - 1
        ''''                    If grp_names(i_grp_name_1) = grp_names(i_grp_name_2) Then
        ''''                        cur_group(cur_count) = doc_names(i_grp_name_2)       ' Adds the doctor matching the current doc (from the outside loop), to the cur_group array containing all of the matching docs
        ''''                        ReDim Preserve cur_group(cur_group.Length())         ' Increases the cur_group size by one
        ''''                        cur_count += 1
        ''''                    End If
        ''''                Next i_grp_name_2

        ''''                ReDim Preserve cur_group(cur_count - 1)    ' Removes the empty element in the last index of the array

        ''''            Else    ' The current doctor does not have a group
        ''''                cur_group(0) = doc_names(i_grp_name_1)
        ''''            End If

        ''''            'excel_new_referrals_wrksheet.EnableAutoFilter = True    ' Enables the auto-filter

        ''''            '' Gets the worksheet ready to be exported to a PDF
        ''''            'GetWorksheetReadyForExport(cur_group, excel_new_referrals_wrksheet, columns)

        ''''            ' Gets the full filename for the new pdf file yet-to-be-created
        ''''            full_output_filename = CreateFilename(output_folder, cur_group, grp_names(i_grp_name_1))


        ''''            'Exports the worksheet to a PDF (TF: will also now fax as needed by the program)
        ''''            ExportToEXL(full_output_filename, excel_workbook, excel_new_referrals_wrksheet, cur_group, grp_names(i_grp_name_1))


        ''''            full_output_filename = output_folder    ' Resets the output filename

        ''''skip1:
        ''''            ReDim cur_group(0)  ' Resets the array containing the current group of doctors
        ''''            cur_group(0) = ""   ' Cleans the first element
        ''''            cur_count = 0       ' Resets the count
        ''''        Next i_grp_name_1

        'excel_new_referrals_wrksheet.AutoFilter.ShowAllData()    ' Clears all the filters
        'ShowColumns(excel_new_referrals_wrksheet, columns)       ' Shows all hidden columns
        'excel_new_referrals_wrksheet.EnableAutoFilter = False    ' Disables the auto-filter
        'ApplyDefaultFilters(excel_workbook, excel_new_referrals_wrksheet)   ' Applies the default filters which were just cleared - 8/8/2020 TF commented out, due to error thrown

    End Sub

    ' Get the worksheet ready for export
    'Private Sub GetWorksheetReadyForExport(ByVal doctor_names() As String, ByVal xl_new_referrals_worksheet As Excel._Worksheet, ByVal columns() As String)
    '    Dim physician_range As Excel.Range  ' Contains the whole of the "Physician" column that is not blank
    '    Dim status_range As Excel.Range     ' Contains the whole of the "Status" column that is not blank
    '    Dim statuses() As String            ' Contains the statues that are to be shown
    '    Dim row_count As Integer = 0        ' Contains the row count of non-blank rows
    '    Dim columns_to_be_hidden() As String = {"B", "C", "G", "I", "K", "L", "N", "M", "O", "P", "Q"}   ' Array of columns to be hidden, and then shown after the worksheet is exported to a pdf


    '    statuses = {"Apt Scheduled",
    '                "Call insurance",
    '                "COVID",
    '                "Not Completed",
    '                "On-Hold(waiting For...)",
    '                "PAR Submitted",
    '                "Ready To Schedule"}


    '    ' Counts the number of non-blank rows in the support worksheet.
    '    ' Exits the loop once it reaches a blank physician cell
    '    For i As Integer = 2 To xl_new_referrals_worksheet.UsedRange.Rows.Count()
    '        If Not xl_new_referrals_worksheet.Cells(i, 1).value = "" Then
    '            row_count += 1
    '        Else
    '            Exit For
    '        End If
    '    Next i

    '    physician_range = xl_new_referrals_worksheet.Range("F1", "F" & row_count)     ' Gets the range to be filtered, which is the entire "Physician" column
    '    status_range = xl_new_referrals_worksheet.Range("A1", "A" & row_count)        ' Gets the range to be filtered, which is the entire "Status" column

    '    UnprotectWorksheet(xl_new_referrals_worksheet)  ' Unprotects the worksheet, which allows filters to be added

    '    physician_range.AutoFilter("6", doctor_names, Excel.XlAutoFilterOperator.xlFilterValues)    ' Autofilters the range with the doctors from the current group as the filter values
    '    status_range.AutoFilter("1", statuses, Excel.XlAutoFilterOperator.xlFilterValues)           ' Autofilters the range with the doctors from the current group as the filter values

    '    ProtectWorksheet(xl_new_referrals_worksheet)    ' Protects the worksheet

    '    HideColumns(xl_new_referrals_worksheet, columns)    ' Hides the columns uneeded for exportation
    'End Sub


    '' Exports the excel worksheet to a PDF
    'Private Sub ExportToEXL(ByVal pdf_output_filename As String, ByVal excel_wrkbk As Excel._Workbook, ByVal excel_wrksheet As Excel._Worksheet, ByVal doctor_names As String(), ByVal group_name As String)
    '    Dim str_header_name As String           ' The group or doctor name for the header
    '    Dim arry_worksheet_names As String()    ' An array of the worksheet names to be selected and then exported to a PDF
    '    Dim FaxStatusCheck As String 'string to check if report will be empty, then skip publishing it
    '    Dim vFaxNum, vContactName As String
    '    Dim EmptyReportCheck As String '11/8/2020 added by TF to skip reports that are empty
    '    Dim MissingFaxNum As String

    '    arry_worksheet_names = {"2nd Page", "CoverPage"}

    '    If group_name <> "" Then    ' There IS a group name
    '        str_header_name = group_name
    '    Else                        ' There is NO group name
    '        str_header_name = doctor_names(0)
    '    End If


    '    ' Fixes the settings for the excel worksheets so that they export to PDF in a coherent and readable way
    '    For i As Integer = arry_worksheet_names.Length() - 1 To 0 Step -1
    '        UnprotectWorksheet(excel_wrkbk.Worksheets(arry_worksheet_names(i)))
    '        With excel_wrkbk.Worksheets(arry_worksheet_names(i)).PageSetup
    '            .CenterVertically = False
    '            .Orientation = Excel.XlPageOrientation.xlPortrait
    '            .Zoom = False
    '            .FitToPagesWide = 1
    '            .FitToPagesTall = False
    '            .BottomMargin = 25
    '            .TopMargin = 50
    '            .RightMargin = 25
    '            .LeftMargin = 25
    '            .HeaderMargin = 25
    '            .PrintHeadings = False
    '            .PrintGridlines = False
    '            'If worksheet = excel_wrksheet.Name Then
    '            '    .CenterHeader = "Pending Patients For " & str_header_name
    '            'End If
    '        End With

    '        excel_wrkbk.Worksheets(arry_worksheet_names(i)).Calculate()

    '        If arry_worksheet_names(i) = "2nd Page" Then
    '            excel_wrkbk.Worksheets(arry_worksheet_names(i)).Cells.Rows.Autofit         ' Autofits all of the rows being used

    '            ' Loops through rows of the specified named range and hides and blank rows
    '            For row_i As Integer = excel_wrkbk.Worksheets(arry_worksheet_names(i)).Range("HideUnhideRange").Row To excel_wrkbk.Worksheets(arry_worksheet_names(i)).Range("HideUnhideRange").Rows.Count()
    '                If excel_wrkbk.Worksheets(arry_worksheet_names(i)).Cells(row_i, 2).Value = Nothing Then
    '                    excel_wrkbk.Worksheets(arry_worksheet_names(i)).Rows(row_i).Hidden = True
    '                End If
    '            Next row_i
    '        End If

    '        If arry_worksheet_names(i) = "2nd Page" Then
    '            'excel_wrkbk.Worksheets(arry_worksheet_names(i)).Cells.Rows.Autofit         ' Autofits all of the rows being used

    '            ' Loops through rows of the specified named range and hides and blank rows
    '            For row_i As Integer = excel_wrkbk.Worksheets(arry_worksheet_names(i)).Range("HideUnhideRange2").Row To excel_wrkbk.Worksheets(arry_worksheet_names(i)).Range("HideUnhideRange2").Rows.Count()
    '                If excel_wrkbk.Worksheets(arry_worksheet_names(i)).Cells(row_i, 2).Value = Nothing Then
    '                    excel_wrkbk.Worksheets(arry_worksheet_names(i)).Rows(row_i).Hidden = True
    '                End If
    '            Next row_i
    '        End If

    '        If arry_worksheet_names(i) = "CoverPage" Then
    '            PopulateGroupOrDoctorNamedRange(excel_wrkbk.Worksheets(arry_worksheet_names(i)), str_header_name)    ' Populates the group/doctor named range with the current group/doctor name

    '            '11/8/2020 - added by TF to skip empty reports
    '            excel_wrkbk.Worksheets(arry_worksheet_names(i)).Calculate()
    '            EmptyReportCheck = excel_wrkbk.Names.Item("CoverPage_FaxReport").RefersToRange.Value

    '            If EmptyReportCheck = "SKIP" Then
    '                'do nothing, skipping the export of the file
    '                'SendErrorEmail(pdf_output_filename & vbNewLine & EmptyReportCheck & vbNewLine & "SKIPPED")
    '                Exit Sub
    '            End If


    '        End If

    '        ProtectWorksheet(excel_wrkbk.Worksheets(arry_worksheet_names(i)))
    '    Next i


    '    excel_wrkbk.Sheets(arry_worksheet_names).Select     ' Selects the worksheets
    '    excel_wrkbk.ActiveSheet.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdf_output_filename, Excel.XlFixedFormatQuality.xlQualityStandard, True, False)  ' Exports the worksheets to a PDF
    '    excel_wrkbk.Sheets(arry_worksheet_names(0)).Select  ' Selects the first worksheet in the array [FIXES A BUG, DON"T DELETE]

    '    excel_wrkbk.Save()  ' Saves the workbook.

    '    'Checks for Faxing Report Options 11/5/2020 TF - added while integrating Auto-Faxing
    '    FaxStatusCheck = excel_wrkbk.Names.Item("CoverPage_FaxReport").RefersToRange.Value
    '    vFaxNum = excel_wrkbk.Names.Item("CoverPage_vFaxNumber").RefersToRange.Value
    '    vContactName = excel_wrkbk.Names.Item("CoverPage_vContactName").RefersToRange.Value


    '    If FaxStatusCheck = "SEND" Then
    '        'MessageBox.Show("Created " & pdf_output_filename)
    '        'SendFax(vFaxNum, vContactName, pdf_output_filename)
    '        'SEND FAX
    '        'SendErrorEmail(str_header_name & vbNewLine & FaxStatusCheck & vbNewLine & vFaxNum & vbNewLine & excel_wrkbk.Names.Item("CoverPage_vFaxYesNo").RefersToRange.Value & vbNewLine & "SEND")
    '    ElseIf FaxStatusCheck = "ERROR" Then
    '        'MessageBox.Show("Skipped " & pdf_output_filename)
    '        If vFaxNum.Length < 2 Then
    '            MissingFaxNum = "Missing Fax Number in BOX"
    '        Else
    '            MissingFaxNum = ""
    '        End If


    '        SendErrorEmail("File Name: " & str_header_name & vbNewLine & "Status: " & FaxStatusCheck & vbNewLine & "Fax Number: " & MissingFaxNum & " " & vFaxNum & vbNewLine & "Y/N - Supposed to send fax (per BOX): " & excel_wrkbk.Names.Item("CoverPage_vFaxYesNo").RefersToRange.Value)
    '        Exit Sub
    '    End If

    '    Exit Sub


    'End Sub


    '' Creates a filename for the output pdf, using the folder entered.
    '' Format: GROUPNAME.pdf or DOCTORNAME.pdf (uses the doctor's name if they don't have a group name)
    'Private Function CreateFilename(ByVal output_exl_folder As String, ByVal doctor_names() As String, ByVal group_name As String)
    '    ' Creates the filename
    '    If group_name <> "" Then    ' There IS a group name
    '        output_exl_folder += group_name
    '    Else                        ' There is NO group name
    '        output_exl_folder += doctor_names(0)
    '    End If

    '    output_exl_folder += ".pdf"  ' Adds the file type to the end of the filename

    '    Return output_exl_folder
    'End Function


    ' Hides the columns uneeded for exportation
    Private Sub HideColumns(ByVal excel_worksheet As Excel._Worksheet, ByVal columns() As String)
        UnprotectWorksheet(excel_worksheet) ' Unprotects the worksheet
        For Each col In columns
            excel_worksheet.Columns(col).Hidden = True  ' Hides the column
        Next col
        ProtectWorksheet(excel_worksheet)   ' Protects the worksheet
    End Sub


    ' Shows the columns uneeded for exportation
    Private Sub ShowColumns(ByVal excel_worksheet As Excel._Worksheet, ByVal columns() As String)
        UnprotectWorksheet(excel_worksheet) ' Unprotects the worksheet
        For Each col In columns
            excel_worksheet.Columns(col).Hidden = False  ' Unhides the column
        Next col
        ProtectWorksheet(excel_worksheet)   ' Protects the worksheet
    End Sub


    ' Creates a new directory within the folder entered
    ' with the current date as directory's name.
    ' Returns the directory's file path
    Private Function CreateNewDir(ByVal folder As String)
        Dim new_folder As String        ' The new folder
        Dim cur_date As Date            ' The current date
        Dim str_cur_date As String
        Dim replacment_folder_name As String

        cur_date = Date.Now.ToShortDateString()     ' Sets the current date
        str_cur_date = cur_date.ToString().Replace("/", "-").Remove(cur_date.ToString().IndexOf(" "))   ' Turns the date into a string and formats it
        str_cur_date = Format(Date.Now.AddMonths(0), "MMM-yyyy").ToString()

        replacment_folder_name = folder & str_cur_date
        If (Not System.IO.Directory.Exists(replacment_folder_name)) Then
            Directory.CreateDirectory(replacment_folder_name)    ' Creates a new directory with the current date as the folder name
        End If

        If (Not System.IO.Directory.Exists(replacment_folder_name.Replace("Excel Reports", "Published PDF Reports"))) Then
            Directory.CreateDirectory(replacment_folder_name.Replace("Excel Reports", "Published PDF Reports"))    ' Creates a new directory with the current date as the folder name
        End If


        new_folder = folder & str_cur_date & "\"    ' Saves the new folder path

        Return new_folder
    End Function


    ' Applies the default filters to the specified worksheet
    'Private Sub ApplyDefaultFilters(ByVal xl_workbook As Excel._Workbook, ByVal xl_worksheet As Excel._Worksheet)
    '    Dim statuses() As String            ' Contains the all of the possible statuses that are to be filtered (not all files use the same filters)
    '    Dim row_count As Integer = 0        ' Contains the row count of non-blank rows
    '    Dim status_range As Excel.Range     ' Contains the whole of the "Status" column that is not blank
    '    Dim first_row_after_header_i As Integer ' The index for the first row after the header

    '    Try
    '        statuses = {"Apt Scheduled",
    '                    "COVID",
    '                    "Not Completed",
    '                    "On-Hold(waiting for...)",
    '                    "Ready to Schedule",
    '                    "Call insurance",
    '                    "PAR Submitted",
    '                    ""}

    '        ' Gets the first row after the headers
    '        first_row_after_header_i = 1
    '        Do Until xl_worksheet.Cells(first_row_after_header_i, 1).Value = "Status"
    '            first_row_after_header_i += 1
    '        Loop
    '        first_row_after_header_i += 1

    '        row_count = xl_worksheet.UsedRange.Rows.Count() - 1   ' Gets the last row in the worksheet

    '        status_range = xl_worksheet.Range("A" & first_row_after_header_i, "A" & row_count)        ' Gets the range to be filtered, which is the entire "Status" column

    '        UnprotectWorksheet(xl_worksheet)        ' Unprotects the worksheet, which allows filters to be added
    '        xl_worksheet.EnableAutoFilter = True    ' Enables the autofilter
    '        status_range.AutoFilter("1", statuses, Excel.XlAutoFilterOperator.xlFilterValues)  ' Autofilters the range with the doctors from the current group as the filter values
    '        ProtectWorksheet(xl_worksheet)          ' Protects the worksheet

    '    Catch ex As Exception
    '        Windows.Forms.MessageBox.Show("Error: failed to apply status filters.", "Error")
    '        Exit Sub
    '    End Try

    '    Try
    '        xl_workbook.Save()              ' Saves the workbook
    '    Catch ex As Exception
    '        Windows.Forms.MessageBox.Show("Error: failed to save applied status filters", "Error")
    '        Exit Sub
    '    End Try

    'End Sub


    ' Unprotects the worksheet
    Private Sub UnprotectWorksheet(ByRef excel_worksheet As Excel._Worksheet)
        excel_worksheet.Unprotect("")
    End Sub


    ' Protects the worksheet
    Private Sub ProtectWorksheet(ByRef excel_worksheet As Excel._Worksheet)
        excel_worksheet.Protect("", False, True, False, , , , , , , True, , True, True, True, )
    End Sub


    ' Populates the group/doctor named range with the current group/doctor whose report is being created
    'Private Sub PopulateGroupOrDoctorNamedRange(ByVal xl_worksheet As Excel._Worksheet, ByVal group_or_doc_name As String)
    '    Dim coverpage_doctorname_range As Excel.Range ' The named range to be populated

    '    coverpage_doctorname_range = xl_worksheet.Range("Coverpage_DoctorName")   ' Gets the named range
    '    coverpage_doctorname_range.Cells.Value = group_or_doc_name                ' Sets the named range equal to the current group/doctor name
    'End Sub

    'Private Sub SendFax(vFaxNumber As String, vContactName As String, vFilePath As String)

    '    'Exit Sub ' TESTING

    '    Dim username As String = "vortechsolutions"
    '    Dim password As String = "!nt3rFax74"
    '    Dim faxNumbers As String = "+" & vFaxNumber '"+18884732963"
    '    Dim contacts As String = vContactName 'FILL IN
    '    Dim path1 As String = vFilePath 'FILL IN
    '    'Dim path1 As String = "C:\Users\timfr\Lehan Drugs\Lehan's HME - Order Processing - Order Processing\Reports\10-18-2020\Alberti, Lawrence E..pdf"
    '    'Dim path2 As String = "c:\temp\1.docx"
    '    ' read files data
    '    Dim file1data() As Byte = IO.File.ReadAllBytes(path1)   '1st document
    '    'Dim file2data() As Byte = IO.File.ReadAllBytes(path2)   '2nd document
    '    ' combine into a single byte array
    '    Dim data(file1data.Length - 1) As Byte 'FILL IN - maybe need to remove the "-1"
    '    Array.Copy(file1data, data, file1data.Length)
    '    'Array.Copy(file2data, 0, data, file1data.Length, file2data.Length)
    '    Dim fileTypes As String = IO.Path.GetExtension(path1).TrimStart("."c)
    '    Dim fileSizes As String = file1data.Length.ToString

    '    Dim postponeTime As DateTime = DateTime.Now.AddHours(-1) ' in two hours. use any PAST time to send ASAP
    '    Dim retriesToPerform As Integer = 10
    '    Dim csid As String = "VFAX" 'FILL IN??
    '    Dim pageHeader As String = ""
    '    'Dim pageHeader As String = "To: {To} From: {From} Pages: {TotalPages}" 'FILL IN??
    '    Dim subject As String = "PAP/NIV Patient Report - Lehan Drugs" '"PAP/NIV Patient Report - Lehan Drugs"
    '    Dim replyAddress As String = "tim@vortechsolutions.com;Mike@LehanDrugs.com"
    '    Dim pageSize As String = "Letter"
    '    Dim pageorientation As String = "Portrait"
    '    Dim isHighResolution As Boolean = True 'this will speed up your transmission
    '    Dim isFineRendering As Boolean = True  'fine will fit more graphics, while normal (false) will fit more textual documents

    '    Dim ifws As New interfax.InterFax()
    '    Dim st As Long = ifws.SendfaxEx_2(
    '    username,
    '    password,
    '    faxNumbers,
    '    contacts,
    '    data,
    '    fileTypes,
    '    fileSizes,
    '    postponeTime,
    '    retriesToPerform,
    '    csid,
    '    pageHeader,
    '    "",
    '    subject,
    '    replyAddress,
    '    pageSize,
    '    pageorientation,
    '    isHighResolution,
    '    isFineRendering)

    '    Console.WriteLine("Status is " & st)
    '    'MsgBox(st)
    'End Sub
    Private Sub SendErrorEmail(ByVal EmailText As String)
        Try
            Dim smtp_server As New SmtpClient
            Dim email As New MailMessage()
            Dim sender As String = "bot@vortechsolutions.com"
            Dim pw As String = "hYtYWTkrF5#onvhSFDGCe4AiHPM26V"
            'Dim sender As String = "vortechbot@lehandrugs.com"
            'Dim pw As String = "Lehan2020!"
            'Dim recipient As String = "dennis@vortechsolutions.com"
            Dim CCrecipient As String = "tim@vortechsolutions.com"
            'Dim recipient As String = "tim@vortechsolutions.com;Mike@LehanDrugs.com"
            Dim recipient1 As String = "tim@vortechsolutions.com"
            'Dim recipient2 As String = "Mike@LehanDrugs.com"

            smtp_server.UseDefaultCredentials = False
            smtp_server.Credentials = New Net.NetworkCredential(sender, pw)
            smtp_server.Port = 587
            smtp_server.EnableSsl = True
            smtp_server.Host = "smtp.office365.com"

            email.From = New MailAddress(sender)
            email.To.Add(recipient1)
            'email.To.Add(recipient2)
            email.CC.Add(CCrecipient)
            email.Subject = "Error Log Automation Error - " & Date.Now
            email.IsBodyHtml = False
            email.Body = EmailText

            smtp_server.Send(email)

            smtp_server.Dispose()

        Catch error_t As Exception
            MsgBox(error_t.ToString)
        End Try

    End Sub

End Class

Attribute VB_Name = "mdlStatic"
Option Explicit

Public Sub createReport()
    
    Dim audience As String: audience = "teacher"
    
    Select Case audience
        Case "teacher": Call createTeacherReport
        Case "student": Call createStudentReport
        Case "administrator": Call createAdministratorReport
    End Select
    
End Sub

Private Sub createTeacherReport()

    ThisWorkbook.Sheets(1).Cells(1, 1).Value = "Oktatsi_azonosito"
    ThisWorkbook.Sheets(1).Cells(1, 2).Value = "Diak_neve"
    ThisWorkbook.Sheets(1).Cells(1, 3).Value = "Osztaly"
    ThisWorkbook.Sheets(1).Cells(1, 4).Value = "Oktato"
    ThisWorkbook.Sheets(1).Cells(1, 5).Value = "Tantargy"
    ThisWorkbook.Sheets(1).Cells(1, 6).Value = "Erdemjegy"
    ThisWorkbook.Sheets(1).Cells(1, 7).Value = "Szazalek"

End Sub

Private Sub createStudentReport()

    ThisWorkbook.Sheets(1).Cells(1, 1).Value = "Oktatsi_azonosito"
    ThisWorkbook.Sheets(1).Cells(1, 2).Value = "Szuletesi_ido"
    ThisWorkbook.Sheets(1).Cells(1, 3).Value = "Osztaly"
    ThisWorkbook.Sheets(1).Cells(1, 4).Value = "Oktato"
    ThisWorkbook.Sheets(1).Cells(1, 5).Value = "Tantargy"
    ThisWorkbook.Sheets(1).Cells(1, 6).Value = "Erdemjegy"

End Sub

Private Sub createAdministratorReport()

    ThisWorkbook.Sheets(1).Cells(1, 1).Value = "Oktatsi_azonosito"
    ThisWorkbook.Sheets(1).Cells(1, 2).Value = "Szuletesi_ido"
    ThisWorkbook.Sheets(1).Cells(1, 3).Value = "Diak_neve"
    ThisWorkbook.Sheets(1).Cells(1, 4).Value = "Osztaly"
    ThisWorkbook.Sheets(1).Cells(1, 5).Value = "Oktato"
    ThisWorkbook.Sheets(1).Cells(1, 6).Value = "Tantargy"
    ThisWorkbook.Sheets(1).Cells(1, 7).Value = "Erdemjegy"
    ThisWorkbook.Sheets(1).Cells(1, 8).Value = "Szazalek"

End Sub

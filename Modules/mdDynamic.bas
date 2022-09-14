Attribute VB_Name = "mdDynamic"
Option Explicit

Private header_coll As Collection

Private Enum REPORT_AUDIENCE
    TEACHER = 1
    STUDENT = 2
    ADMINISTRATOR = 3
End Enum

Public Sub createReport()

    Rem create new workbook for report
    Dim report_wb As Workbook
    Set report_wb = Workbooks.Add
    
    Dim report_tpye As REPORT_AUDIENCE: report_tpye = TEACHER
    
    Rem crate header items
    Dim header As Variant
    header = createHeader(report_tpye)
    
    Rem add header to report workbook
    Call addHeader(report_wb.Worksheets(1), header)
    
End Sub

Private Sub addHeaderItem(ByRef header_arr As Variant, ByRef header_lv As Integer, ByVal header_text As String)

    If header_lv < 1 Then Err.Raise vbObjectError + 1, Description:="header_lv must be positive integer"

    ReDim Preserve header_arr(1 To header_lv)
    header_arr(header_lv) = header_text
    header_coll.Add header_lv, header_text
    header_lv = header_lv + 1
    
End Sub

Private Function createHeader(ByVal audience As REPORT_AUDIENCE) As Variant

    Set header_coll = New Collection
    
    Dim header_arr() As Variant
    Dim header_lv As Integer: header_lv = 1
    
    
    Call addHeaderItem(header_arr, header_lv, "Oktatasi_azonosito")
    
    If audience = REPORT_AUDIENCE.STUDENT _
        Or audience = REPORT_AUDIENCE.ADMINISTRATOR Then
        Call addHeaderItem(header_arr, header_lv, "Szuletesi_ido")
    End If
    
    If audience = REPORT_AUDIENCE.TEACHER _
        Or audience = REPORT_AUDIENCE.ADMINISTRATOR Then
        Call addHeaderItem(header_arr, header_lv, "Diak_neve")
    End If
    
    Call addHeaderItem(header_arr, header_lv, "Osztaly")
    Call addHeaderItem(header_arr, header_lv, "Oktato")
    Call addHeaderItem(header_arr, header_lv, "Tantargy")
    Call addHeaderItem(header_arr, header_lv, "Erdemjegy")
    
    If audience = REPORT_AUDIENCE.TEACHER _
        Or audience = REPORT_AUDIENCE.ADMINISTRATOR Then
        Call addHeaderItem(header_arr, header_lv, "Szazalek")
    End If
    
    createHeader = header_arr

End Function

Private Sub addHeader(ByRef sh As Worksheet, ByRef header_arr As Variant)

    sh.Range("A1").Resize(1, header_coll.Count).Value = header_arr

End Sub

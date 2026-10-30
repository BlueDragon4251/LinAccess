Attribute VB_Name = "Kundenformular"
Option Explicit

' Beispiel für die spätere VBA-Kompatibilitätsprüfung.
' Wird in LinAccess 0.2.0 bewusst NICHT ausgeführt.
Private Sub Form_Load()
    MsgBox "Kundenformular geöffnet"
End Sub

Private Sub btnSpeichern_Click()
    If IsNull(Me!Nachname) Then
        MsgBox "Bitte einen Nachnamen eingeben"
        Exit Sub
    End If

    DoCmd.RunCommand acCmdSaveRecord
    MsgBox "Datensatz gespeichert"
End Sub

Public Function AnzahlKunden() As Long
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Set db = CurrentDb()
    Set rs = db.OpenRecordset("SELECT COUNT(*) AS Total FROM Kunden")
    AnzahlKunden = rs!Total
    rs.Close
End Function

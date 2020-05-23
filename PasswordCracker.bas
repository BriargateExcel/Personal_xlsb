Attribute VB_Name = "PasswordCracker"
Option Explicit

Sub BreakPassword()

    Dim I As Integer, J As Integer, k As Integer
    Dim l As Integer, m As Integer, n As Integer
    Dim i1 As Integer, i2 As Integer, i3 As Integer
    Dim i4 As Integer, i5 As Integer, i6 As Integer
    Dim Password As String
    
    On Error Resume Next
    
    For I = 65 To 66: For J = 65 To 66: For k = 65 To 66
    For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
    For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
    For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
    
        Password = Chr(I) & Chr(J) & Chr(k) & _
        Chr(l) & Chr(m) & Chr(i1) & Chr(i2) & Chr(i3) & _
        Chr(i4) & Chr(i5) & Chr(i6) & Chr(n)
        
        Debug.Print Password
        
        ActiveSheet.Unprotect Password
        
        If ActiveSheet.ProtectContents = False Then
            MsgBox Password
            Debug.Print Password
            Exit Sub
        End If
    
    Next: Next: Next: Next: Next: Next
    Next: Next: Next: Next: Next: Next

End Sub


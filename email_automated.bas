Attribute VB_Name = "Módulo1"
Sub EnviarCorreo()
  

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    Dim correo As String
    Dim correo2 As String
    Dim correo3 As String
    Dim correo4 As String
    
    correo = ActiveSheet.Range("e14")
    correo2 = ActiveSheet.Range("e15")
    correo3 = ActiveSheet.Range("e16")
    correo4 = ActiveSheet.Range("e17")
    
 
    
    
    
    
    
    Dim adjunto As String
    Dim cc2 As Integer
    
    If Range("B12").Value = 1 Then
    adjunto = Range("K19")

    ElseIf Range("B12").Value = 2 Then
    adjunto = Range("K19") & vbCr & Range("K20")

    ElseIf Range("B12").Value = 3 Then
    adjunto = Range("K19") & vbCr & Range("K20") & vbCr & Range("K21")
    
    ElseIf Range("B12").Value = 4 Then
    adjunto = Range("K19") & vbCr & Range("K20") & vbCr & Range("K21") & vbCr & Range("K22")
    
    ElseIf Range("B12").Value = 5 Then
    adjunto = Range("K19") & vbCr & Range("K20") & vbCr & Range("K21") & vbCr & Range("K22") & vbCr & Range("K23")
    
    ElseIf Range("B12").Value = 6 Then
    adjunto = Range("K19") & vbCr & Range("K20") & vbCr & Range("K21") & vbCr & Range("K22") & vbCr & Range("K23") & vbCr & Range("K24")
    
    ElseIf Range("B12").Value = 7 Then
    adjunto = Range("K19") & vbCr & Range("K20") & vbCr & Range("K21") & vbCr & Range("K22") & vbCr & Range("K23") & vbCr & Range("K24") & vbCr & Range("K25")

    ElseIf Range("B12").Value = 8 Then
    adjunto = Range("K19") & vbCr & Range("K20") & vbCr & Range("K21") & vbCr & Range("K22") & vbCr & Range("K23") & vbCr & Range("K24") & vbCr & Range("K25") & vbCr & Range("K26")

    End If
        
    'If Range("e2").Value = 0 Then
    'Range("e2").Value = ""
    'End If
    'If Worksheets("Setting").Range("E16").Value = 0 Then
    'Worksheets("Setting").Range("E16").Value = ""
    'End If
    'If Worksheets("Setting").Range("E15").Value = 0 Then
    'Worksheets("Setting").Range("E15").Value = ""
    'End If
    'If Range("f2").Value = 0 Then
    'Range("f2").Value = ""
    'End If
    
 
    
 If Range("B11") = "Si" Then
    
On Error Resume Next

With OutMail
    .To = correo
    
    .cc = correo2 & ";" & correo3 & ";" & correo4 & ";" & Range("e2") & ";" & Range("f2") & ";" & Range("g2")
    
    .Body = Range("b3") & " " & Range("b4") & ":" & vbCr & vbCr & Range("b14") & " " & _
            vbCr & Range("b15") & " " & Range("b16") & vbCr & vbCr & Range("b17") & vbCr & vbCr & _
            adjunto & vbCr & vbCr & Range("b24") & vbCr & vbCr & Range("b25") & vbCr & vbCr & Range("b26")
            
            
            
    .Subject = Range("B10") 'asunto
    
    .Display
    
    .HTMLBody = .HTMLBody & "<img  src='C:\Users\cristian.gomez\Desktop\macros\Captura.png'>"
    
End With

On Error GoTo 0

Set PutMail = Nothing
Set OutApp = Nothing

Else
     
    
On Error Resume Next

With OutMail
    .To = correo
    
    .cc = correo2 & ";" & correo3 & ";" & correo4 & ";" & Range("e2") & ";" & Range("f2") & ";" & Range("g2")
    
    .Body = Range("b3") & " " & Range("b4") & ":" & vbCr & vbCr & Range("b14") & " " & _
            vbCr & Range("b15") & " " & Range("b16") & vbCr & vbCr & Range("b24") & vbCr & vbCr & Range("b25")
            
    .Subject = Range("B10") 'asunto
    
    .Display
    
    .HTMLBody = .HTMLBody & "<img  src='C:\Users\cristian.gomez\Desktop\macros\Captura.png'>"
    
   
End With

On Error GoTo 0

Set PutMail = Nothing
Set OutApp = Nothing

End If

End Sub


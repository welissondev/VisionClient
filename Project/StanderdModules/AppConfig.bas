Attribute VB_Name = "AppConfig"
Option Explicit

Public Function GetConnectionStrig() As String

       On Error GoTo Exception
       
       GetConnectionStrig = "Provider=Microsoft.ACE.Oledb.12.0;Data Source = " & ThisWorkbook.Path & "\App\Data\VisionBase.mdb"
       
       Exit Function
       
Exception:

       Call MsgBox("A função ''GetConnectionString'' da classe ''AppSettings'' gerou uma exeção!" + vbCr + vbCr + _
       "Tente novamente e se o erro persistir por favor, entre em contato com nosso suporte em:" + vbCr + vbCr + _
       "Suporte: diarioexcel.com.br", vbCritical, "Exception")
              
End Function


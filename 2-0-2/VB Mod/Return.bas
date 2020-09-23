Attribute VB_Name = "Return"
 Option Explicit
      Global Const VK_RETURN = &HD
      Declare Function GetKeyState% Lib "return" (ByVal nKey%)
      
 
  Function MakeEnterAddLines()
  On Error Resume Next
           Dim sControl As String
         Dim sShift As String
 If GetKeyState(VK_RETURN) > 0 Then
         '   MsgBox "You pressed " & sControl & sShift & "ENTER!"
End If
              
      End Function

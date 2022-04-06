Attribute VB_Name = "Scripts"
Sub Script01_UpdateAdhamProducts()
On Error GoTo errorhandler
Dim sqlText As String
    sqlText = ""
    sqlText = sqlText & " update adhamproducts set gasno = left(ProdBaCodeNo,1)where gasno is  null "
    sqlText = sqlText & " update adhamproducts set BarcodeSerNo = right(ltrim(rtrim(ProdBaCodeNo)),3) where BarcodeSerNo is  null and (len(ProdBaCodeNo)=16 or len(ProdBaCodeNo)=18)"
    de.con.Execute (sqlText)
Exit Sub
errorhandler:
MsgBox Err.Description
End Sub


Sub Script02_UpdateMaintCall_CallReceiverEmpNo_EmpNo()
On Error GoTo errorhandler
Dim sqlText As String
    sqlText = ""
    sqlText = sqlText & " Update m1  set callreceiverempno = a1.empno , callentryEmpNo = a1.empno"
    sqlText = sqlText & " From "
    sqlText = sqlText & " maintcall m1 inner join  "
    sqlText = sqlText & " empfullname a1 on ltrim(rtrim(m1.callreceiver)) = a1.fullname"
    sqlText = sqlText & "  Where m1.CallReceiverEmpNo Is Null Or m1.CallEntryEmpNo Is Null"
    
    de.con.Execute (sqlText)



Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

Sub Script03_UpdateManitCall_MaintState()
On Error GoTo errorhandler
Dim sqlText As String
sqlText = ""

sqlText = sqlText & " update m1 set callstate = " & EnumMaintCallState.Repared + EnumMaintCallState.UnderwayAndPrinted & "  from maintcall m1 inner join reparation r1 on m1.callno = r1.callno "
sqlText = sqlText & " update m1 set callstate = " & EnumMaintCallState.UnderwayAndPrinted & "   from maintcall m1 left outer  join reparation r1 on  m1.callno = r1.callno where r1.callno is null and m1.callstate <> 0"

de.con.Execute (sqlText)

Exit Sub
errorhandler:
MsgBox Err.Description
End Sub

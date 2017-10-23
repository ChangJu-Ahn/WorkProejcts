<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<%	
Dim pBDC002																'☆ : 조회용 ComProxy Dll 사용 변수 
Dim szRetValue
Dim szProcID
Dim szSpread1
Dim szSpread2

Call LoadBasisGlobalInf() 

Call HideStatusWnd

On Error Resume Next
Err.Clear

szProcID = Trim(Request("txtProcID"))
szSpread1 = Trim(Request("txtSpread1"))
szSpread2 = Trim(Request("txtSpread2"))

Set pBDC002 = Server.CreateObject("BDC002.clsVerify")

If Err.Number <> 0 Then
	Set pBDC002 = Nothing
	Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If

szRetValue = pBDC002.CreateVerify(gStrGlobalCollection, _
								  szProcID, _
								  szSpread1, _
								  szSpread2)

Set pBDC002 = Nothing													'☜: ComProxy Unload
If CheckSYSTEMError(Err,True) = True Then
	Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)
	Response.End														'☜: 비지니스 로직 처리를 종료함 
End If
%>
<SCRIPT LANGUAGE=vbscript>
	With parent
		.DbSaveOk
	End With
</SCRIPT>

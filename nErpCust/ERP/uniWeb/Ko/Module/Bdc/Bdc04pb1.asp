<%@ LANGUAGE=VBSCript %>
<% Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<%    
'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
'---------------------------------------Common-----------------------------------------------------------
Dim strKeyNo
Dim strUsrId
Dim strMode	                                                            '☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strProcesID
Dim strJobTitle
Dim strPlanTime
Dim arrStrExcel
Dim iArrRowVal
Dim objBDC004
Dim istrCode
Dim iLngRow

Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount

Dim iCUCount

Dim ii

Err.Clear																		'☜: Protect system from crashing

Call LoadBasisGlobalInf()

itxtSpread = ""
             
iCUCount = Request.Form("txtCUSpread").Count
             
itxtSpreadArrCount = -1
             
ReDim itxtSpreadArr(iCUCount)

For ii = 1 To iCUCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
Next

itxtSpread = Join(itxtSpreadArr,"")

'Call ServerMesgBox("itxtSpread : " & itxtSpread, vbInformation, I_MKSCRIPT)

strMode     = Request.Form("txtMode")										'☜ : 현재 상태를 받음 
strKeyNo    = Request.Form("txtKeyNo") 
strUsrId    = Replace(gUsrId, "'", "''")
strProcesID = FilterVar(Request.Form("txtProcessID"), "", "SNM")
strJobTitle = FilterVar(Request.Form("txtJobTitle"), "", "SNM")
strPlanTime = FilterVar(Request.Form("tmPlanTime"), "", "SNM")
'arrStrExcel = Request.Form("txtExcel")
'Response.Write arrStrExcel



On Error Resume Next
Err.Clear
Set objBDC004 = Server.CreateObject("BDC004.clsJobManager")
If CheckSYSTEMError(Err,True) = True Then
    Set objBDC004 = Nothing
    Response.End
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
End If

Call objBDC004.AddJob(gStrGlobalCollection, _
                      strProcesID, _
                      strJobTitle, _
                      strPlanTime, _
                      itxtSpread, _
                      istrCode)

If CheckSYSTEMError(Err,True) = True Then
    Set objBDC004 = Nothing
    Response.End
	Response.Write "<Script Language=VBScript>" & vbCrLF
	Response.Write "Call Parent.RemovedivTextArea" & vbCrLF
	Response.Write "</Script>" & vbCrLF
End If
Set objBDC004 = Nothing

Call DisplayMsgBox("210030", vbOKOnly, "", "", I_MKSCRIPT)	    ' 등록되었습니다!		

'등록후 화면 다시 로딩 
Response.Write "<Script Language=vbscript>"	& vbCr
Response.Write "    Parent.DbSaveOk "		& vbCr
Response.Write "</Script>"
%>



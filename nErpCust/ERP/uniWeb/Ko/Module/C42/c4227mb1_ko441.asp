<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%'======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        :
'*  3. Program ID           : C4227MB1_KO441.asp
'*  4. Program Name         : 품목그룹별손익차이(B)
'*  5. Program Desc         :
'*  6. Modified date(First) : 2008-06-16
'*  7. Modified date(Last)  :
'*  8. Modifier (First)     : Lee Sang Ho
'*  9. Modifier (Last)      :
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->


<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다
Call LoadBasisGlobalInf
Call loadInfTB19029B("Q", "S", "NOCOOKIE", "MB")

On Error Resume Next

Dim lgStrColorFlag
Dim ADF
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim Rs0, rs1, rs2, rs3, rs4, rs5, rs6
Dim strQryMode
Dim lgStrPrevKey
Dim lgStrPrevKey2
Dim strFlag
Dim strItemCd
Dim StrProdOrderNo
Dim StrWcCd
Dim StrTrackingNo
Dim StrSlCd
Dim strDeleteFlag
Dim strItemGroupCd
Dim strStatus
Dim strTemp
Dim i
Dim strSQL
Dim tmpTot
Dim txtfrdate, txttodate, txtplantcd, txtitemcd, txtsalesgrp, txtbpcd, txtlvl, txtitemacct

txtfrdate = request("txtConFromDt")
txttodate = request("txtConToDt")

if request("txtplantcd") = "" Then
	txtplantcd = "%"
Else
	txtplantcd = request("txtplantcd")
End If

if Request("txtitemcd") = "" Then
	txtitemcd = "%"
Else
	txtitemcd = Request("txtitemcd")
End If
If Request("txtsalesgrp") = "" Then
	txtsalesgrp = "%"
Else
	txtsalesgrp = Request("txtsalesgrp")
End If

If request("txtbpcd") = "" Then
	txtbpcd = "%"
Else
	txtbpcd = request("txtbpcd")
End If

If Request("txtitemacct") = "" Then
	txtitemacct = "%"
Else
	txtitemacct = Request("txtitemacct")
End If

txtlvl = Request("txtlvl")

	Const C_SHEETMAXROWS_D = 9999

	Call HideStatusWnd

	strQryMode = Request("lgIntFlgMode")
	
	'=======================================================================================================
	'	Main Query - Production Results Display
	'=======================================================================================================
	strSQL = "EXEC dbo.usp_C_C4227MA1_KO441 '" &  txtfrdate & "', '"  & _
                                                  txttodate & "', '"    & _
                                                  txtplantcd & "', '"    & _
                                                  txtitemacct & "', '"    & _
                                                  txtitemcd & "', '"    & _
                                                  txtsalesgrp & "', '"    & _
                                                  txtbpcd & "', "    & _
                                                  txtlvl
    Call SubOpenDB(lgObjConn)
    Call FncOpenRs("R",lgObjConn,lgObjRs,strSQL,"X","X")
    If lgObjRs.EOF And lgObjRs.BOF Then
	Call DisplayMsgBox("800506", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
	lgObjRs.Close
	Set rs0 = Nothing
	Response.Write "<Script Language=vbscript>" & vbcr
	Response.Write "	parent.DbQueryNotOk" & vbcr
	Response.Write "</Script>" & vbcr
	Response.End
    End If

    Response.Write "<Script Language=""VBScript"">"         & vbCrLf
    Response.Write "With Parent"                            & vbCrLf
    Response.Write "    .ggoSpread.Source = .frm1.vspdDataH" & vbCrLf
    Response.Write "    .ggoSpread.SSShowDataByClip """    

	Do While Not lgObjRs.EOF
		i = i+1
		Response.Write Chr(11) & ConvSPChars(lgObjRs(0))
		Response.Write Chr(11) & lgObjRs(1)
		Response.Write Chr(11) & lgObjRs(2)
		Response.Write Chr(11) & lgObjRs(3)
		Response.Write Chr(11) & lgObjRs(4)
		Response.Write Chr(11) & lgObjRs(5)
		Response.Write Chr(11) & lgObjRs(6)
		Response.Write Chr(11) & lgObjRs(7)
		Response.Write Chr(11) & lgObjRs(8)
		Response.Write Chr(11) & lgObjRs(9)
		Response.Write Chr(11) & lgObjRs(10)
		Response.Write Chr(11) & lgObjRs(11)
		Response.Write Chr(11) & i
		Response.Write Chr(11) & Chr(12)
			
		lgObjRs.MoveNext
    Loop
	
    Response.Write """"                                     & vbCrLf
    'Response.Write "    .lgStrColorFlag  = lgStrColorFlag  "                       & vbCrLf
    Response.Write "    .DBQueryOk  "                       & vbCrLf
	Response.Write "End with"                               & vbCrLf
    Response.Write "</Script>"                              & vbCrLf

    Call SubCloseRs(lgObjRs)
    Call SubCloseRs(lgObjConn)
    Response.End
    
'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub    

%>


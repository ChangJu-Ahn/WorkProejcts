<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m4114qb1
'*  4. Program Name         : 월별매입가계정현황 
'*  5. Program Desc         :  
'*  6. Component List       : 
'*  7. Modified date(First) : 2005/10/20
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Sim Hae Young
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
Dim rs1, rs2, rs3, rs4,rs5
Dim istrData
Dim iStrPoNo
Dim iLngMaxRow		' 현재 그리드의 최대Row
Dim iLngRow
Dim GroupCount
Dim lgCurrency
Dim index,Count     	' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
Dim lgDataExist
Dim lgPageNo


Dim strBP_NM

Dim lgOpModeCRUD
Dim Inti
Dim intARows
Dim intTRows
intARows=0
intTRows=0
On Error Resume Next                                                             '☜: Protect system from crashing
Err.Clear                                                                        '☜: Clear Error status

Dim strSpread																'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Call HideStatusWnd                                                               '☜: Hide Processing message

lgOpModeCRUD  = Request("txtMode")

Select Case lgOpModeCRUD
	Case CStr(UID_M0001)
		Call  SubBizQueryMulti()
End Select

Sub SubBizQueryMulti()
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgDataExist      = "No"
	iLngMaxRow = CLng(Request("txtMaxRows"))



	Call FixUNISQLData()		'☜ : DB-Agent로 보낼 parameter 데이타 set

	Call QueryData()			'☜ : DB-Agent를 통한 ADO query

	'-----------------------
	'Result data display area
	'-----------------------
%>
	<Script Language=vbscript>
		With parent
			.frm1.txtBpNm.Value	= "<%=strBP_NM%>"

			.frm1.txtBpCd.focus
			Set .gActiveElement = .document.activeElement

			If "<%=lgDataExist%>" = "Yes" Then

				'Show multi spreadsheet data from this line

				.ggoSpread.Source    = .frm1.vspdData
				.ggoSpread.SSShowData "<%=istrData%>"                  '☜: Display data

				.lgPageNo			 =  "<%=lgPageNo%>"				    '☜: Next next data tag
				
				.DbQueryOk <%=intARows%>,<%=intTRows%>

			End If
		End with
	</Script>
<%
End Sub

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_S  = 100
	Dim iLoopCount
	Dim iRowStr
	Dim ColCnt

	Const C_MV_DT               = 0
	Const C_BP_CD               = 1
	Const C_BP_NM               = 2
	Const C_MVMT_AMT_SUM        = 3
	Const C_IV_AMT_SUM          = 4
	Const C_BALANCE_AMT			= 5


	lgDataExist    = "Yes"

	If CLng(lgPageNo) > 0 Then
		rs0.Move     	= CLng(C_SHEETMAXROWS_S) * CLng(lgPageNo)                
		intTRows	= CLng(C_SHEETMAXROWS_S) * CLng(lgPageNo)
	End If

	'//Response.end

	'----- 레코드셋 칼럼 순서 ----------
	'-----------------------------------
	iLoopCount = 0

    	ReDim PvArr(C_SHEETMAXROWS_S - 1)

	Do while Not (rs0.EOF Or rs0.BOF)
		iLoopCount =  iLoopCount + 1
		iRowStr = ""


		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_MV_DT))	        
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_BP_CD))	        
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_BP_NM))		    

		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(C_MVMT_AMT_SUM), ggQty.DecPoint, 0)
		iRowStr = iRowStr & Chr(11) & ""
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(C_IV_AMT_SUM), ggQty.DecPoint, 0)
		iRowStr = iRowStr & Chr(11) & ""
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs0(C_BALANCE_AMT), ggQty.DecPoint, 0)	

		iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount

		If iLoopCount < C_SHEETMAXROWS_S Then
		   istrData = istrData & iRowStr & Chr(11) & Chr(12)

        Else
		   istrData = istrData & iRowStr & Chr(11) & Chr(12)
		   lgPageNo = lgPageNo + 1
		   Exit Do
		End If

		rs0.MoveNext
	Loop

	
	intARows = iLoopCount
	If iLoopCount  < C_SHEETMAXROWS_S Then                                      '☜: Check if next data exists
	  lgPageNo = ""
	End If

	rs0.Close                                                       '☜: Close recordset object
	Set rs0 = Nothing	                                            '☜: Release ADF
End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    On Error Resume Next

    SetConditionData = false

	If Not(rs1.EOF Or rs1.BOF) Then
		strBP_NM = rs1("BP_NM")
		Set rs1 = Nothing

	Else
		Set rs1 = Nothing
		If Len(Request("txtBpCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "공급처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
		    Exit Function
		End If
	End If

    SetConditionData = True
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
   	Dim strVal
   	Dim strVal1
	ReDim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	Redim UNIValue(1,4)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 
	strVal = ""
	UNISqlId(0) = "M4114QA1_KO441"		'상의Splead Query
	UNISqlId(1) = "M3111QA102"		'공급처 PopUp

	UNIValue(1,0) = "'zzzz'"

     If Request("gPlant") <> "" Then
        strVal = strVal & " AND a.PLANT_CD=" & FilterVar(Request("gPlant"),"''","S")
        strVal1 = strVal1 & " AND a.PLANT_CD=" & FilterVar(Request("gPlant"),"''","S")
     End If
     If Request("gPurGrp") <> "" Then
        strVal = strVal & " AND a.PUR_GRP=" & FilterVar(Request("gPurGrp"),"''","S")
        strVal1 = strVal1 & " AND b.PUR_GRP=" & FilterVar(Request("gPurGrp"),"''","S")
     End If
     If Request("gPurOrg") <> "" Then
        strVal = strVal & " AND a.PUR_ORG=" & FilterVar(Request("gPurOrg"),"''","S")
        strVal1 = strVal1 & " AND b.PUR_ORG=" & FilterVar(Request("gPurOrg"),"''","S")
     End If
     If Request("gBizArea") <> "" Then
        strVal = strVal & " AND a.MVMT_BIZ_AREA=" & FilterVar(Request("gBizArea"),"''","S")
        strVal1 = strVal1 & " AND b.IV_BIZ_AREA=" & FilterVar(Request("gBizArea"),"''","S")
     End If   
    UNIValue(0,0) = strVal
    UNIValue(0,1) = strVal1
    
	'From Date 
	UNIValue(0,2) = " '"& FilterVar(Trim(UCase(Request("txtFromDt"))), " " , "SNM") & "' "
	If Trim(Request("txtFromDt")) <> "" Then
		UNIValue(0,2) = " '"& FilterVar(Trim(UCase(Request("txtFromDt"))), " " , "SNM") & "' "
	End If

	'To Date
	UNIValue(0,3) = " '"& FilterVar(Trim(UCase(Request("txtToDt"))), " " , "SNM") & "' "
	If Trim(Request("txtToDt")) <> "" Then
		UNIValue(0,3) = " '"& FilterVar(Trim(UCase(Request("txtToDt"))), " " , "SNM") & "' "
	End If

	'공급처 
	UNIValue(0,4) = " '"& FilterVar(Trim(UCase(Request("txtBpCd"))), " " , "SNM") & "' "
	If Trim(Request("txtBpCd")) <> "" Then
		UNIValue(0,4) = " '"& FilterVar(Trim(UCase(Request("txtBpCd"))), " " , "SNM") & "' "
		UNIValue(1,0) = " '"& FilterVar(Trim(UCase(Request("txtBpCd"))), " " , "SNM") & "' "
	End If

	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode


End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
	On Error Resume Next
	Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
	Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
	Dim iStr

	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

	Set lgADF   = Nothing

	iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If

	'팝업필드 체크 
	If Setconditiondata = False Then Exit Sub

	If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
	Else

		Call  MakeSpreadSheetData()

	End If
End Sub



%>

<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1203mb4.asp
'*  4. Program Name         : Routing Detail Query
'*  5. Program Desc         :
'*  6. Component List       :
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2002/11/20
'*  9. Modifier (First)     : Im, HyunSoo
'* 10. Modifier (Last)      : Hong Chang Ho 
'* 11. Comment              : add 1 more line for rccp (2003.04.23) kjpark
'**********************************************************************************************
%>
<%
On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf

Dim ADF										'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg								'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag	'DBAgent Parameter 선언 
Dim rs0, rs1								'DBAgent Parameter 선언 
Dim gToday, iStrPrevKey, iLngMaxRows, iIntCnt
Dim BaseDate
Dim strData, strTemp
Dim TmpBuffer
Dim iTotalStr

Const C_SHEETMAXROWS_D = 50
	
BaseDate = UniConvYYYYMMDDToDate(gAPDateFormat, "2999", "12", "31")

iStrPrevKey = Request("lgStrPrevKey2")
gToday = UniConvDate(Request("lgCurDt"))
iLngMaxRows = Request("txtMaxRows")

'=======================================================================================================
'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
'=======================================================================================================
Redim UNISqlId(0)
Redim UNIValue(0, 4)

UNISqlId(0) = "p1203mb4"
	
UNIValue(0, 0) = FilterVar(Request("txtPlantCd"), "''", "S")
UNIValue(0, 1) = FilterVar(Request("RoutNo"), "''", "S")
If iStrPrevKey <> "" Then
	UNIValue(0, 2) = FilterVar(iStrPrevKey, "''", "S")
Else
	UNIValue(0, 2) = "''"
End If
UNIValue(0, 3) = " " & FilterVar(UniConvDate(Request("lgCurDt")), "''", "S") & "" 
UNIValue(0, 4) = " " & FilterVar(UniConvDate(Request("lgCurDt")), "''", "S") & "" 
	
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)

If (rs1.EOF And rs1.BOF) Then
	Call DisplayMsgBox("181400", vbOKOnly, "", "", I_MKSCRIPT)
	rs1.Close
	Set rs1 = Nothing
	Set ADF = Nothing
	Response.End
End If
	
Response.Write "<Script Language = VBScript>" & vbCrLf
Response.Write "Dim arrRet(0)" & vbCrLf
Response.Write "With parent" & vbCrLf
	
	If Not(rs1.EOF And rs1.BOF) Then
		
		'If C_SHEETMAXROWS_D < rs1.RecordCount Then 
		'	ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)
		'Else
			ReDim TmpBuffer(rs1.RecordCount - 1)
		'End If
	
		For iIntCnt = 0 To rs1.RecordCount - 1
			'If iIntCnt < C_SHEETMAXROWS_D Then 		
				strData = ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ConvSPChars(rs1("Wc_Cd"))
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ConvSPChars(rs1("Job_Cd"))
				strData = strData & Chr(11) & ""
       						
				strTemp = rs1("Inside_Flg")
				If  strTemp = "Y" Then
					strData = strData & Chr(11) & "사내"
				ElseIf strTemp = "N" Then
					strData = strData & Chr(11) & "외주"
				Else
					strData = strData & Chr(11) & ""		
				End If
		
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & UniDateClientFormat(gToday)
				strData = strData & Chr(11) & UniDateClientFormat(BaseDate)
				strData = strData & Chr(11) & rs1("Inside_Flg")
				strData = strData & Chr(11) & rs1("Rout_Order")
				strData = strData & Chr(11) & (iLngMaxRows + iIntCnt)
				strData = strData & Chr(11) & Chr(12)
				rs1.MoveNext
				
				TmpBuffer(iIntCnt) = strData
			'End If
		Next
		
		iTotalStr = Join(TmpBuffer, "")
		
		Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf
		
		'If rs1("OPR_NO") = Null Then
		'	Response.Write ".lgStrPrevKey2 = """"" & vbCrLf
		'Else
		'	Response.Write ".lgStrPrevKey2 = """ & Trim(rs1("OPR_NO")) & """" & vbCrLf
		'End If
	End If

	rs1.Close
	Set rs1 = Nothing

	'Response.Write "If .lgStrPrevKey2 <> """" Then"	& vbCrLf' GroupView 사이즈로 화면 Row수보다 쿼리가 작으면 다시 쿼리함 
	'	Response.Write "arrRet(0) = """ & ConvSPChars(Request("RoutNo")) & """" & vbCrLf
	'	Response.Write ".SetRoutCopy(arrRet)" & vbCrLf
	'Response.Write "Else" & vbCrLf    
		Response.Write ".SetRoutCopyOk(" & iLngMaxRows & " + 1)" & vbCrLf
	'Response.Write "End If" & vbCrLf

Response.Write "End With" & vbCrLf

Response.Write "</Script>" & vbCrLf	

Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing

%>

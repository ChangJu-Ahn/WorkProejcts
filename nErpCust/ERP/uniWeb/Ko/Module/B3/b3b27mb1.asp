<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%'======================================================================================================
'*  1. Module Name          : Item Master
'*  2. Function Name        : 
'*  3. Program ID           : b3b27mb1.asp 
'*  4. Program Name         : Called By B3B27MA1 (Class Query)
'*  5. Program Desc         : Lookup Class Information
'*  6. Modified date(First) : 2003/02/05
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Lee Woo Guen
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================


'On Error Resume Next								'☜: 

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "*", "NOCOOKIE", "MB")

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2				'DBAgent Parameter 선언 
Dim iIntCnt, iLngMaxRows, strQryMode, iStrPrevKey
Dim strData

Dim strClassCd
Dim strClassMgr

Dim strAvailableItem

Dim TmpBuffer
Dim iTotalStr

Const C_SHEETMAXROWS_D = 100

strQryMode = Request("lgIntFlgMode")
iStrPrevKey = Request("lgStrPrevKey1")
iLngMaxRows = Request("txtMaxRows")

'======================================================================================================
'	클래스명 처리해주는 부분 
'======================================================================================================
	IF Request("txtClassCd") <> "" Then

		Redim UNISqlId(0)
		Redim UNIValue(0, 0)
	
		UNISqlId(0) = "b3b22mb1a"

		strClassCd = Request("txtClassCd") 
		UNIValue(0, 0) = FilterVar(strClassCd, "''", "S")

		UNILock = DISCONNREAD :	UNIFlag = "1"

		Set ADF = Server.CreateObject("prjPublic.cCtlTake")
		strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
		If rs0.EOF And rs0.BOF Then
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtClassNm.value = """"" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		Else
			Response.Write "<Script Language=VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtClassNm.value = """ & ConvSPChars(rs0(0)) & """" & vbCrLf
			Response.Write "</Script>" & vbCrLf
		End If
	
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing	

	ELSE
		Response.Write "<Script Language=VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtClassNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	END IF

'=======================================================================================================
'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
'=======================================================================================================
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 2)
	
	UNISqlId(0) = "b3b27mb1b"
	
	strClassCd = Trim(Request("txtClassCd"))
	
	UNIValue(0, 0) = "^"	
	
	Select Case strQryMode
		Case CStr(OPMD_CMODE)
			UNIValue(0, 1) = FilterVar(strClassCd, "''", "S")
		Case CStr(OPMD_UMODE) 
			UNIValue(0, 1) = FilterVar(iStrPrevKey, "''", "S")
	End Select

	IF Request("cboClassMgr") = "" Then
		UNIValue(0, 2) = "|"
	ELSE
		strClassMgr = Request("cboClassMgr") 
		UNIValue(0, 2) = FilterVar(strClassMgr, "''", "S")
	END IF

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)
      
	If rs1.EOF And rs1.BOF Then
		Call DisplayMsgBox("122650", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
		rs1.Close
		Set rs1 = Nothing					
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
Response.Write "<Script Language = VBScript>" & vbCrLf
Response.Write "With parent" & vbCrLf
	
    If Not(rs1.EOF And rs1.BOF) Then
		
		If C_SHEETMAXROWS_D < rs1.RecordCount Then 

			ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)

		Else
			
			ReDim TmpBuffer(rs1.RecordCount - 1)

		End If
		
		For iIntCnt = 0 To rs1.RecordCount - 1 
			If iIntCnt < C_SHEETMAXROWS_D Then
				
				strData = ""
				strData = strData & Chr(11) & ConvSPChars(rs1("CLASS_CD"))
				strData = strData & Chr(11) & ConvSPChars(rs1("CLASS_NM"))
				strData = strData & Chr(11) & rs1("CLASS_DIGIT")
				strData = strData & Chr(11) & ConvSPChars(rs1("CHAR_CD1"))
				strData = strData & Chr(11) & ConvSPChars(rs1("CHAR_NM1"))
				strData = strData & Chr(11) & rs1("CHAR_VALUE_DIGIT1")
				strData = strData & Chr(11) & ConvSPChars(rs1("CHAR_CD2"))
				strData = strData & Chr(11) & ConvSPChars(rs1("CHAR_NM2"))
				strData = strData & Chr(11) & rs1("CHAR_VALUE_DIGIT2")
				strData = strData & Chr(11) & ConvSPChars(rs1("MINOR_NM_CLASS_MGR"))		

		        strData = strData & Chr(11) & (iLngMaxRows + iIntCnt)
				strData = strData & Chr(11) & Chr(12)
			
				rs1.MoveNext
				
				TmpBuffer(iIntCnt) = strData
				
			End If
		Next
		
		iTotalStr = Join(TmpBuffer, "")
		
		Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf
		
		If rs1.EOF Then
			Response.Write ".lgStrPrevKey1 = """"" & vbCrLf
		Else
			Response.Write ".lgStrPrevKey1 = """ & Trim(rs1("CLASS_CD")) & """" & vbCrLf
		End If
	End If	

	rs1.Close
	Set rs1 = Nothing
	
	Response.Write ".frm1.hClassCd.value = """ & ConvSPChars(Request("txtClassCd")) & """" & vbCrLf
	Response.Write ".frm1.hClassMgr.value = """ & ConvSPChars(Request("cboClassMgr")) & """" & vbCrLf

	Response.Write ".DbQueryOk()" & vbCrLf

Response.Write "End With" & vbCrLf

Response.Write "</Script>" & vbCrLf

Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>
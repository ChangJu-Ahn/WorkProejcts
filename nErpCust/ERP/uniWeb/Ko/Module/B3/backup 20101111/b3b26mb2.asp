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
'*  3. Program ID           : b3b26mb2.asp 
'*  4. Program Name         : Called By B3B26MA1 (Characteristic Value Query)
'*  5. Program Desc         : Lookup Characteristic Value Information
'*  6. Modified date(First) : 2003/02/07
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Lee Woo Guen
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

On Error Resume Next								'☜: 여기서 부터 개발자 비지니스 로직을 

Call HideStatusWnd									'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "*","NOCOOKIE","MB")

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1			'DBAgent Parameter 선언 
Dim iIntCnt, iLngMaxRows, strQryMode, iStrPrevKey
Dim strData
Dim strMode													'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim strCharCd
Dim strCharValueCd

Dim TmpBuffer
Dim iTotalStr

Const C_SHEETMAXROWS_D = 50

strQryMode = Request("lgIntFlgMode")
iStrPrevKey = Request("lgStrPrevKey2")
iLngMaxRows = Request("txtMaxRows")


'=======================================================================================================
'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
'=======================================================================================================
	Redim UNISqlId(0)
	Redim UNIValue(0, 1)
	
	UNISqlId(0) = "b3b26mb2b"	
	
	strCharCd = UCase(Trim(Request("txtCharCd")))
	UNIValue(0, 0) = FilterVar(strCharCd, "''", "S")

	If iStrPrevKey = "" Then
		strCharValueCd = UCase(Trim(Request("txtCharValueCd")))
	Else
		strCharValueCd = iStrPrevKey
	End If
	UNIValue(0, 1) = FilterVar(strCharValueCd, "|", "S")

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("122640", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
		Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "Call parent.SetActiveCell(parent.frm1.vspdData1,parent.frm1.vspdData1.ActiveCol,parent.frm1.vspdData1.ActiveRow,""M"",""X"",""X"")" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		rs0.Close
		Set rs0 = Nothing					
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	End If
	
Response.Write "<Script Language = VBScript>" & vbCrLf
Response.Write "With parent" & vbCrLf

    If Not(rs0.EOF And rs0.BOF) Then
		
		If C_SHEETMAXROWS_D < rs0.RecordCount Then 

			ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)

		Else
			
			ReDim TmpBuffer(rs0.RecordCount - 1)

		End If
		
		For iIntCnt = 0 To rs0.RecordCount - 1 
			If iIntCnt < C_SHEETMAXROWS_D Then
			
				strData = ""
				strData = strData & Chr(11) & ConvSPChars(rs0("CHAR_VALUE_CD"))
				strData = strData & Chr(11) & ConvSPChars(rs0("CHAR_VALUE_NM"))

		        strData = strData & Chr(11) & (iLngMaxRows + iIntCnt)
				strData = strData & Chr(11) & Chr(12)
			
				rs0.MoveNext
				
				ReDim Preserve TmpBuffer(iIntCnt)
				TmpBuffer(iIntCnt) = strData
				
			End If
		Next
		
		iTotalStr = Join(TmpBuffer, "")
		
		Response.Write ".ggoSpread.Source = .frm1.vspdData2" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf
		
		If rs0("CHAR_VALUE_CD") = Null Then
			Response.Write ".lgStrPrevKey2 = """"" & vbCrLf
		Else
			Response.Write ".lgStrPrevKey2 = """ & Trim(rs0("CHAR_VALUE_CD")) & """" & vbCrLf
		End If
	End If	

	rs0.Close
	Set rs0 = Nothing
	
	Response.Write ".frm1.hCharCd.value = """ & ConvSPChars(Request("txtCharCd")) & """" & vbCrLf

	Response.Write ".DbDtlQueryOk()" & vbCrLf

Response.Write "End With" & vbCrLf

Response.Write "</Script>" & vbCrLf

Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>

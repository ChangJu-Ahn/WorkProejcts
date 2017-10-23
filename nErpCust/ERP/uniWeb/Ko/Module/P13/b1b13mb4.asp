<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<%
'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : ADO Template
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/09/15
'*  7. Modified date(Last)  : 2000/09/26
'*  8. Modifier (First)     : KimTaeHyun
'*  9. Modifier (Last)      : KimTaeHyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

On Error Resume Next								'☜: 

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf() 

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'DBAgent Parameter 선언 
Dim strData
Dim TmpBuffer
Dim iTotalStr

Const C_SHEETMAXROWS_D = 100

Dim iIntPrevKey, iIntCnt, iLngMaxRows

iIntPrevKey = Request("lgStrPrevKey2")
	
Redim UNISqlId(0)
Redim UNIValue(0, 2)
	
UNISqlId(0) = "b1b13mb4"
UNIValue(0, 0) = FilterVar(Request("txtPlantCd") , "''", "S")			'Trim(Request("txtPlantCd"))
UNIValue(0, 1) = FilterVar(Request("txtItemCd") , "''", "S")			'Trim(Request("txtItemCd"))
	
If iIntPrevKey = "" Then	
	UNIValue(0, 2) = 0
Else
	UNIValue(0, 2) = FilterVar(Request("lgStrPrevKey2") , "''", "S")		'Ucase(Trim(lgIntPrevKey))		
End If
		
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

If rs0.EOF And rs0.BOF Then
	Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
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
				strData = strData & Chr(11) & ConvSPChars(rs0("ALT_ITEM_CD"))
				strData = strData & Chr(11) & ConvSPChars(rs0("ITEM_NM"))
				strData = strData & Chr(11) & ConvSPChars(rs0("SPEC"))
				strData = strData & Chr(11) & rs0("PRIORITY")
				strData = strData & Chr(11) & UniDateClientFormat(rs0("VALID_FROM_DT"))
				strData = strData & Chr(11) & UniDateClientFormat(rs0("VALID_TO_DT"))
				strData = strData & Chr(11) & rs0("SEQ")
		        strData = strData & Chr(11) & (iLngMaxRows + iIntCnt)
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(iIntCnt) = strData

				rs0.MoveNext
			End If
		Next
		
		iTotalStr = Join(TmpBuffer, "")

		Response.Write ".ggoSpread.Source = .frm1.vspdData2" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf
		
		If rs0("SEQ") = Null Then
			Response.Write ".lgStrPrevKey2 = """"" & vbCrLf
		Else
			Response.Write ".lgStrPrevKey2 = """ & Trim(rs0("SEQ")) & """" & vbCrLf
		End If
	End If
		
	rs0.Close
	Set rs0 = Nothing

	Response.Write ".DbDtlQueryOk()" & vbCrLf

Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>

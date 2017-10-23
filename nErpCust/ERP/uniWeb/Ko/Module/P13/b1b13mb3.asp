<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<%
'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b13mb3.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Component List		: PB6S101.cBLkUpPlt.B_LOOK_UP_PLANT
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Jung Yu Kyung
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
On Error Resume Next								'☜: 

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf

Const C_SHEETMAXROWS_D = 30

Dim ADF1
Dim ADF 														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag,rs0, rs1			'DBAgent Parameter 선언 

Dim iIntCnt, iLngMaxRows, iStrPrevKey
Dim strData
Dim TmpBuffer
Dim iTotalStr

Dim strPlantCd
Dim strItemCd
Dim strItemAcct
Dim strProcType
Dim strFromDt
Dim strToDt
Dim strValidFlg
Dim I1_plant_cd
Dim pPB6S101

'-----------------------
'Com action area
'-----------------------
Set pPB6S101 = Server.CreateObject("PB6S101.cBLkUpPlt")

If CheckSYSTEMError(Err,True) = True Then
	Response.End
End If

I1_plant_cd = Request("txtPlantCd")

Call pPB6S101.B_LOOK_UP_PLANT(gStrGlobalCollection, I1_plant_cd)

If CheckSYSTEMError(Err, True) = True Then
	Set pPB6S101 = Nothing															'☜: Unload Component
	Response.End
End If

Set pPB6S101 = Nothing															'☜: Unload Component

iStrPrevKey = Request("lgStrPrevKey")
	
'======================================================================================================
'	품목이름 처리해주는 부분 
'======================================================================================================
Redim UNISqlId(1)
Redim UNIValue(1, 0)
	
UNISqlId(0) = "122600sac"
UNISqlId(1) = "122700sab"
	
	
strItemCd = FilterVar(Request("txtItemCd") , "''", "S")

strPlantCd = FilterVar(Request("txtPlantCd") , "''", "S")

UNIValue(0, 0) = strItemCd
UNIValue(1, 0) = strPlantCd
	
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
If rs0.EOF And rs0.BOF Then
	Response.Write "<Script Language = VBScript>" &vbCrLf
		Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf							'☜: 화면 처리 ASP 를 지칭함 
	Response.Write "</Script>" &vbCrLf
Else
	Response.Write "<Script Language = VBScript>" &vbCrLf
		Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs0(0)) & """" & vbCrLf
	Response.Write "</Script>" &vbCrLf
End If
	
	
If rs1.EOF And rs1.BOF Then
	Response.Write "<Script Language = VBScript>" &vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf					'☜: 화면 처리 ASP 를 지칭함 
	Response.Write "</Script>" &vbCrLf
Else
	Response.Write "<Script Language = VBScript>" &vbCrLf
		Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1(0)) & """" & vbCrLf
	Response.Write "</Script>" &vbCrLf
End If

rs0.Close
rs1.Close
		
Set rs0 = Nothing
Set rs1 = Nothing

Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
	
'=======================================================================================================
'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
'=======================================================================================================
Redim UNISqlId(0)
Redim UNIValue(0, 10)
	
UNISqlId(0) = "127500sab"	
	
If iStrPrevKey <> "" Then
	strItemCd = FilterVar(iStrPrevKey, "''", "S")	
Else
	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(UCase(Request("txtItemCd")) , "''", "S")
	End IF
End If

IF Request("txtItemAcct") = "" Then
	strItemAcct = "|"
Else
	strItemAcct = FilterVar(UCase(Request("txtItemAcct")) , "''", "S")
End IF

IF Request("txtProcType") = "" Then
	strProcType = "|"
Else
	strProcType = FilterVar(UCase(Request("txtProcType")) , "''", "S")
End IF

IF Request("txtFromDt") = "" Then
	strFromDt = FilterVar("1900-01-01", "''", "S")
Else
	strFromDt = FilterVar(UniConvDate(Request("txtFromDt")) , "''", "S")
End IF
	
IF Request("txtToDt") = "" Then
	strToDt = FilterVar("2999-12-31", "''", "S")
Else
	strToDt = FilterVar(UniConvDate(Request("txtToDt")) , "''", "S")
End IF
	
	
IF Request("rdoValidFlg") = "A" Then
	strValidFlg = "|"
Else
	strValidFlg = FilterVar(UCase(Request("rdoValidFlg")) , "''", "S")
End IF

UNIValue(0, 0) = "^"
UNIValue(0, 1) = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
UNIValue(0, 2) = strItemCd
UNIValue(0, 3) = strItemAcct
UNIValue(0, 4) = strProcType
UNIValue(0, 5) = strFromDt	
UNIValue(0, 6) = strToDt
UNIValue(0, 7) = strValidFlg
UNIValue(0, 8) = FilterVar("P1001" , "''", "S")
UNIValue(0, 9) = FilterVar("P1003" , "''", "S")
UNIValue(0, 10) = FilterVar("P1002" , "''", "S")
	
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF1 = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF1.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1)
      
If rs1.EOF And rs1.BOF Then
	Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
	rs1.Close		
	Set rs1 = Nothing
	Set ADF = Nothing		
		
	Call HideStatusWnd
		
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
				strData = strData & Chr(11) & ConvSPChars(rs1("ITEM_CD"))
				strData = strData & Chr(11) & ConvSPChars(rs1("ITEM_NM"))
				strData = strData & Chr(11) & ConvSPChars(rs1("SPEC"))
				strData = strData & Chr(11) & rs1("MINOR_NM_ITEM_ACCT")
				strData = strData & Chr(11) & rs1("MINOR_NM_ITEM_CLASS")
				strData = strData & Chr(11) & rs1("MINOR_NM_PROC_TYPE")
				strData = strData & Chr(11) & ConvSPChars(rs1("BASIC_UNIT"))
				strData = strData & Chr(11) & rs1("PROD_ENV_NM")
				strData = strData & Chr(11) & UCase(rs1("PHANTOM_FLG"))
				strData = strData & Chr(11) & UCase(rs1("MPS_FLG"))
				strData = strData & Chr(11) & UCase(rs1("TRACKING_FLG"))
				strData = strData & Chr(11) & UCase(rs1("ORDER_TYPE"))
				strData = strData & Chr(11) & ConvSPChars(rs1("ORDER_RULE_NM"))
				strData = strData & Chr(11) & UCase(rs1("LOT_FLG"))
				strData = strData & Chr(11) & UCase(rs1("VALID_FLG"))				
				strData = strData & Chr(11) & UNIDateClientFormat(rs1("VALID_FROM_DT"))
				strData = strData & Chr(11) & UNIDateClientFormat(rs1("VALID_TO_DT"))
		        strData = strData & Chr(11) & (iLngMaxRows + iIntCnt)
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(iIntCnt) =strData

				rs1.MoveNext
			End If
		Next
		
		iTotalStr = Join(TmpBuffer,"")
		Response.Write ".ggoSpread.Source = .frm1.vspdData1" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf
		
		If rs1("ITEM_CD") = Null Then
			Response.Write ".lgStrPrevKey = """"" & vbCrLf
		Else
			Response.Write ".lgStrPrevKey = """ & Trim(rs1("ITEM_CD")) & """" & vbCrLf
		End If
	End If

	Response.Write ".frm1.hPlantCd.value = """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCrLf
	Response.Write ".frm1.hItemCd.value = """ & ConvSPChars(Request("txtItemCd")) & """" & vbCrLf
	Response.Write ".frm1.hItemAcct.value = """ & Request("txtItemAcct") & """" & vbCrLf
	Response.Write ".frm1.hProcType.value = """ & Request("txtProcType") & """" & vbCrLf
	Response.Write ".frm1.hFromDt.value = """ & Request("txtFromDt") & """" & vbCrLf
	Response.Write ".frm1.hToDt.value = """ & Request("txtToDt") & """" & vbCrLf
	Response.Write ".frm1.hValidFlg.value = """ & Request("rdoValidFlg") & """" & vbCrLf
		
	rs1.Close
	Set rs1 = Nothing

	Response.Write ".DbQueryOk()" & vbCrLf

Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf
Set ADF1 = Nothing												'☜: ActiveX Data Factory Object Nothing
%>

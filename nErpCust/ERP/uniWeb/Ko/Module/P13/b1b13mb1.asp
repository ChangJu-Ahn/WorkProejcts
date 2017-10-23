<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<%
'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : Manage Alternative Item (Query)
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Component List		: PB3S106.cBLkUpItemByPlt.B_LOOK_UP_ITEM_BY_PLANT_SVR
'*  6. Modified date(First) : 2000/09/15
'*  7. Modified date(Last)  : 2000/09/26
'*  8. Modifier (First)     : KimTaeHyun
'*  9. Modifier (Last)      : KimTaeHyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf

Dim pPB3S106
Dim I1_select_char
Dim I2_plant_cd
Dim I3_item_cd
Dim E6_b_plant
Dim E7_b_item

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'DBAgent Parameter 선언 
Dim iIntPrevKey, iIntCnt, iLngMaxRows
Dim strData

Dim TmpBuffer
Dim iTotalStr

' E6_b_plant
Const P027_E6_plant_cd = 0
Const P027_E6_plant_nm = 1

' E7_b_item
Const P027_E7_item_cd = 0
Const P027_E7_item_nm = 1
Const P027_E7_phantom_flg = 13

Const C_SHEETMAXROWS_D = 50

iLngMaxRows = Request("txtMaxRows")
iIntPrevKey = Request("lgStrPrevKey")	             
I2_plant_cd = Request("txtPlantCd")
I3_item_cd = Request("txtItemCd")

Set pPB3S106 = Server.CreateObject("PB3S106.cBLkUpItemByPlt")
    
If CheckSYSTEMError(Err, True) = True Then
	Response.End
End If
    
Call pPB3S106.B_LOOK_UP_ITEM_BY_PLANT_SVR(gStrGlobalCollection, "", _
            I2_plant_cd, I3_item_cd, , , , , _
            , E6_b_plant, E7_b_item)

If CheckSYSTEMError(Err, True) = True Then
	Set pPB3S106 = Nothing															'☜: Unload Component
	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "With parent.frm1" & vbCrLf
		Response.Write ".txtPlantNm.value = """ & ConvSPChars(E6_b_plant(P027_E6_plant_nm)) & """" & vbCrLf
		Response.Write ".txtItemNm.value = """ & ConvSPChars(E7_b_item(P027_E7_item_nm)) & """" & vbCrLf
	Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Response.End
End If

Set pPB3S106 = Nothing	

'------------------------------------------					
' 공장별 품목 일반정보 
'------------------------------------------
Response.Write "<Script Language=VBScript>" & vbCrLf
Response.Write "With parent.frm1" & vbCrLf
	Response.Write ".txtPlantNm.value = """ & ConvSPChars(E6_b_plant(P027_E6_plant_nm)) & """" & vbCrLf
	Response.Write ".txtItemNm.value = """ & ConvSPChars(E7_b_item(P027_E7_item_nm)) & """" & vbCrLf
Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf

'=======================================================================================================
'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
'=======================================================================================================
	
Redim UNISqlId(0)
Redim UNIValue(0, 2)
	
UNISqlId(0) = "b1b13mb1"
	
UNIValue(0, 0) = FilterVar(UCase(Request("txtPlantCd")), "''", "S") 
UNIValue(0, 1) = FilterVar(UCase(Request("txtItemCd")), "''", "S") 
	
If iIntPrevKey = "" Then	
	UNIValue(0, 2) = 0
Else
	UNIValue(0, 2) = UCase(Trim(iIntPrevKey))		
End If
		
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

If rs0.EOF And rs0.BOF Then
	Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
	Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "With parent.frm1" & vbCrLf
	Response.Write ".hPlantCd.value = Trim(.txtPlantCd.value)" & vbCrLf	
	Response.Write ".hItemCd.value = Trim(.txtItemCd.value)" & vbCrLf	
	Response.Write "Call parent.SetToolbar(""11001111001111"")" & vbCrLf	
	Response.Write "End With" & vbCrLf
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
				strData = strData & Chr(11) & ConvSPChars(rs0("ALT_ITEM_CD"))
				strData = strData & Chr(11) & ""
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
		
		Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf
		
		If rs0("SEQ") = Null Then
			Response.Write ".lgStrPrevKey = """"" & vbCrLf
		Else
			Response.Write ".lgStrPrevKey = """ & Trim(rs0("SEQ")) & """" & vbCrLf
		End If
	End If

	rs0.Close
	Set rs0 = Nothing

	Response.Write ".DbQueryOk()" & vbCrLf

Response.Write "End With" & vbCrLf
Response.Write "</Script>" & vbCrLf
Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>

<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b1b02mb3.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 2002/11/14
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Hong Chang Ho
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

On Error Resume Next								'☜: 

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "*", "NOCOOKIE", "MB")

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg												'Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2				'DBAgent Parameter 선언 
Dim iIntCnt, iLngMaxRows, strQryMode, iStrPrevKey
Dim strData
Dim TmpBuffer
Dim iTotalStr

Dim strItemCd
Dim strSumItemClass
Dim strItemAccount
Dim strItemGroup
Dim strStartDt
Dim strEndDt
Dim strAvailableItem

Const C_SHEETMAXROWS_D = 100

strQryMode = Request("lgIntFlgMode")
iStrPrevKey = Request("lgStrPrevKey1")
iLngMaxRows = Request("txtMaxRows")
	
'======================================================================================================
'	품목이름 처리해주는 부분 
'======================================================================================================
Redim UNISqlId(1)
Redim UNIValue(1, 0)
	
UNISqlId(0) = "122600sac"
UNISqlId(1) = "127400saa"
	
	
IF Request("txtItemCd") = "" Then
   strItemCd = "|"
ELSE
   strItemCd = Request("txtItemCd") 
END IF
	
IF Request("txtHighItemGroupCd") = "" Then
   strItemGroup = "|"
ELSE
   strItemGroup = Request("txtHighItemGroupCd") 
END IF
	
UNIValue(0, 0) = FilterVar(strItemCd, "", "SNM")
UNIValue(1, 0) = FilterVar(strItemGroup, "", "SNM")
	
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
      
If rs0.EOF And rs0.BOF Then
	Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
	Response.Write "</Script>" & vbCrLf
Else
	Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs0(0)) & """" & vbCrLf
	Response.Write "</Script>" & vbCrLf
End If
	
If Not Trim(Request("txtHighItemGroupCd")) = "" Then
	If rs1.EOF And rs1.BOF Then
		Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtHighItemGroupNm.value = """"" & vbCrLf
		Response.Write "</Script>" & vbCrLf
		Call DisplayMsgBox("127400", vbInformation, "", "", I_MKSCRIPT)
		Response.End
	Else
		Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "parent.frm1.txtHighItemGroupNm.value = """ & ConvSPChars(rs1(0)) & """" & vbCrLf
		Response.Write "</Script>" & vbCrLf
	End If
End If

rs0.Close
rs1.Close
		
Set rs0 = Nothing
Set rs1 = Nothing

Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
								'☜: ActiveX Data Factory Object Nothing
	
'=======================================================================================================
'	만약, 선언한 변수가 배열이라면 아래와같은 Fix 된 배열로 Redim 을 해서 넘겨줘야 한다.
'=======================================================================================================
	
Redim UNISqlId(0)
Redim UNIValue(0, 9)
	
UNISqlId(0) = "B1B02MB3"
IF Request("txtItemCd") = "" Then
   strItemCd = "|"
ELSE
   strItemCd = Request("txtItemCd") 
END IF
		
IF Request("cboItemAcct") = "" Then
   strItemAccount = "|"
ELSE
   strItemAccount = Request("cboItemAcct") 
END IF
	
IF Request("cboItemClass") = "" Then
   strSumItemClass = "|"
ELSE
   strSumItemClass = Request("cboItemClass") 
END IF
	
IF Request("txtFinishStartDt") = "" Then
   strStartDt = "|"
ELSE
   strStartDt = UniConvDate(Request("txtFinishStartDt"))
END IF
	
IF Request("txtFinishEndDt") = "" Then
   strEndDt = "|"
ELSE
   strEndDt = UniConvDate(Request("txtFinishEndDt")) 
END IF
	
IF Request("rdoDefaultFlg") = "A" Then
   strAvailableItem = "|"
ELSE
   strAvailableItem = Request("rdoDefaultFlg") 
	   
END IF
	
	
UNIValue(0, 0) = "^"
	
Select Case strQryMode
	Case CStr(OPMD_CMODE)
		UNIValue(0, 1) = FilterVar(strItemCd, "''", "S")
	Case CStr(OPMD_UMODE) 
		UNIValue(0, 1) = FilterVar(iStrPrevKey, "''", "S")
End Select

UNIValue(0, 2) = strSumItemClass
UNIValue(0, 3) = strItemAccount
UNIValue(0, 4) = FilterVar(strStartDt, "''", "S")
UNIValue(0, 5) = FilterVar(strEndDt, "''", "S")
UNIValue(0, 6) = strAvailableItem
UNIValue(0, 7) = FilterVar("P1001", "''", "S")
UNIValue(0, 8) = FilterVar("P1002", "''", "S")
IF Request("txtHighItemGroupCd") = "" Then
	UNIValue(0,9) = "|"
Else
	UNIValue(0,9) = "a.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtHighItemGroupCd"))	, "''", "S") & " ))"
End IF
	
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")
strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
If rs0.EOF And rs0.BOF Then
	Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
	Response.Write "<Script Language = VBScript>" & vbCrLf
	Response.Write "parent.DbQueryNotOk()" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	Set rs0 = Nothing					
	Response.End													'☜: 비지니스 로직 처리를 종료함 
End If
	
Response.Write "<Script Language = VBScript>" & vbCrLf
Response.Write "With parent" & vbCrLf
	strData = ""
    If Not(rs0.EOF And rs0.BOF) Then
    
	    If C_SHEETMAXROWS_D < rs0.RecordCount Then 

			ReDim TmpBuffer(C_SHEETMAXROWS_D - 1)

		Else
			
			ReDim TmpBuffer(rs0.RecordCount - 1)

		End If
		
		For iIntCnt = 0 To rs0.RecordCount - 1 
			If iIntCnt < C_SHEETMAXROWS_D Then
				strData = ""
				strData = strData & Chr(11) & ConvSPChars(rs0("ITEM_CD"))
				strData = strData & Chr(11) & ConvSPChars(rs0("ITEM_NM"))
				strData = strData & Chr(11) & ConvSPChars(rs0("FORMAL_NM"))
				strData = strData & Chr(11) & ConvSPChars(rs0("MINOR_NM_ITEM_ACCT"))
				strData = strData & Chr(11) & ConvSPChars(rs0("BASIC_UNIT"))
				strData = strData & Chr(11) & ConvSPChars(rs0("ITEM_GROUP_CD"))
				strData = strData & Chr(11) & ConvSPChars(rs0("ITEM_GROUP_NM"))
				strData = strData & Chr(11) & rs0("PHANTOM_FLG")
				strData = strData & Chr(11) & rs0("BLANKET_PUR_FLG")
				strData = strData & Chr(11) & ConvSPChars(rs0("BASE_ITEM_CD"))
				strData = strData & Chr(11) & ""
				strData = strData & Chr(11) & rs0("MINOR_NM_ITEM_CLASS")
				strData = strData & Chr(11) & rs0("VALID_FLG")
				strData = strData & Chr(11) & rs0("ITEM_IMAGE_FLG")
				strData = strData & Chr(11) & ConvSPChars(rs0("SPEC"))
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("UNIT_WEIGHT"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
				strData = strData & Chr(11) & ConvSPChars(rs0("UNIT_OF_WEIGHT"))
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("GROSS_WEIGHT"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
				strData = strData & Chr(11) & ConvSPChars(rs0("GROSS_UNIT"))
				strData = strData & Chr(11) & UniConvNumberDBToCompany(rs0("CBM"), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0)
				strData = strData & Chr(11) & ConvSPChars(rs0("CBM_DESCRIPTION"))
				strData = strData & Chr(11) & ConvSPChars(rs0("DRAW_NO"))
				strData = strData & Chr(11) & ConvSPChars(rs0("HS_CD"))
				strData = strData & Chr(11) & ConvSPChars(rs0("HS_UNIT"))
				strData = strData & Chr(11) & UNIDateClientFormat(rs0("VALID_FROM_DT"))
				strData = strData & Chr(11) & UNIDateClientFormat(rs0("VALID_TO_DT"))

		        strData = strData & Chr(11) & (iLngMaxRows + iIntCnt)
				strData = strData & Chr(11) & Chr(12)
				
				TmpBuffer(iIntCnt) = strData
			
				rs0.MoveNext
			End If
		Next
	
		iTotalStr = Join(TmpBuffer, "")
		Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf
		
		If rs0("ITEM_CD") = Null Then
			Response.Write ".lgStrPrevKey1 = """"" & vbCrLf
		Else
			Response.Write ".lgStrPrevKey1 = """ & Trim(rs0("ITEM_CD")) & """" & vbCrLf
		End If
	End If	
	
	Response.Write ".frm1.hItemCd.value = """ & ConvSPChars(Request("txtItemCd")) & """" & vbCrLf
	Response.Write ".frm1.hItemAcct.value = """ & ConvSPChars(Request("cboItemAcct")) & """" & vbCrLf
	Response.Write ".frm1.hSumItemClass.value = """ & ConvSPChars(Request("CboItemClass")) & """" & vbCrLf
	Response.Write ".frm1.hItemGroup.value = """ & ConvSPChars(Request("txtHighItemGroupCd")) & """" & vbCrLf
	Response.Write ".frm1.hStartDt.value = """ & UNIDateClientFormat(strstartdt) & """" & vbCrLf
	Response.Write ".frm1.hEndDt.value = """ & UNIDateClientFormat(strEnddt) & """" & vbCrLf
	Response.Write ".frm1.hAvailableItem.value = """ & Request("rdoDefaultFlg") & """" & vbCrLf

	Response.Write ".DbQueryOk()" & vbCrLf

Response.Write "End With" & vbCrLf

Response.Write "</Script>" & vbCrLf

Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>

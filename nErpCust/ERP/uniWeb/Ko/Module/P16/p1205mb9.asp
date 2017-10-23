<%@LANGUAGE = VBScript%>
<%'======================================================================================================
'*  1. Module Name          : Product
'*  2. Function Name        : 
'*  3. Program ID           : p1205mb9.asp 
'*  4. Program Name         : 
'*  5. Program Desc         : 자원구성정보조회 
'*  6. Modified date(First) : 2003/04/09
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Ryu Sung Won
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "P", "NOCOOKIE", "MB")

Dim ADF														'ActiveX Data Factory 지정 변수선언 
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag
Dim rs0, rs1, rs2, rs3, rs4
Dim iIntCnt, strQryMode
Dim strData, strTemp

Const C_SHEETMAXROWS_D = 100

Dim StrNextKey		' 다음 값 
Dim lgStrNextKey1	'item_cd
Dim lgStrNextKey2	'rout_no
Dim lgStrNextKey3	'opr_no
Dim lgStrNextKey4	'resourcd_cd
Dim lgStrNextKey5	'rank
Dim LngMaxRow		' 현재 그리드의 최대Row

Dim TmpBuffer
Dim iTotalStr

Dim pPB6S101
Dim strPlantCd
Dim strItemCd
Dim strRoutNo
Dim strNextKeys		'Next Keys' Set

Call HideStatusWnd

On Error Resume Next
Err.Clear

	strQryMode		= Request("lgIntFlgMode")
	LngMaxRow		= Request("txtMaxRows")

	If strQryMode = CStr(OPMD_CMODE) Then
			
		'======================================================================================================
		'	품목이름 처리해주는 부분 
		'======================================================================================================
		Redim UNISqlId(2)
		Redim UNIValue(2, 0)
			
		UNISqlId(0) = "122600sac"
		UNISqlId(1) = "122700sab"
		UNISqlId(2) = "181300sac"
			
		strPlantCd = Request("txtPlantCd")
		strItemCd = Request("txtItemCd")
		strRoutNo = Request("txtRoutNo") 

			
		UNIValue(0, 0) = FilterVar(strItemCd ,"''","S")
		UNIValue(1, 0) = FilterVar(strPlantCd,"''","S")
		UNIValue(2, 0) = FilterVar(strRoutNo,"''","S")

		UNILock = DISCONNREAD :	UNIFlag = "1"
			
		Set ADF = Server.CreateObject("prjPublic.cCtlTake")
		strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1,rs2)

		Response.Write "<Script Language=VBScript>" & vbCrLf

		If rs0.EOF And rs0.BOF Then
			Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf						'☜: 화면 처리 ASP 를 지칭함 
		Else
			Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(rs0(0)) & """" & vbCrLf		'☜: 화면 처리 ASP 를 지칭함 
		End If
			
		If rs1.EOF And rs1.BOF Then
			Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf					'☜: 화면 처리 ASP 를 지칭함 
		Else
			Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(rs1(0)) & """" & vbCrLf				'☜: 화면 처리 ASP 를 지칭함 
		End If
			
		If rs2.EOF And rs2.BOF Then
			Response.Write "parent.frm1.txtRoutNm.value = """"" & vbCrLf				'☜: 화면 처리 ASP 를 지칭함 
		Else
			Response.Write "parent.frm1.txtRoutNm.value = """ & ConvSPChars(rs2(0)) & """" & vbCrLf				'☜: 화면 처리 ASP 를 지칭함 

		End If
		Response.Write "</Script>" & vbCrLf

		rs0.Close
		rs1.Close
		rs2.Close
				
		Set rs0 = Nothing
		Set rs1 = Nothing
		Set rs2 = Nothing

		Set ADF = Nothing
	
	End If	'CMODE END


	'--------------------------
	' Main Query
	'--------------------------
	Redim UNISqlId(0)
	Redim UNIValue(0, 4)

	UNISqlId(0) = "p1205mb9a"

	strPlantCd = FilterVar(UCase(Request("txtPlantCd")), "''", "S")
	
	IF Trim(Request("txtItemCd")) = "" Then
	   strItemCd = "|"
	ELSE
	   strItemCd = FilterVar(UCase(Request("txtItemCd")), "''", "S")
	END IF
	
	IF Trim(Request("txtRoutNo")) = "" Then
	   strRoutNo = "|"
	ELSE
	   strRoutNo = FilterVar(UCase(Request("txtRoutNo")), "''", "S")
	END IF		
	
	lgStrNextKey1	= Request("lgStrNextKey1")	'item_cd
	lgStrNextKey2	= Request("lgStrNextKey2")	'rout_no
	lgStrNextKey3	= Request("lgStrNextKey3")	'opr_no
	lgStrNextKey4	= Request("lgStrNextKey4")	'resource_cd
	lgStrNextKey5	= Request("lgStrNextKey5")	'rank
	
	'---------------------------------------------------------------
	' Make Statements using Next Keys
	'---------------------------------------------------------------
	If CInt(strQryMode) = OPMD_UMODE Then
		strNextKeys = "(a.item_cd > " & FilterVar(UCase(lgStrNextKey1), "''", "S") & " OR " & vbCrLf
		
		strNextKeys = strNextKeys & "(a.item_cd = " & FilterVar(UCase(lgStrNextKey1), "''", "S") & " AND " & _
									"a.rout_no > " & FilterVar(UCase(lgStrNextKey2), "''", "S") & ") OR " & vbCrLf

		strNextKeys = strNextKeys & "(a.item_cd = " & FilterVar(UCase(lgStrNextKey1), "''", "S") & " AND " & _
									"a.rout_no = " & FilterVar(UCase(lgStrNextKey2), "''", "S") & " AND " & _
									"a.opr_no > " & FilterVar(UCase(lgStrNextKey3), "''", "S") & ") OR " & vbCrLf
									
		strNextKeys = strNextKeys & "(a.item_cd = " & FilterVar(UCase(lgStrNextKey1), "''", "S") & " AND " & _
									"a.rout_no = " & FilterVar(UCase(lgStrNextKey2), "''", "S") & " AND " & _
									"a.opr_no = " & FilterVar(UCase(lgStrNextKey3), "''", "S") & " AND " & _
									"a.resource_cd > " & FilterVar(UCase(lgStrNextKey4), "''", "S") & ") OR " & vbCrLf
		
		strNextKeys = strNextKeys & "(a.item_cd = " & FilterVar(UCase(lgStrNextKey1), "''", "S") & " AND " & _
									"a.rout_no = " & FilterVar(UCase(lgStrNextKey2), "''", "S") & " AND " & _
									"a.opr_no = " & FilterVar(UCase(lgStrNextKey3), "''", "S") & " AND " & _
									"a.resource_cd = " & FilterVar(UCase(lgStrNextKey4), "''", "S") & " AND " & _
									"a.rank > " & FilterVar(UCase(lgStrNextKey5), "''", "S") & ") OR " & vbCrLf
		
		strNextKeys = strNextKeys & "(a.item_cd = " & FilterVar(UCase(lgStrNextKey1), "''", "S") & " AND " & _
									"a.rout_no = " & FilterVar(UCase(lgStrNextKey2), "''", "S") & " AND " & _
									"a.opr_no = " & FilterVar(UCase(lgStrNextKey3), "''", "S") & " AND " & _
									"a.resource_cd = " & FilterVar(UCase(lgStrNextKey4), "''", "S") & " AND " & _
									"a.rank = " & FilterVar(UCase(lgStrNextKey5), "''", "S") & "))" & vbCrLf
					
	Else	'OPMD_CMODE
		strNextKeys = "|"
	End If
	
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = strPlantCd
	UNIValue(0, 2) = strItemCd
	UNIValue(0, 3) = strRoutNo
	UNIValue(0, 4) = strNextKeys

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)	'⊙: DB 에러코드, 메세지타입, %처리, 스크립트유형 
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.End
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
				strData = strData & Chr(11) & ConvSPChars(rs0("ITEM_CD"))
				strData = strData & Chr(11) & ConvSPChars(rs0("ITEM_NM"))
				strData = strData & Chr(11) & ConvSPChars(rs0("SPEC"))
				strData = strData & Chr(11) & ConvSPChars(rs0("ROUT_NO"))
				strData = strData & Chr(11) & ConvSPChars(rs0("ROUT_NM"))
				strData = strData & Chr(11) & ConvSPChars(rs0("OPR_NO"))
				strData = strData & Chr(11) & UniNumClientFormat(rs0("RANK"),ggQty.DecPoint,0)
				strData = strData & Chr(11) & ConvSPChars(rs0("RESOURCE_CD"))
				strData = strData & Chr(11) & ConvSPChars(rs0("RESOURCE_NM"))
				strData = strData & Chr(11) & ConvSPChars(rs0("RESOURCE_TYPE_NM"))
				strData = strData & Chr(11) & UniNumClientFormat(rs0("BOR_EFFICIENCY"),ggQty.DecPoint,0)
				
		        strData = strData & Chr(11) & (LngMaxRow + iIntCnt)
				strData = strData & Chr(11) & Chr(12)
			
				rs0.MoveNext
				
				TmpBuffer(iIntCnt) = strData
			End If
		Next
		
		iTotalStr = Join(TmpBuffer, "")

		Response.Write ".ggoSpread.Source = .frm1.vspdData" & vbCrLf
		Response.Write ".ggoSpread.SSShowDataByClip """ & iTotalStr & """" & vbCrLf
		
		If Trim(rs0("ITEM_CD")) = "" Then
			Response.Write ".lgStrNextKey1 = """"" & vbCrLf
			Response.Write ".lgStrNextKey2 = """"" & vbCrLf
			Response.Write ".lgStrNextKey3 = """"" & vbCrLf
			Response.Write ".lgStrNextKey4 = """"" & vbCrLf
			Response.Write ".lgStrNextKey5 = """"" & vbCrLf
		Else
			Response.Write ".lgStrNextKey1 = """ & Trim(rs0("ITEM_CD")) & """" & vbCrLf
			Response.Write ".lgStrNextKey2 = """ & Trim(rs0("ROUT_NO")) & """" & vbCrLf
			Response.Write ".lgStrNextKey3 = """ & Trim(rs0("OPR_NO")) & """" & vbCrLf
			Response.Write ".lgStrNextKey4 = """ & Trim(rs0("RESOURCE_CD")) & """" & vbCrLf
			Response.Write ".lgStrNextKey5 = """ & Trim(rs0("RANK")) & """" & vbCrLf
		End If
	End If	

	rs0.Close
	Set rs0 = Nothing
	
	Response.Write ".frm1.hPlantCd.value	= """ & ConvSPChars(Request("txtPlantCd")) & """" & vbCrLf
	Response.Write ".frm1.hItemCd.value		= """ & ConvSPChars(Request("txtItemCd")) & """" & vbCrLf
	Response.Write ".frm1.hRoutNo.value		= """ & ConvSPChars(Request("txtRoutNo")) & """" & vbCrLf

	Response.Write ".DbQueryOk" & vbCrLf
	Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf

Set ADF = Nothing												'☜: ActiveX Data Factory Object Nothing
%>

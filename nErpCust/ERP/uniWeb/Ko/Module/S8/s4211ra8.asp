<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : B/L관리 
'*  3. Program ID           : S4211RA8
'*  4. Program Name         : 통관참조(B/L등록)
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/07/30
'*  8. Modified date(Last)  : 2002/07/30
'*  9. Modifier (First)     : Hwangseongbae
'* 10. Modifier (Last)      : Hwangseongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2002/07/30	ADO Version
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>통관참조</TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              '☜: indicates that All variables must be declared in advance
'========================================================================================================
Const BIZ_PGM_ID 		= "s4211rb8.asp"                              '☆: Biz Logic ASP Name
'========================================================================================================
Const C_MaxKey          = 4                                            '☆: key count of SpreadSheet

Const C_PopApplicant	= 1
Const C_PopSalesGrp		= 2
Const C_PopPayTerms		= 3
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
Dim IsOpenPop  
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 
														    'Window가 여러 개 뜨는 것을 방지하기 위해 
														    'PopUp Window가 사용중인지 여부를 나타냄 
Dim lgBlnOpenedFlag
Dim	lgBlnApplicantChg
Dim lgBlnSalesGrpChg
Dim	lgBlnPayTermsChg

Dim arrParent

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
'20021227 kangjungu dynamic popup
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate

iDBSYSDate = "<%=GetSvrDate%>"

'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
'========================================================================================================
Function InitVariables()
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
    gblnWinEvent = False
        
	lgBlnApplicantChg = False		' 수입자 변경여부 
	lgBlnSalesGrpChg	= False		' 영업그룹 변경여부 
	lgBlnPayTermsChg	= False		' 매출채권 변경여부 
End Function
'=======================================================================================================
Sub SetDefaultVal()
	Dim iArrReturn
		
	With frm1
		.txtFromDt.Text = UNIDateClientFormat(UniConvDateAToB(UniConvDateToYYYYMM(EndDate, PopupParent.gDateFormat, "-") & "-01", PopupParent.gServerDateFormat ,PopupParent.gAPDateFormat))
		.txtToDt.Text = EndDate

		If PopupParent.gSalesGrp <> "" Then
			.txtSalesGrp.value = PopupParent.gSalesGrp
			Call txtSalesGrp_OnChange1()
		End If
	End With
		
	Redim iArrReturn(0)
	Self.Returnvalue = iArrReturn
End Sub
'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>	
End Sub
'========================================================================================================
Sub InitSpreadSheet()
			
	' 정상출고에 의한 통관 
	Call SetZAdoSpreadSheet("S4211RA801","S","A","V20030103",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
								C_MaxKey, "X","X")		
	' 외주출고에 의한 통관 적용 
	Call SetZAdoSpreadSheet("S4211RA802","S","B","V20030103",PopupParent.C_SORT_DBAGENT,frm1.vspdData1, _
								C_MaxKey, "X","X")
	
	Call SetSpreadLock() 	
	
End Sub
'========================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		.vspdData1.ReDraw = False
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		ggoSpread.SpreadLock 1 , -1
		.vspddata.OperationMode = 3
		.vspddata1.OperationMode = 3
		'------ Developer Coding part (End   ) -------------------------------------------------------------- 
		.vspdData.ReDraw = True
		.vspdData1.ReDraw = True
    End With
End Sub	
'========================================================================================================
Function OKClick()
	Dim iArrReturn
	
	' 정상출고에 의한 통관 
    If frm1.rdoSubcontract2.checked  Then		 
		With frm1
			If .vspdData.ActiveRow > 0 Then	
				Redim iArrReturn(3)
				
				.vspdData.Row = .vspdData.ActiveRow
				.vspdData.Col =  GetKeyPos("A",1)			' 통관관리번호 
				iArrReturn(0) = .vspdData.Text
				.vspdData.Col =  GetKeyPos("A",2)			' 수주번호 
				iArrReturn(1) = .vspdData.Text
				.vspdData.Col =  GetKeyPos("A",3)			' 매출채권형태 
				iArrReturn(2) = .vspdData.Text
				.vspdData.Col =  GetKeyPos("A",4)			' 매출채권형태명 
				iArrReturn(3) = .vspdData.Text
				
				Self.Returnvalue = iArrReturn
			End If
		End With
    End If
    
    ' 외주출고에 의한 통관 적용 
    If frm1.rdoSubcontract1.checked  Then
		 With frm1
			If .vspdData1.ActiveRow > 0 Then	
				Redim iArrReturn(3)
				
				.vspdData1.Row = .vspdData1.ActiveRow
				.vspdData1.Col =  GetKeyPos("B",1)			' 통관관리번호 
				iArrReturn(0) = .vspdData1.Text
				.vspdData1.Col =  GetKeyPos("B",2)			' 수주번호 
				iArrReturn(1) = .vspdData1.Text
				.vspdData1.Col =  GetKeyPos("B",3)			' 매출채권형태 
				iArrReturn(2) = .vspdData1.Text
				.vspdData1.Col =  GetKeyPos("B",4)			' 매출채권형태명 
				iArrReturn(3) = .vspdData1.Text
				
				Self.Returnvalue = iArrReturn
			End If
		End With
    End If
	
	Self.Close()
	
End Function
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function
'-------------------------------------------------------------------------------------------------------- 
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case pvIntWhere

	Case C_PopApplicant												
		iArrParam(1) = "dbo.b_biz_partner BP"			<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtApplicant.value)	<%' Code Condition%>
		iArrParam(3) = ""								<%' Name Cindition%>
		iArrParam(4) = "BP.bp_type IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") AND BP.usage_flag = " & FilterVar("Y", "''", "S") & " "		<%' Where Condition%>
		iArrParam(5) = frm1.txtApplicant.alt '"수입자"						<%' TextBox 명칭 %>
			
		iArrField(0) = "ED15" & PopupParent.gColSep & "BP.bp_cd"	<%' Field명(0)%>
		iArrField(1) = "ED30" & PopupParent.gColSep & "BP.bp_nm"	<%' Field명(1)%>
		    
		iArrHeader(0) = "수입자"					<%' Header명(0)%>
		iArrHeader(1) = "수입자명"					<%' Header명(1)%>

		frm1.txtApplicant.focus 
	Case C_PopSalesGrp												
		iArrParam(1) = "dbo.B_SALES_GRP"				<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtSalesGrp.value)		<%' Code Condition%>
		iArrParam(3) = ""								<%' Name Cindition%>
		iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "					<%' Where Condition%>
		iArrParam(5) = Trim(frm1.txtSalesGrp.alt)		<%' TextBox 명칭 %>
		
		iArrField(0) = "ED15" & PopupParent.gColSep & "SALES_GRP"		<%' Field명(0)%>
		iArrField(1) = "ED30" & PopupParent.gColSep & "SALES_GRP_NM"	<%' Field명(1)%>
    
	    iArrHeader(0) = "영업그룹"					<%' Header명(0)%>
	    iArrHeader(1) = "영업그룹명"				<%' Header명(1)%>

		frm1.txtSalesGrp.focus 
	Case C_PopPayTerms												
		iArrParam(1) = "dbo.b_minor"				<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtPayTerms.value)		<%' Code Condition%>
		iArrParam(3) = ""								<%' Name Cindition%>
		iArrParam(4) = " major_cd = " & FilterVar("B9004", "''", "S") & " "	<%' Where Condition%>
		iArrParam(5) = Trim(frm1.txtPayTerms.alt)		<%' TextBox 명칭 %>

		iArrField(0) = "ED15" & PopupParent.gColSep & "minor_cd"	<%' Field명(0)%>
		iArrField(1) = "ED30" & PopupParent.gColSep & "minor_nm"	<%' Field명(1)%>

		iArrHeader(0) = "결제방법"				<%' Header명(0)%>
		iArrHeader(1) = "결제방법명"				<%' Header명(1)%>

		frm1.txtPayTerms.focus 
	End Select
 
	iArrParam(0) = iArrParam(5)							<%' 팝업 명칭 %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPopup(iArrRet,pvIntWhere)
		OpenConPopup = True
	End If	
	
End Function
'========================================================================================================
Function OpenSortPopup()
	Dim arrRet,pvSheetNo,pvPgmID
	
	' 정상출고에 의한 통관 
    If frm1.rdoSubcontract2.checked Then
	    pvSheetNo = "A"
	    pvPgmID   = "S4211RA801"
	   
    End If
    
    ' 외주출고에 의한 통관 적용 
    If frm1.rdoSubcontract1.checked  Then
    
	    pvSheetNo = "B"
	    pvPgmID   = "S4211RA802"
	   
    End If
    
	On Error Resume Next 
	
	If IsOpenPop = True Then Exit Function
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(pvSheetNo),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "X" Then
	     Exit Function
	Else
	     Call ggoSpread.SaveXMLData(pvSheetNo,arrRet(0),arrRet(1))
         Call InitVariables
         Call InitSpreadSheet()
    End If
End Function
'-------------------------------------------------------------------------------------------------------
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	Select Case pvIntWhere
	Case C_PopApplicant
		frm1.txtApplicant.value = pvArrRet(0) 
		frm1.txtApplicantNm.value = pvArrRet(1)   
	Case C_PopSalesGrp
		frm1.txtSalesGrp.value = pvArrRet(0) 
		frm1.txtSalesGrpNm.value = pvArrRet(1)   
	Case C_PopPayTerms
		frm1.txtPayTerms.value = pvArrRet(0) 
		frm1.txtPayTermsNm.value = pvArrRet(1)   
	End Select

	SetConPopup = True

End Function
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029											  '⊙: Load table , B_numeric_format

    'Html에서 tag 숫자가 1과 2로 시작하는 부분 각각Format
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    
	Call InitVariables											  '⊙: Initializes local global variables
	Call SetDefaultVal	
	
	Call InitSpreadSheet()
	
	'정상출고에 의한 통관 display
	frm1.vspdData.style.display = "inline"   
    frm1.vspdData1.style.display = "none"
	
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	lgBlnOpenedflag = True
	DbQuery()
End Sub
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub
'==========================================================================================
Function txtApplicant_OnChange1()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtApplicant.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("C%", "''", "S") & "", "default", "default", "default", "" & FilterVar("BP", "''", "S") & "", C_PopApplicant) Then
				.txtApplicant.value = ""
				.txtApplicantNm.value = ""
				.txtApplicant.focus
			Else
				.txtfromDt.focus
			End If
			txtApplicant_OnChange1 = False
		Else
			.txtApplicantNm.value = ""
		End If
	End With
End Function
'==========================================================================================
Function txtSalesGrp_OnChange1()
	Dim iStrCode

	With frm1	
		iStrCode = Trim(.txtSalesGrp.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("SG", "''", "S") & "", C_PopSalesGrp) Then
				.txtSalesGrp.value = ""
				.txtSalesGrpNm.value = ""
				.txtSalesGrp.focus
			Else
				.txtPayTerms.focus
			End If
			txtSalesGrp_OnChange1 = False
		Else
			.txtSalesGrpNm.value = ""
		End If
	End With
End Function
'==========================================================================================
Function txtPayTerms_OnChange1()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtPayTerms.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("B9004", "''", "S") & "", "default", "default", "default", "" & FilterVar("MJ", "''", "S") & "", C_PopPayTerms) Then
				.txtPayTerms.value = ""
				.txtPayTermsNm.value = ""
				.txtPayTerms.focus
			Else
				.txtApplicant.focus
			End If
			txtPayTerms_OnChange1 = False
		Else
			.txtPayTermsNm.value = ""
		End If
	End With
End Function
'==========================================================================================
Function txtApplicant_OnKeyDown()
	lgBlnApplicantChg = True
	lgBlnFlgChgValue = True
End Function
'==========================================================================================
Function txtSalesGrp_OnKeyDown()
	lgBlnSalesGrpChg = True
	lgBlnFlgChgValue = True
End Function
'==========================================================================================
Function txtPayTerms_OnKeyDown()
	lgBlnPayTermsChg = True
	lgBlnFlgChgValue = True
End Function
'====================================================================================================
Function ChkValidityQueryCon()
	Dim iStrCode

	ChkValidityQueryCon = True

	If lgBlnApplicantChg Then
		iStrCode = Trim(frm1.txtApplicant.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("C%", "''", "S") & "", "default", "default", "default", "" & FilterVar("BP", "''", "S") & "", C_PopApplicant) Then
				Call DisplayMsgBox("970000", "X", frm1.txtApplicant.alt, "X")
				frm1.txtApplicant.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtApplicantNm.value = ""
		End If
		lgBlnApplicantChg	= False
	End If

	If lgBlnSalesGrpChg Then
		iStrCode = Trim(frm1.txtSalesGrp.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("SG", "''", "S") & "", C_PopSalesGrp) Then
				Call DisplayMsgBox("970000", "X", frm1.txtSalesGrp.alt, "X")
				frm1.txtSalesGrp.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtSalesGrpNm.value = ""
		End If
		lgBlnSalesGrpChg = False
	End If
			
	If lgBlnPayTermsChg Then
		iStrCode = Trim(frm1.txtPayTerms.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("B9004", "''", "S") & "", "default", "default", "default", "" & FilterVar("MJ", "''", "S") & "", C_PopPayTerms) Then
				Call DisplayMsgBox("970000", "X", frm1.txtPayTerms.alt, "X")
				frm1.txtPayTerms.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtPayTermsNm.value = ""
		End If
		lgBlnPayTermsChg = False
	End If

End Function
'====================================================================================================
Function GetCodeName(ByVal pvStrArg1, ByVal pvStrArg2, ByVal pvStrArg3, ByVal pvStrArg4, ByVal pvIntArg5, ByVal pvStrFlag, ByVal pvIntWhere)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs(2), iArrTemp
	
	GetCodeName = False
	
	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetCodeName (" & pvStrArg1 & ", " & pvStrArg2 & ", " & pvStrArg3 & ", " & pvStrArg4 & ", " & pvIntArg5 & ", " & pvStrFlag & ") "
	iStrWhereList = ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, Chr(11))
		iArrRs(0) = iArrTemp(1)
		iArrRs(1) = iArrTemp(2)
		iArrRs(2) = iArrTemp(3)
		GetCodeName = SetConPopup(iArrRs, pvIntWhere)
	Else
		'Item Change시 명을 Fetch하는 것으로 표준 변경시 Enable 시킨다.
		'If lgBlnOpenedFlag Then GetCodeName = OpenConPopup(pvIntWhere)
	End if
End Function
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
          Exit Function
    End If

	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function
'========================================================================================================
Function vspdData1_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData1.MaxRows = 0 Then 
          Exit Function
    End If

	If frm1.vspdData1.MaxRows > 0 Then
		If frm1.vspdData1.ActiveRow = Row Or frm1.vspdData1.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function
'========================================================================================================
    Function vspdData_KeyPress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
			Call OKClick()
		ElseIf KeyAscii = 27 Then
			Call CancelClick()
		End If
    End Function
'========================================================================================================
    Function vspdData1_KeyPress(KeyAscii)
		On Error Resume Next
		If KeyAscii = 13 And vspdData1.ActiveRow > 0 Then
			Call OKClick()
		ElseIf KeyAscii = 27 Then
			Call CancelClick()
		End If
    End Function
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
    If OldLeft <> NewLeft Then Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub
'========================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    

	If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub
'========================================================================================================
Sub txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7		
		Call SetFocusToDocument("P")   
		Frm1.txtFromDt.Focus
	End If
End Sub

Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("P")   
		Frm1.txtToDt.Focus
	End If
End Sub
'=======================================================================================================
Sub txtFromDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub

Sub txtToDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'*********************************************************************************************************
Function FncQuery() 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    
	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 크거나 같아야 할때 **
	With frm1
		If ValidDateCheck(.txtFromDt, .txtToDt) = False Then Exit Function

		If UniConvDateToYYYYMMDD(.txtFromDt.text , PopupParent.gDateFormat , "") > UniConvDateToYYYYMMDD(EndDate, PopupParent.gDateFormat , "") Then		
			Call DisplayMsgBox("970025", "X", .txtFromDt.ALT, "현재일" & "(" & EndDate & ")")
			.txtFromDt.focus	
			Exit Function
		End If

		If UniConvDateToYYYYMMDD(.txtToDt.text , PopupParent.gDateFormat , "") > UniConvDateToYYYYMMDD(EndDate, PopupParent.gDateFormat , "") Then		
			Call DisplayMsgBox("970025", "X", .txtToDt.ALT, "현재일" & "(" & EndDate & ")")	
			.txtToDt.Focus
			Exit Function
		End If
	End With
   
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")	         						'⊙: Clear Contents  Field

	' 조회조건 유효값 check
	If 	lgBlnFlgChgValue Then
		If Not ChkValidityQueryCon Then	Exit Function
	End If
	
    Call InitVariables 														'⊙: Initializes local global variables
    
    '-----------------------
    'Query function call area
    '-----------------------	
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
End Function
'========================================================================================================
Function DbQuery() 
	
	Dim pvSheetNo,pvPgmID
	Err.Clear														'☜: Protect system from crashing
	DbQuery = False													'⊙: Processing is NG

	If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
			
		' 정상출고에 의한 통관 
		If frm1.rdoSubcontract2.checked  Then
		     pvSheetNo = "A"
		     pvPgmID   = "S4211RA801"
		   
		     '정상출고에 의한 통관 display
		     frm1.vspdData.style.display = "inline"   
			 frm1.vspdData1.style.display = "none"
		End If
		
		' 외주출고에 의한 통관 적용 
         If frm1.rdoSubcontract1.checked  Then
	          pvSheetNo = "B"
	          pvPgmID   = "S4211RA802"
	   
	          '외주출고에 의한 통관 display
	          frm1.vspdData.style.display = "none"   
              frm1.vspdData1.style.display = "inline"
         End If
    Else
		If UCase(frm1.txtHSubcontractFlg.value) = "N" Then
			pvSheetNo = "A"
		ElseIf UCase(frm1.txtHSubcontractFlg.value) = "Y" Then
         		pvSheetNo = "B"
		End If     
         
    End If
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal
	
    With frm1
		strVal = BIZ_PGM_ID & "?txtHMode=" & PopupParent.UID_M0001					<%'☜: 비지니스 처리 ASP의 상태 %>
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			' Scroll시 
			strVal = strVal & "&txtApplicant=" & .txtHApplicant.value
			strVal = strVal & "&txtFromDt=" & .txtHFromDt.value
			strVal = strVal & "&txtToDt=" & .txtHToDt.value
			strVal = strVal & "&txtSalesGrp=" & .txtHSalesGrp.value
			strVal = strVal & "&txtPayTerms=" & .txtHPayTerms.value
			strVal = strVal & "&txtSubcontractFlg=" & Trim(.txtHSubcontractFlg.value)
		Else
			strVal = strVal & "&txtApplicant=" & Trim(.txtApplicant.value)
			' 처음 조회시 
			strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.Text)				<%'☆: 조회 조건 데이타 %>
			If Len(Trim(.txtToDt.text)) Then
				strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
			Else
				strVal = strVal & "&txtToDt=" & EndDate
			End if
			strVal = strVal & "&txtSalesGrp=" & Trim(.txtSalesGrp.value)
			strVal = strVal & "&txtPayTerms=" & Trim(.txtPayTerms.value)		
			IF .rdoSubcontract2.checked Then
				strVal = strVal & "&txtSubcontractFlg=" & "N"
			Else
				strVal = strVal & "&txtSubcontractFlg=" & "Y"
			End If
		End If
		
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType(pvSheetNo)
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList(pvSheetNo)
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList(pvSheetNo))

	End With    
	
	Call RunMyBizASP(MyBizASP, strVal)									<%'☜: 비지니스 ASP 를 가동 %>

    DbQuery = True    

End Function
'=========================================================================================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 

	If frm1.vspdData.MaxRows > 0 Then
		lgIntFlgMode = PopupParent.OPMD_UMODE
		frm1.vspdData.Focus
	Else
		frm1.txtApplicant.focus
	End If
	If frm1.vspdData1.MaxRows > 0 Then
		lgIntFlgMode = PopupParent.OPMD_UMODE
		frm1.vspdData1.Focus
	Else
		frm1.txtApplicant.focus
	End If

End Function
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<BODY SCROLL=NO TABINDEX="-1">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5>수입자</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtApplicant" ALT="수입자" SIZE=10 MAXLENGTH=10 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnApplicant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopApplicant">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="14"></TD>
						<TD CLASS="TD5" NOWRAP>송장작성일</TD>
						<TD CLASS="TD6" NOWRAP>
							<TABLE CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtFromDt" CLASS=FPDTYYYYMMDD tag="11X1" Alt="시작일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD>
										&nbsp;~&nbsp;
									</TD>
									<TD>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtToDt" CLASS=FPDTYYYYMMDD tag="11X1" Alt="종료일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5>영업그룹</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtSalesGrp" ALT="영업그룹" SIZE=10 MAXLENGTH=4 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopSalesGrp">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGrpNm" SIZE=20 TAG="14"></TD>
						<TD CLASS=TD5>결제방법</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtPayTerms" ALT="결제방법" SIZE=10 MAXLENGTH=20 TAG="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPayTerms" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopPayTerms">&nbsp;<INPUT TYPE=TEXT NAME="txtPayTermsNm" SIZE=20 TAG="14"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>외주출고여부</TD>
						<TD CLASS=TD6 NOWRAP>
							<INPUT TYPE="RADIO" CLASS="RADIO" VALUE="Y" NAME="rdoSubcontract" TAG="11X" ID="rdoSubcontract1"><LABEL FOR="rdoSubcontract1">Y</LABEL>&nbsp;&nbsp;&nbsp;
							<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoSubcontract" TAG="11X" VALUE="N" CHECKED ID="rdoSubcontract2"><LABEL FOR="rdoSubcontract2">N</LABEL>
						</TD>
						<TD CLASS=TD5></TD>
						<TD CLASS=TD6></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT="100%">
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TAG="23" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxRows" Value=0> <PARAM NAME="MaxCols" Value=0> </OBJECT>');</SCRIPT>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% TAG="23" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxRows" Value=0> <PARAM NAME="MaxCols" Value=0> </OBJECT>');</SCRIPT>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
										      <IMG SRC="../../../CShared/image/zpConfig_d.gif"  Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)"  OnClick="OpenSortPopup()"></IMG>			</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO NORESIZE framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
	
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtHApplicant" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSalesGrp" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHPayTerms" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSubcontractFlg" TAG="24">  
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>

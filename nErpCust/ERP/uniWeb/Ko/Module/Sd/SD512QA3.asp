<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 영업실적관리 
'*  3. Program ID           : SD512QA3
'*  4. Program Name         : 매출채권조회(판매유형2)
'*  5. Program Desc         : 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/04/08
'*  8. Modified date(Last)  : 2003/06/09
'*  9. Modifier (First)     : kang su hwan
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit
	
' External ASP File
'========================================
Const BIZ_PGM_ID 		= "SD512QB301.asp"
Const BIZ_PGM_ID1 		= "SD512QB302.asp"

' Constant variables 
'========================================
Const C_MaxKey          = 20

' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

' User-defind Variables
'========================================
Dim lgIsOpenPop                                          

Dim lgStartRow
Dim lgEndRow

Const C_PopBizArea		=	0		'사업장								
Const C_PopSalesGrp		=	1		'영업그룹 
Const C_PopBillToParty	=	2		'발행처 
Const C_PopTaxBiz		=	3		'세금신고사업장							

Dim lgStrColorFlag
Dim lgStrColorFlag1

Dim ToDateOfDB

ToDateOfDB = UNIConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat,parent.gDateFormat)

'========================================	
Sub InitVariables()

    lgPageNo     = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                      'Indicates that current mode is Create mode
    lgSortKey        = 1
    
    Call SetToolBar("1100000000001111")										
    
End Sub

'========================================
Sub SetDefaultVal()
	frm1.txtConFromDt.Text = UNIGetFirstDay(ToDateOfDB, Parent.gDateFormat)
	frm1.txtConToDt.Text = 	ToDateOfDB
End Sub							
	
'========================================
Sub LoadInfTB19029()

	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "S","NOCOOKIE","QA") %>                                 
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub

'========================================
Sub InitSpreadSheet()

    Call SetZAdoSpreadSheet("SD512QA3","S","A", "V20030404", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
    Call SetZAdoSpreadSheet("SD512QA31","S","B", "V20030404", parent.C_SORT_DBAGENT,frm1.vspdData1,C_MaxKey, "X", "X")
    Call SetSpreadLock()   
     
End Sub

'========================================
Sub SetSpreadLock()
	
    With frm1
		ggoSpread.Source = .vspdData
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock 1 , -1
		.vspdData.ReDraw = True
    End With
    
    With frm1
		ggoSpread.Source = .vspdData1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock 1 , -1
		.vspdData.ReDraw = True
    End With
End Sub

'========================================
Sub Form_Load()

    Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")

	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()

    Call SetFocusToDocument("M")	
    frm1.txtConFromDt.focus
End Sub

'========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================
Function FncQuery() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncQuery = False                                                              '⊙: Processing is NG
    
    If Not chkField(Document, "1") Then Exit Function

	If ValidDateCheck(frm1.txtConFromDt, frm1.txtConToDt) = False Then Exit Function
	
    Call ggoOper.ClearField(Document, "2")
    
    Call InitVariables
    
    If DbQuery Then Exit Function

    If Err.number = 0 Then FncQuery = True

    Set gActiveElement = document.ActiveElement  
    
End Function

'========================================
Function FncPrint()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
	Call Parent.FncPrint()                                                        '☜: Protect system from crashing

    If Err.number = 0 Then
       FncPrint = True                                                            '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function FncExcel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	Call Parent.FncExport(parent.C_MULTI)
    
    If Err.number = 0 Then
       FncExcel = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function FncFind() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

	Call Parent.FncFind(parent.C_MULTI, True)

    If Err.number = 0 Then
       FncFind = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function FncExit()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = True                                                             '⊙: Processing is OK
End Function

'========================================
Function DbQuery() 

	Dim strVal
	Dim strBillConfFlag
	Dim strExceptFlag
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbQuery = False
    
	Call LayerShowHide(1)
    
    With frm1

		If .rdoBillConfFlag(0).checked Then
			strBillConfFlag=""
		ElseIf .rdoBillConfFlag(1).checked Then
			strBillConfFlag="Y"
		ElseIf .rdoBillConfFlag(2).checked Then
			strBillConfFlag="N"
		End If
		
		If .rdoExceptFlag(0).checked Then
			strExceptFlag=""
		ElseIf .rdoExceptFlag(1).checked Then
			strExceptFlag="Y"
		ElseIf .rdoExceptFlag(2).checked Then
			strExceptFlag="N"
		End If

        If lgIntFlgMode  <> parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?ConFromDt=" & Trim(.txtConFromDt.Text)
			strVal = strVal & "&ConToDt=" & Trim(.txtConToDt.Text)
			strVal = strVal & "&BizAreaCd=" & Trim(.txtConBizAreaCd.value)
			strVal = strVal & "&SalesGrpCd=" & Trim(.txtConSalesGrpCd.value)
			strVal = strVal & "&BillToPartyCd=" & Trim(.txtConBillToPartyCd.value)
			strVal = strVal & "&TaxBizCd=" & Trim(.txtConTaxBizCd.value)
			strVal = strVal & "&BillConfFlag=" & Trim(strBillConfFlag)	
			strVal = strVal & "&ExceptFlag=" & Trim(strExceptFlag)	
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))				
        End If    
                  
        lgStartRow = .vspdData.MaxRows + 1										'포멧팅 적용하는 시작Row
        
    End With

    Call RunMyBizASP(MyBizASP, strVal)
		
    If Err.number = 0 Then
       DbQuery = True
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function DbQueryOk()												

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    lgIntFlgMode     = parent.OPMD_UMODE										  '⊙: Indicates that current mode is Update mode
    
	Call SetQuerySpreadColor
	Call DbQuery2(1)
    Call SetToolBar("1100000000011111")
    
	frm1.vspdData.focus
    Set gActiveElement = document.ActiveElement   

End Function

Function DbQueryOk1()												

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	Call SetQuerySpreadColor1
	
    Set gActiveElement = document.ActiveElement   

End Function

Function DbQuery2(byVal pRow)
	Dim strVal
	Dim strBillToParty
	Dim strDealType
	Dim strBillConfFlag
	Dim strExceptFlag
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbQuery2 = False

	Call LayerShowHide(1)

	frm1.vspddata.row = pRow
	frm1.vspddata.col = GetKeyPos("A", 4)	'Bill to party
	strBillToParty = Trim(frm1.vspddata.text)
	frm1.vspddata.col = GetKeyPos("A", 2)	'Deal Type
	strDealType = Trim(frm1.vspddata.text)

	If frm1.rdoBillConfFlag(0).checked Then
		strBillConfFlag=""
	ElseIf frm1.rdoBillConfFlag(1).checked Then
		strBillConfFlag="Y"
	ElseIf frm1.rdoBillConfFlag(2).checked Then
		strBillConfFlag="N"
	End If
		
	If frm1.rdoExceptFlag(0).checked Then
		strExceptFlag=""
	ElseIf frm1.rdoExceptFlag(1).checked Then
		strExceptFlag="Y"
	ElseIf frm1.rdoExceptFlag(2).checked Then
		strExceptFlag="N"
	End If

	strVal = BIZ_PGM_ID1 & "?BillToParty=" & strBillToParty			
	strVal = strVal & "&DealType=" & strDealType
	strVal = strVal & "&ConFromDt=" & Trim(frm1.txtConFromDt.Text)
	strVal = strVal & "&ConToDt=" & Trim(frm1.txtConToDt.Text)
	strVal = strVal & "&BillConfFlag=" & Trim(strBillConfFlag)	
	strVal = strVal & "&ExceptFlag=" & Trim(strExceptFlag)	
	strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("B")			 
	strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("B")
	strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("B"))				

    Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

    If Err.number = 0 Then
       DbQuery2 = True																'⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   
	
End Function

'========================================
Sub SetQuerySpreadColor()

	Dim iArrColor1, iArrColor2
	Dim iLoopCnt
	
	iArrColor1 = Split(lgStrColorFlag,Parent.gRowSep)
	
	For iLoopCnt=0 to ubound(iArrColor1,1) - 1
		iArrColor2 = Split(iArrColor1(iLoopCnt),Parent.gColSep)

		frm1.vspdData.Col = -1
		frm1.vspdData.Row =  iArrColor2(0)
		
		Select Case iArrColor2(1)
			Case "1"
				frm1.vspdData.BackColor = RGB(204,255,153) '연두 
			Case "2"
				frm1.vspdData.BackColor = RGB(176,234,244) '하늘색 
			Case "3"
				frm1.vspdData.BackColor = RGB(224,206,244) '연보라 
			Case "4"  
				frm1.vspdData.BackColor = RGB(251,226,153) '연주황 
			Case "5" 
				frm1.vspdData.BackColor = RGB(255,255,153) '연노랑 
		End Select
	Next

End Sub

Sub SetQuerySpreadColor1()

	Dim iArrColor1, iArrColor2
	Dim iLoopCnt
	iArrColor1 = Split(lgStrColorFlag1,Parent.gRowSep)
	
	For iLoopCnt=0 to ubound(iArrColor1,1) - 1
		iArrColor2 = Split(iArrColor1(iLoopCnt),Parent.gColSep)

		frm1.vspdData1.Col = -1
		frm1.vspdData1.Row =  iArrColor2(0)
		
		Select Case iArrColor2(1)
			Case "1"
				frm1.vspdData1.BackColor = RGB(204,255,153) '연두 
			Case "2"
				frm1.vspdData1.BackColor = RGB(176,234,244) '하늘색 
			Case "3"
				frm1.vspdData1.BackColor = RGB(224,206,244) '연보라 
			Case "4"  
				frm1.vspdData1.BackColor = RGB(251,226,153) '연주황 
			Case "5" 
				frm1.vspdData1.BackColor = RGB(255,255,153) '연노랑 
		End Select
	Next
End Sub

'========================================
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case pvIntWhere
	
	'사업장 
	Case C_PopBizArea									
		iArrParam(1) = "B_BIZ_AREA"									' TABLE 명칭 
		iArrParam(2) = Trim(frm1.txtConBizAreaCd.value)				' Code Condition
		iArrParam(3) = ""											' Name Cindition
		iArrParam(4) = ""											' Where Condition
		iArrParam(5) = frm1.txtConBizAreaCd.alt						' TextBox 명칭 
		
		iArrField(0) = "ED15" & Parent.gColSep & "BIZ_AREA_CD"		' Field명(0)
		iArrField(1) = "ED30" & Parent.gColSep & "BIZ_AREA_NM"		' Field명(1)
    
	    iArrHeader(0) = frm1.txtConBizAreaCd.alt					' Header명(0)
	    iArrHeader(1) = frm1.txtConBizAreaNm.alt					' Header명(1)

		frm1.txtConBizAreaCd.focus 
	'영업그룹 
	Case C_PopSalesGrp	
		iArrParam(1) = "B_SALES_GRP"
		iArrParam(2) = Trim(frm1.txtConSalesGrpCd.value)
		iArrParam(3) = ""
		iArrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & " "
		iArrParam(5) = frm1.txtConSalesGrpCd.alt
		
		iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"
		iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"
    
	    iArrHeader(0) = frm1.txtConSalesGrpCd.alt
	    iArrHeader(1) = frm1.txtConSalesGrpNm.alt

		frm1.txtConSalesGrpCd.focus 
	'발행처 
	Case C_PopBillToParty
		iArrParam(1) = "B_BIZ_PARTNER_FTN PF, B_BIZ_PARTNER PA"
		iArrParam(2) = Trim(frm1.txtConBillToPartyCd.value)
		iArrParam(3) = ""
		iArrParam(4) = "PF.USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND PF.PARTNER_FTN = " & FilterVar("SBI", "''", "S") & "" _
						& "AND PA.BP_CD = PF.PARTNER_BP_CD AND PA.BP_TYPE <= " & FilterVar("CS", "''", "S") & ""
		iArrParam(5) = frm1.txtConBillToPartyCd.alt
		
		iArrField(0) = "ED15" & Parent.gColSep & "PA.BP_CD"
		iArrField(1) = "ED30" & Parent.gColSep & "PA.BP_NM"
    
	    iArrHeader(0) = frm1.txtConBillToPartyCd.alt
	    iArrHeader(1) = frm1.txtConBillToPartyNm.alt

		frm1.txtConBillToPartyCd.focus
	'세금신고사업장 
	Case C_PopTaxBiz
		iArrParam(1) = "B_TAX_BIZ_AREA"
		iArrParam(2) = Trim(frm1.txtConTaxBizCd.value)
		iArrParam(3) = ""
		iArrParam(4) = ""
		iArrParam(5) = frm1.txtConTaxBizCd.alt
		
		iArrField(0) = "ED15" & Parent.gColSep & "TAX_BIZ_AREA_CD"
		iArrField(1) = "ED30" & Parent.gColSep & "TAX_BIZ_AREA_NM"
    
	    iArrHeader(0) = frm1.txtConTaxBizCd.alt
	    iArrHeader(1) = frm1.txtConTaxBizNm.alt

		frm1.txtConTaxBizCd.focus

	End Select
	
	iArrParam(0) = iArrParam(5)										' 팝업 명칭 
	
	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	

End Function

'========================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	With frm1
		Select Case pvIntWhere
		Case C_PopBizArea
			.txtConBizAreaCd.value = pvArrRet(0) 
			.txtConBizAreaNm.value = pvArrRet(1)   			
			
		Case C_PopSalesGrp
			.txtConSalesGrpCd.value = pvArrRet(0)
			.txtConSalesGrpNm.value = pvArrRet(1)

		Case C_PopBillToParty
			frm1.txtConBillToPartyCd.value = pvArrRet(0) 
			frm1.txtConBillToPartyNm.value = pvArrRet(1)  

		Case C_PopTaxBiz
			frm1.txtConTaxBizCd.value = pvArrRet(0) 
			frm1.txtConTaxBizNm.value = pvArrRet(1)  
			
		End Select
	End With

	SetConPopup = True		
	
End Function

'========================================
Sub PopZAdoConfigGrid()

  If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
     Exit Sub
  End If

  If UCase(Trim(gActiveSpdSheet.id)) = "A" Then
	Call OpenOrderBy("A")
  Else
	Call OpenOrderBy("B")
  End If
  
End Sub

'========================================
Sub OpenOrderBy(ByVal pvPsdNo)
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Sub
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(pvPsdNo),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then												' Means that nothing is happened!!!
	   Exit Sub
	Else
	   Call ggoSpread.SaveXMLData(pvPsdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Sub

'========================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
    
    If Frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
End Sub

'========================================
Sub vspdData1_Click( Col,  Row)
	Call SetPopupMenuItemInf("00000000001")
	Set gActiveSpdSheet = frm1.vspdData1
	gMouseClickStatus = "SP2C"
	ggoSpread.Source    = frm1.vspdData1
	Call SetSpreadColumnValue("B",Frm1.vspdData1, Col, Row)
End Sub

'========================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub

'========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'========================================
Sub vspdData1_MouseDown(Button , Shift , x , y)

	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub    

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DbQuery
    	End If
    End If
    
End Sub

'========================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    If Row <> NewRow And NewRow > 0 Then
		frm1.vspdData1.MaxRows = 0
		Call DbQuery2(NewRow)
    End If
End Sub

'========================================
Sub txtConFromDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConFromDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConFromDt.Focus
	End If
End Sub

'========================================
Sub txtConToDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConToDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConToDt.Focus
	End If
End Sub

'========================================
Sub txtConFromDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================
Sub txtConToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출채권조회(판매유형2)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="30" align=right></td>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>									
									<TD CLASS="TD5" NOWRAP>매출채권일</TD>									
									<TD CLASS="TD6" NOWRAP>							        
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
											<script language =javascript src='./js/sd512qa3_OBJECT1_txtConFromDt.js'></script>
											</TD>
											<TD>
											&nbsp;~&nbsp;
											</TD>
											<TD>
											<script language =javascript src='./js/sd512qa3_OBJECT2_txtConToDt.js'></script>
											</TD>
										</TR>
									</TABLE>							        
							        </TD>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConBizArea" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopBizArea) ">
															<INPUT TYPE=TEXT NAME="txtConBizAreaNm" SIZE=20 tag="14" ALT="사업장명"></TD>
	                            </TR>		                            
								<TR>
									<TD CLASS="TD5" NOWRAP>영업그룹</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConSalesGrpCd" SIZE=10 MAXLENGTH=5 tag="11NXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSalesGrp) ">
															<INPUT TYPE=TEXT NAME="txtConSalesGrpNm" SIZE=20 tag="14" ALT="영업그룹명"></TD>
									<TD CLASS="TD5" NOWRAP>발행처</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConBillToPartyCd" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="발행처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConBillToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopBillToParty) ">
															<INPUT TYPE=TEXT NAME="txtConBillToPartyNm" SIZE=20 tag="14" ALT="발행처명"></TD>
								</TR>								
								<TR>						
									<TD CLASS="TD5" NOWRAP>세금신고사업장</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConTaxBizCd" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="세금신고사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConTaxBizCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopTaxBiz)  ">
															<INPUT TYPE=TEXT NAME="txtConTaxBizNm" SIZE=20 tag="14" ALT="세금신고사업장명"></TD>
									<TD CLASS=TD5 NOWRAP>확정여부</TD>
									<TD CLASS=TD6 NOWRAP colspan = 3>
										<input type=radio CLASS="RADIO" name="rdoBillConfFlag" id="rdoAll" value="A" tag = "11X" checked>
											<label for="rdoAll">전체</label>&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoBillConfFlag" id="rdoConf" value="S" tag = "11X">
											<label for="rdoConf">확정</label>&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoBillConfFlag" id="rdoNonConf" value="D" tag = "11X">
											<label for="rdoNonConf">미확정</label>
									</TD>								
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>예외여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoExceptFlag" id="rdoAll1" value="A" tag = "11X" checked>
											<label for="rdoAll1">전체</label>&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoExceptFlag" id="rdoExcept" value="Y" tag = "11X">
											<label for="rdoExcept">예외</label>&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoExceptFlag" id="rdoNormal" value="N" tag = "11X">
											<label for="rdoNormal">정상</label>
									</TD>
									<TD CLASS=TD5 NOWRAP>&nbsp;</TD>
									<TD CLASS=TD6 NOWRAP>&nbsp;</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="50%">
								<TD WIDTH="100%">
									<script language =javascript src='./js/sd512qa3_A_vspdData.js'></script>
								</TD>
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="100%">
									<script language =javascript src='./js/sd512qa3_B_vspdData1.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
    <TR>
      <TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>

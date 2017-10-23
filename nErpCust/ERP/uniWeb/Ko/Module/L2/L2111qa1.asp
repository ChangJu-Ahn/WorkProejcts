<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : 주문진행현황(I/F)
'*  5. Program Desc         : 주문진행현황(I/F)
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/04/01
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     :
'* 10. Modifier (Last)      :
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<Script Language="VBScript">
Option Explicit                             

Const BIZ_PGM_ID 		= "L2111QB1.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 22				                          '☆: SpreadSheet의 키의 갯수 

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgIsOpenPop                                          
Dim IsOpenPop  
Dim lgCookValue 
Dim lgSaveRow 
Dim lgStrFirstDt
Dim lgStrBaseDt

lgStrBaseDt = "<%=GetSvrDate%>"
lgStrFirstDt = UNIConvDateAToB(UNIGetFirstDay(lgStrBaseDt,parent.gServerDateFormat), parent.gServerDateFormat,parent.gDateFormat)
lgStrBaseDt	 = UNIConvDateAToB(lgStrBaseDt, parent.gServerDateFormat,parent.gDateFormat)

Dim lgStartRow
Dim lgEndRow

Const C_PopSoldToParty	= 1
Const C_PopPoNo	=	2

'========================================================================================================	
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                      'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
End Sub

'========================================================================================================
Sub SetDefaultVal()
	With frm1
		.txtConFromDt.Text = lgStrFirstDt
		.txtConToDt.Text   = lgStrBaseDt
		.txtConFromDt.focus
	End With
End Sub							
	
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	<% Call loadInfTB19029A("Q", "S","NOCOOKIE","QA") %>                                 
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal

	Const CookieSplit = 4877						

	If Kubun = 0 Then                                              ' Called Area
       strTemp = ReadCookie(CookieSplit)

       If strTemp = "" then Exit Function

       arrVal = Split(strTemp, parent.gRowSep)

       Frm1.txtSchoolCd.Value = ReadCookie ("SchoolCd")
       Frm1.txtGrade.Value   = arrVal(0)
				
       Call MainQuery()

       WriteCookie CookieSplit , ""
	
	ElseIf Kubun = 1 then                                         ' If you want to call
		Call vspdData_Click(Frm1.vspdData.ActiveCol,Frm1.vspdData.ActiveRow)
		WriteCookie CookieSplit , lgCookValue		
		Call PgmJump(BIZ_PGM_JUMP_ID2)
	End IF
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
		
End Function

'========================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("L2111QA101","S","A", "V20030401", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
    Call SetSpreadLock()   
End Sub

'========================================================================================================
Sub SetSpreadLock()

    With frm1
		.vspdData.ReDraw = False
		'------ Developer Coding part (Start ) -------------------------------------------------------------- 		
		ggoSpread.SpreadLock 1 , -1		
		'------ Developer Coding part (End   ) -------------------------------------------------------------- 
		.vspdData.ReDraw = True
    End With
    
End Sub

Sub SetSpreadColColor()
    With frm1.vspdData
		.Row = -1
		.Col = GetKeyPos("A",15)		:		.BackColor = RGB(204,255,153) '연두 
		.Col = GetKeyPos("A",16)		:		.BackColor = RGB(204,255,153) '연두 
		.Col = GetKeyPos("A",17)		:		.BackColor = RGB(204,255,153) '연두 
		.Col = GetKeyPos("A",18)		:		.BackColor = RGB(204,255,153) '연두 
    End With
    
End Sub

'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")

	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolBar("1100000000001111")										
    Call CookiePage(0)
    
    Call ggoOper.FormatDate(frm1.txtConYMFromDt, Parent.gDateFormat, 2)			'YYYYMM으로 포멧팅 
    Call ggoOper.FormatDate(frm1.txtConYMToDt, Parent.gDateFormat, 2)
    
    frm1.txtConFromDt.focus
    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
Function FncQuery() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncQuery = False                                                              
    
    Call ggoOper.ClearField(Document, "2")									      '⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.ClearSpreadData()

	'주문일 
	If ValidDateCheck(frm1.txtConFromDt, frm1.txtConToDt) = False Then Exit Function

	'접수일 
	If ValidDateCheck(frm1.txtConRcptFromDt, frm1.txtConRcptToDt) = False Then Exit Function
    
    Call InitVariables 														      '⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then								              '⊙: This function check indispensable field
       Exit Function
    End If

    If DbQuery = False Then 
       Exit Function
    End If   

    If Err.number = 0 Then
       FncQuery = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement  
    
End Function

'========================================================================================================
Function FncNew()

    Dim IntRetCD 
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNew = False	                                                              '☜: Processing is NG
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncNew = True                                                              '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   
    
End Function
	
	
'========================================================================================================
Function FncDelete()

    Dim intRetCD
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDelete = False                                                             '☜: Processing is NG
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
       FncDelete = True                                                           '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
Function FncSave()
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncSave = False                                                               '☜: Processing is NG
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncSave = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
Function FncCopy()

	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncCopy = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
Function FncCancel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCancel = False                                                             '☜: Processing is NG

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncCancel = True                                                           '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function



'========================================================================================================
Function FncInsertRow()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncInsertRow = False                                                          '☜: Processing is NG

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncInsertRow = True                                                        '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
Function FncDeleteRow()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDeleteRow = False                                                          '☜: Processing is NG

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncDeleteRow = True                                                        '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
Function FncPrint()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                        

    If Err.number = 0 Then
       FncPrint = True                                                            '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
Function FncPrev() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrev = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then
       FncPrev = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
Function FncNext() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNext = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then
       FncNext = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
Function FncExcel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call Parent.FncExport(parent.C_MULTI)
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then
       FncExcel = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
Function FncFind() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call Parent.FncFind(parent.C_MULTI, True)
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then
       FncFind = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'========================================================================================================
Function FncExit()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then
       FncExit = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
Function DbQuery() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbQuery = False
    
	Call LayerShowHide(1)
    
    With frm1
        If lgIntFlgMode  <> parent.OPMD_UMODE Then									'☜: This means that it is first search
			
			.txtHConFromDt.value	= .txtConFromDt.text
			.txtHConToDt.value		= .txtConToDt.text	
			.txtHConSoldToPartyCd.value	= .txtConSoldToPartyCd.value
			.txtHConPoNo.value			= .txtConPoNo.value
			.txtHConRcptFromDt.value	= .txtConRcptFromDt.text
			.txtHConRcptToDt.value		= .txtConRcptToDt.text
			
			.txtHlgSelectListDT.value	= GetSQLSelectListDataType("A") 
			.txtHlgTailList.value		= MakeSQLGroupOrderByList("A")
			.txtHlgSelectList.value		= EnCoding(GetSQLSelectList("A"))
        End If    
        
        .txtHlgPageNo.value	= lgPageNo
        
        lgStartRow = .vspdData.MaxRows + 1										'포멧팅 적용하는 시작Row
        
    End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    If Err.number = 0 Then
       DbQuery = True																'⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
Function DbQueryOk()												

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE										  '⊙: Indicates that current mode is Update mode
    lgSaveRow        = 1
    
    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================================
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	With frm1
		Select Case pvIntWhere
			'주문처 
			Case C_PopSoldToParty											
				iArrParam(1) = "B_BIZ_PARTNER"								
				iArrParam(2) = Trim(.txtConSoldToPartyCd.value)			
				iArrParam(3) = ""										
				iArrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND BP_TYPE IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"	
				iArrParam(5) = .txtConSoldToPartyCd.alt					
				
				iArrField(0) = "ED15" & Parent.gColSep & "BP_CD"		
				iArrField(1) = "ED30" & Parent.gColSep & "BP_NM"		
	    
			    iArrHeader(0) = .txtConSoldToPartyCd.alt			
			    iArrHeader(1) = .txtConSoldToPartyNm.alt				

				.txtConSoldToPartyCd.focus
			
			' 주문번호 
			Case C_PopPoNo
				iArrParam(1) = "S_INF_SO_HDR ISH INNER JOIN B_BIZ_PARTNER SP ON (SP.BP_CD = ISH.SOLD_TO_PARTY)"
				iArrParam(2) = Trim(.txtConPoNo.value)
				iArrParam(3) = ""
				iArrParam(4) = "ISH.DOC_ISSUE_DT >=  " & FilterVar(UNIConvDate(.txtConFromDt.Text), "''", "S") & ""

				If Trim(.txtConToDt.Text) <> "" Then
					iArrParam(4) = iArrParam(4) & " AND ISH.DOC_ISSUE_DT <=  " & FilterVar(UNIConvDate(.txtConToDt.Text), "''", "S") & ""
				End If
				
				If Trim(.txtConSoldToPartyCd.value) <> "" Then
					iArrParam(4) = iArrParam(4) & " AND ISH.SOLD_TO_PARTY =  " & FilterVar(.txtConSoldToPartyCd.value, "''", "S") & ""
				End If
				
				iArrParam(5) = .txtConPoNo.alt
				
				iArrField(0) = "ED20" & Parent.gColSep & "ISH.DOC_NO"		
				iArrField(1) = "DD15" & Parent.gColSep & "ISH.DOC_ISSUE_DT"
				iArrField(2) = "ED15" & Parent.gColSep & "ISH.SOLD_TO_PARTY"
				iArrField(3) = "ED20" & Parent.gColSep & "SP.BP_NM"
	    
			    iArrHeader(0) = .txtConPoNo.alt					' 주문번호 
			    iArrHeader(1) = "주문일"					' 주문일 
			    iArrHeader(2) = .txtConSoldToPartyCd.alt		' 주문처 
			    iArrHeader(3) = .txtConSoldToPartyNm.alt		' 주문처명 

				.txtConPoNo.focus
		End Select
	End With
	
	iArrParam(0) = iArrParam(5)										<%' 팝업 명칭 %> 
	
	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	

End Function

'========================================================================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	With frm1
		Select Case pvIntWhere
			Case C_PopSoldToParty
				.txtConSoldToPartyCd.value = pvArrRet(0) 
				.txtConSoldToPartyNm.value = pvArrRet(1)   
				
			Case C_PopPoNo
				.txtConPoNo.value = pvArrRet(0) 
		End Select
	End With

	SetConPopup = True		
	
End Function


'==================================================================================
Sub PopZAdoConfigGrid()

  If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
     Exit Sub
  End If

  Call OpenOrderBy("A")
  
End Sub


'========================================================================================================
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

'========================================================================================================
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
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)		
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
End Sub

'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub    

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
		'If CheckRunningBizProcess = True Then Exit Sub	
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DBQuery
    	End If
    End If    
End Sub


'========================================================================================================
Sub txtConFromDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConFromDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConFromDt.Focus
	End If
End Sub


'========================================================================================================
Sub txtConToDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConToDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConToDt.Focus
	End If
End Sub


'========================================================================================================
Sub txtConRcptFromDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConRcptFromDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConRcptFromDt.Focus
	End If
End Sub

'========================================================================================================
Sub txtConRcptToDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConRcptToDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConRcptToDt.Focus
	End If
End Sub

'========================================================================================================
Sub txtConFromDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================================================================================
Sub txtConToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================================================================================
Sub txtConRcptFromDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================================================================================
Sub txtConRcptToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>주문진행현황(I/F)</font></td>
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
									<TD CLASS="TD5" NOWRAP>주문일</TD>									
									<TD CLASS="TD6" NOWRAP>							        
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
											<script language =javascript src='./js/l2111qa1_OBJECT1_txtConFromDt.js'></script>
											</TD>
											<TD>
											&nbsp;~&nbsp;
											</TD>
											<TD>
											<script language =javascript src='./js/l2111qa1_OBJECT2_txtConToDt.js'></script>
											</TD>
										</TR>
									</TABLE>							        
							        </TD>
									<TD CLASS="TD5" NOWRAP>주문처</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConSoldToPartyCd" SIZE=10 MAXLENGTH=5 tag="11NXXU" ALT="주문처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSoldToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSoldToParty) ">
															<INPUT TYPE=TEXT NAME="txtConSoldToPartyNm" SIZE=20 tag="14" ALT="주문처명"></TD>
	                            </TR>	
	                            <TR>
									<TD CLASS="TD5" NOWRAP>주문번호</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConPoNo" SIZE=20 MAXLENGTH=18 tag="11NXXU" ALT="주문번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConBizArea" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopPoNo) "></TD>
									<TD CLASS="TD5" NOWRAP>접수일</TD>									
									<TD CLASS="TD6" NOWRAP>							        
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
											<script language =javascript src='./js/l2111qa1_OBJECT1_txtConRcptFromDt.js'></script>
											</TD>
											<TD>
											&nbsp;~&nbsp;
											</TD>
											<TD>
											<script language =javascript src='./js/l2111qa1_OBJECT2_txtConRcptToDt.js'></script>
											</TD>
										</TR>
									</TABLE>							        
							        </TD>
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
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/l2111qa1_vspdData_vspdData.js'></script>
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

<INPUT TYPE=HIDDEN NAME="txtHConFromDt"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConToDt"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConSoldToPartyCd"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConPoNo"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConRcptFromDt"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConRcptToDt"		tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHlgPageNo"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgSelectListDT"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgTailList"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgSelectList"		tag="24" TABINDEX="-1">				

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>

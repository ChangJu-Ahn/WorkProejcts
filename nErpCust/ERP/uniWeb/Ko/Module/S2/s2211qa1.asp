<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2211QA1
'*  4. Program Name         : 판매계획확정정보조회 
'*  5. Program Desc         : 판매계획확정정보조회 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2001/04/18
'*  8. Modified date(Last)  : 2001/04/18
'*  9. Modifier (First)     :
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
	
Const BIZ_PGM_ID 		= "S2211QB1.asp"                              '☆: Biz Logic ASP Name

Const C_MaxKey          = 3					                          '☆: SpreadSheet의 키의 갯수 

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgIsOpenPop                                          
Dim IsOpenPop  

Dim lgCookValue 

Dim lgSaveRow 
<%
Dim lsSvrDate
lsSvrDate = GetsvrDate
%>

'========================================================================================================	
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0

End Sub

'========================================================================================================
Sub SetDefaultVal()

	frm1.cboConSpType.focus
	Call cboConFlag_onChange
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("Q", "S","NOCOOKIE","QA") %>                                '☆: 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub InitSpreadSheet(ByVal pvPsdNo)
    
    if pvPsdNo = "A" then
		Call SetZAdoSpreadSheet("S2211QA101","S","A", "V20030115", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
		Call SetSpreadLock("A")
	else
		Call SetZAdoSpreadSheet("S2211QA102","S","B", "V20030115", parent.C_SORT_DBAGENT,frm1.vspdData2,C_MaxKey, "X", "X")
		Call SetSpreadLock("B")
	end if
	
End Sub

'========================================================================================================
Sub SetSpreadLock(Byval iOpt)
    If iOpt = "A" Then
       With frm1
          .vspdData.ReDraw = False
          ggoSpread.Source = .vspdData 
          ggoSpread.SpreadLockWithOddEvenRowColor()
          .vspdData.ReDraw = True
       End With
    Else
       With frm1
            .vspdData2.ReDraw = False
            ggoSpread.Source = .vspdData2
            ggoSpread.SpreadLock 1, -1
            .vspdData2.ReDraw = True
       End With
    End If   
   
End Sub

'========================================================================================================
Sub Form_Load()

    Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")

	Call InitVariables	
	Call InitComboBox	
	Call SetDefaultVal	
	Call InitSpreadSheet("A")
	Call InitSpreadSheet("B")
    Call SetToolBar("1100000000001111")	
    call ChangeingcboConFlag									
    
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================================================================================
Function FncQuery() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncQuery = False                                                              '⊙: Processing is NG

	
    Call ggoOper.ClearField(Document, "2")									      '⊙: Clear Contents  Field
    
    if frm1.cboConFlag.value = "G" then
		ggoSpread.Source = frm1.vspdData
	else
		ggoSpread.Source = frm1.vspdData2
	end if
    Call ggoSpread.ClearSpreadData()
    
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
Function FncPrint()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                        '☜: Protect system from crashing

    If Err.number = 0 Then
       FncPrint = True                                                            '⊙: Processing is OK
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

	Dim strVal, iSheetNo

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbQuery = False
    
    Call LayerShowHide(1)
    call ChangeingcboConFlag
     
	if frm1.cboConFlag.value = "G" then
		iSheetNo = "A"
	else
		iSheetNo = "B"
	end if
	  
    With frm1

        strVal = BIZ_PGM_ID
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode  <> parent.OPMD_UMODE Then   ' This means that it is first search
           strVal = strVal & "?cboConStep="       & Trim(.cboConStep.value)
           strVal = strVal & "&cboConSpType="     & Trim(.cboConSpType.Value)
           strVal = strVal & "&txtConSalesGrp="   & Trim(.txtConSalesGrp.Value)
           strVal = strVal & "&txtConSpPeriod="   & Trim(.txtConSpPeriod.Value)
        Else
           strVal = strVal & "?cboConStep="       & Trim(.hcboConStep.value)
           strVal = strVal & "&cboConSpType="     & Trim(.hcboConSpType.Value)
           strVal = strVal & "&txtConSalesGrp="   & Trim(.htxtConSalesGrp.Value)
           strVal = strVal & "&txtConSpPeriod="   & Trim(.htxtConSpPeriod.Value)
        End If   

    '--------- Developer Coding Part (End) ------------------------------------------------------------

        strVal = strVal & "&lgPageNo="       & lgPageNo                          '☜: Next key tag
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType(iSheetNo)
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList(iSheetNo)
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList(iSheetNo))
		strVal = strVal & "&cboConFlag="     & Trim(.cboConFlag.Value)
		
        Call RunMyBizASP(MyBizASP, strVal)	                                         '☜: 비지니스 ASP 를 가동 
    
    End With

    If Err.number = 0 Then
       DbQuery = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================================================================
Function DbQueryOk()												

    On Error Resume Next															'☜: If process fails
    Err.Clear																		'☜: Clear error status
    
	frm1.vspdData.ReDraw = false 

	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE												 '⊙: Indicates that current mode is Update mode
    lgSaveRow        = 1

    Set gActiveElement = document.ActiveElement   
    
    if frm1.cboConFlag.value = "G" then
		Set Spread = frm1.vspdData
		frm1.vspdData.focus
	else		
		Set Spread = frm1.vspdData2
		frm1.vspdData2.focus
	end if

	frm1.vspdData.ReDraw = True
	Call SetToolBar("1100000000011111")	
	 
End Function

'========================================================================================================
Sub InitComboBox()
	' 판매계획유형 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("S0023", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboConSpType,lgF0,lgF1,parent.gColSep)

	'진행단계 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("S0021", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboConstep,lgF0,lgF1,Chr(11))
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("F1007", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	'조회기준 
	lgF0 = "G" & Chr(11) & "P" & Chr(11)
	lgF1 = "영업그룹" & Chr(11) & "공장" & Chr(11)
	Call SetCombo2(frm1.cboConFlag,lgF0,lgF1,Chr(11))
	
End Sub

'========================================================================================================
'	Description : 조회기준 onChange 이벤트 처리 
'========================================================================================================
Sub cboConFlag_onChange()
	select case frm1.cboConFlag.value
		Case "G"
			lblTitle.innerHTML = "영업그룹"
		Case Else
			lblTitle.innerHTML = "공장"
	End Select
	frm1.txtConSalesGrp.value= ""
	frm1.txtConSalesGrpNm.value=""
	frm1.txtConSpPeriod.value= ""
	frm1.txtConSpPeriodDesc.value=""
		
End Sub

'========================================================================================================
'	Description : 조회기준 Change시 vspddata설정 
'========================================================================================================
Sub ChangeingcboConFlag()

lgBlnFlgChgValue = true	
	select case frm1.cboConFlag.value
		Case "G"
			frm1.vspdData.style.display = "inline"
			frm1.vspdData2.style.display = "none"
			ggoSpread.Source = frm1.vspdData
		Case Else
			frm1.vspdData.style.display = "none"
			frm1.vspdData2.style.display = "inline"
			ggoSpread.Source = frm1.vspdData2
	End Select
	
End Sub

'========================================================================================================
Function OpenConPopup()

	Dim iarrRet
	Dim iArrParam(6), iArrField(6), iArrHeader(6)
	Dim pvIntWhere

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	pvIntWhere= frm1.cboConFlag.value 
	
	Select Case pvIntWhere
	Case "G"
			iArrParam(1) = "B_SALES_GRP "			<%' TABLE 명칭 %>
			iArrParam(2) = frm1.txtConSalesGrp.value						<%' Code Condition%>
			iArrParam(3) = ""								<%' Name Cindition%>
			iArrParam(4) = ""							<%' Where Condition%>
			iArrParam(5) = "영업그룹"									<%' TextBox 명칭 %>
			
			iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"
			iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"	
			
			
			iArrHeader(0) = "영업그룹"
			iArrHeader(1) = "영업그룹명"
			
			frm1.txtConSalesGrp.focus
	Case "P"
			iArrParam(1) = "B_PLANT "			<%' TABLE 명칭 %>
			iArrParam(2) = frm1.txtConSalesGrp.value						<%' Code Condition%>
			iArrParam(3) = ""								<%' Name Cindition%>
			iArrParam(4) = ""							<%' Where Condition%>
			iArrParam(5) = "공장"						<%' TextBox 명칭 %>
				
			iArrField(0) = "ED15" & Parent.gColSep & "PLANT_CD"
			iArrField(1) = "ED30" & Parent.gColSep & "PLANT_NM"
			
			iArrHeader(0) = "공장"
			iArrHeader(1) = "공장명"
		
			frm1.txtConSalesGrp.focus
			
	End Select
    
    iArrParam(0) = iArrParam(5)
    
	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	
	IsOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPopup(iArrRet)
		OpenConPopup = True		
	End If	


End Function

'========================================================================================================
Function SetConPopup(Byval iArrRet)

	SetConPopup = False

	frm1.txtConSalesGrp.Value = iArrRet(0)
	frm1.txtConSalesGrpNm.Value = iArrRet(1)		

	SetConPopup = True

End Function


'==================================================================================
Sub PopZAdoConfigGrid()

	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
	   Exit Sub
	End If

	Select Case UCase(Trim(gActiveSpdSheet.Name))
	 
		Case "VSPDDATA"
			OpenOrderBy("A")
		Case "VSPDDATA2"			
				OpenOrderBy("B")
	End Select	
  
End Sub

'========================================================================================================
' Sales planning period Popup
Function OpenConSpPeriodPopup(byval pvStrData)

Dim iArrRet
Dim iArrParam(2)
Dim iCalledAspName

	OpenConSpPeriodPopup = False

	iCalledAspName = AskPRAspName("s2211pa3")
	
	If Trim(iCalledAspName) = "" Then
		Call DisplayMsgBox("900040", parent.VB_INFORMATION, "s2211pa3", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	frm1.txtConSpPeriod.focus 	
	iArrParam(0) = pvStrData
	
	iArrRet = window.showModalDialog(iCalledAspName & "?txtDisplayFlag=N", Array(window.parent,iArrParam), _
	 "dialogWidth=690px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtConSpPeriod.Value = iArrRet(0)
		frm1.txtConSpPeriodDesc.Value = iArrRet(1)	
		
	End If	
	
End Function

'========================================================================================================
Sub OpenOrderBy(ByVal pvPsdNo)

Dim arrRet
	
	If lgIsOpenPop = True Then Exit Sub
	lgIsOpenPop = True
	
	If pvPsdNo = "A" then
		ggoSpread.Source = Frm1.vspdData
	else
		ggoSpread.Source = Frm1.vspdData2
	end if

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(pvPsdNo),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then         ' Means that nothing is happened!!!
	   Exit Sub
	Else
	   Call ggoSpread.SaveXMLData(pvPsdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet(pvPsdNo)       
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
Sub vspdData2_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("00000000001")
    
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData2
    
    If Frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData2
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
    Call SetSpreadColumnValue("B",frm1.vspdData2,Col,Row)		
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
End Sub

'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub

'========================================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
    
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub

'========================================================================================================
Sub vspdData_GotFocus()

    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
End Sub

'========================================================================================================
Sub vspdData2_GotFocus()

    ggoSpread.Source = Frm1.vspdData2

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
Sub vspdData2_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub 

'========================================================================================================
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
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
    
End Sub

'========================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If Frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData2,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>판매계획확정정보조회</font></td>
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
									<TD CLASS="TD5" NOWRAP>판매계획유형</TD>
									<TD CLASS="TD6"><SELECT Name="cboConSpType" ALT="판매계획유형" tag="12XXXU"></SELECT></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6"></TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>조회기준</TD>
									<TD CLASS="TD6"><SELECT Name="cboConFlag" ALT="조회기준" STYLE="WIDTH: 150px" tag="12"></SELECT></TD> 
														   
									<TD CLASS="TD5" NOWRAP>진행단계</TD>
									<TD CLASS="TD6"><SELECT Name="cboConStep" ALT="진행단계" STYLE="WIDTH: 200px" tag="1XXXXU"><OPTION VALUE="" selected></OPTION></SELECT></TD> 
											
							         </TD>
	                            </TR>	
	                            <TR>
									<TD CLASS="TD5" id = "lblTitle" NOWRAP>영업그룹</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtConSalesGrp" SIZE=10 MAXLENGTH=10 tag="1XXXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup() ">
														   <INPUT TYPE=TEXT NAME="txtConSalesGrpNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>계획기간</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtConSpPeriod" SIZE=10 MAXLENGTH=10 tag="1XXXXU" ALT="계획기간"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSpPeriodPopup(frm1.txtConSpPeriod.Value) ">
														   <INPUT TYPE=TEXT NAME="txtConSpPeriodDesc" SIZE=20 tag="14"></TD>
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
									<script language =javascript src='./js/s2211qa1_vspdData_vspdData.js'></script>
									<script language =javascript src='./js/s2211qa1_vspdData2_vspdData2.js'></script>
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
<INPUT TYPE=HIDDEN NAME="hcboConStep"     tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hcboConSpType"   tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtConSalesGrp" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="htxtConSpPeriod" tag="24" TABINDEX="-1">

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>

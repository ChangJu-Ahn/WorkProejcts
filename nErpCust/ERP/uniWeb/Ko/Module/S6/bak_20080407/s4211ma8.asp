<%@ LANGUAGE="VBSCRIPT" %>
<%
'************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : S4211MA8
'*  4. Program Name         : 통관현황조회 
'*  5. Program Desc         : 통관현황조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/11
'*  8. Modified date(Last)  : 2002/04/12
'*  9. Modifier (First)     : Cho Sung-Hyun
'* 10. Modifier (Last)      : Park in sik
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/29 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'*                            -2002/04/11 : ADO변환 
'**************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                              
<!-- #Include file="../../inc/lgvariables.inc" --> 

Dim lgIsOpenPop                                              
Dim lgMark                                                  
Dim IsOpenPop
Dim arrParam	

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "s4211mb8.asp"
Const BIZ_PGM_JUMP_ID	= "s4211ma1"
Const C_MaxKey          = 8                                    '☆☆☆☆: Max key value

'------ Minor Code PopUp을 위한 Major Code정의 ------ 
Const gstrEDTypeMajor = "S9012"
Const gstrExportTypeMajor = "S9009"
Const gstrEpTypesMajor = "S9008"
'========================================================================================================= 
Sub InitVariables()
	
	lgBlnFlgChgValue = False                               
    lgStrPrevKey     = ""                                  
    lgSortKey        = 1
	lgPageNo         = ""
    lgIntFlgMode = parent.OPMD_CMODE	

	
End Sub
'=========================================================================================================
Sub SetDefaultVal()
	
	frm1.txtApplicantCd.focus
	frm1.txtFromDate.text = StartDate
	frm1.txtToDate.text = EndDate

End Sub
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub
'==========================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S4211MA8","S","A","V20030901", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock 
      
End Sub
'=========================================================================================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub
'===========================================================================
Function OpenCCHdrRef()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	Dim IntRetCD			

	On Error Resume Next

	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End IF

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("S4211RA9")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S4211RA9", "X")			
		IsOpenPop = False
		Exit Function
	End If
				
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",1)		
		arrParam(0) = frm1.vspdData.Text
		
		frm1.vspdData.Col = GetKeyPos("A",2)
		arrParam(1) = frm1.vspdData.Text

	IsOpenPop = True
  
	arrRet = window.showModalDialog(iCalledAspName ,Array(window.parent,arrParam),"dialogWidth=840px; dialogHeight=481px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function
'===========================================================================
Function OpenCCDtlRef()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD			

	On Error Resume Next

	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End IF

	If IsOpenPop = True Then Exit Function		

	iCalledAspName = AskPRAspName("S4212RA6")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S4212RA6", "X")			
		IsOpenPop = False
		Exit Function
	End If
		
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",1)
		arrParam(0) = frm1.vspdData.Text
		frm1.vspdData.Col = GetKeyPos("A",2)
		arrParam(1) = frm1.vspdData.Text
		frm1.vspdData.Col = GetKeyPos("A",7)
		arrParam(2) = frm1.vspdData.Text
		frm1.vspdData.Col = GetKeyPos("A",8)
		arrParam(3) = frm1.vspdData.Text
		frm1.vspdData.Col = GetKeyPos("A",3)
		arrParam(4) = frm1.vspdData.Text
		frm1.vspdData.Col = GetKeyPos("A",4)
		arrParam(5) = frm1.vspdData.Text	
	IsOpenPop = True
		   
	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent,arrParam),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function
'===========================================================================
Function OpenCCLanRef()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD			

	On Error Resume Next

	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End IF

	If IsOpenPop = True Then Exit Function			

	iCalledAspName = AskPRAspName("S4213RA8")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S4213RA8", "X")			
		IsOpenPop = False
		Exit Function
	End If

		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",1)
		arrParam(0) = frm1.vspdData.Text
		frm1.vspdData.Col = GetKeyPos("A",2)
		arrParam(1) = frm1.vspdData.Text
		frm1.vspdData.Col = GetKeyPos("A",3)
		arrParam(2) = frm1.vspdData.Text
		frm1.vspdData.Col = GetKeyPos("A",5)
		arrParam(3) = frm1.vspdData.Text

		arrParam(4) = parent.gCurrency
		frm1.vspdData.Col = GetKeyPos("A",6)
		arrParam(5) = frm1.vspdData.Text
	
	IsOpenPop = True
		   
	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent,arrParam),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPopup(ByVal iType)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iType
	Case 1		
		arrParam(0) = "수입자"						
		arrParam(1) = "B_BIZ_PARTNER"							
		arrParam(2) = Trim(frm1.txtApplicantCd.value)		
		arrParam(3) = Trim(frm1.txtApplicantNm.value)	
		arrParam(4) = "BP_TYPE <= " & FilterVar("CS", "''", "S") & ""							
		arrParam(5) = "수입자"							
			
		arrField(0) = "BP_CD"									
		arrField(1) = "BP_NM"									
			
		arrHeader(0) = "수입자"							
		arrHeader(1) = "수입자명"						

	Case 2		
		arrParam(0) = "영업그룹"
		arrParam(1) = "B_SALES_GRP"						
		arrParam(2) = Trim(frm1.txtSalesGrpCd.value)		
		arrParam(3) = Trim(frm1.txtSalesGrpNm.value)	
		arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "					
		arrParam(5) = "영업그룹"					
		
	    arrField(0) = "SALES_GRP"						
	    arrField(1) = "SALES_GRP_NM"					
	    
	    arrHeader(0) = "영업그룹"					
	    arrHeader(1) = "영업그룹명"	
	    
	Case 3		
	
		arrParam(0) = "신고구분"					
		arrParam(1) = "B_Minor"							
		arrParam(2) = Trim(frm1.txtEdType.value)		
		arrParam(3) = Trim(frm1.txtEdTypeNm.value)		
		arrParam(4) = "MAJOR_CD= " & FilterVar(gstrEDTypeMajor, "''", "S") & ""	
		arrParam(5) = "신고구분"						
		
		arrField(0) = "Minor_CD"							
		arrField(1) = "Minor_NM"							

		arrHeader(0) = "신고구분"						
		arrHeader(1) = "신고구분명"						
		
	Case 4	
	
		arrParam(0) = "수출거래구분"					
		arrParam(1) = "B_Minor"								
		arrParam(2) = Trim(frm1.txtEpType.value)			
		arrParam(3) = Trim(frm1.txtEpTypeNm.value)			
		arrParam(4) = "MAJOR_CD= " & FilterVar(gstrEpTypesMajor, "''", "S") & ""	
		arrParam(5) = "수출거래구분"					
		
		arrField(0) = "Minor_CD"						
		arrField(1) = "Minor_NM"						

		arrHeader(0) = "수출거래구분"				
		arrHeader(1) = "수출거래구분명"				
		
	Case 5	
	
		arrParam(0) = "수출구분"					
		arrParam(1) = "B_Minor"							
		arrParam(2) = Trim(frm1.txtExportType.value)	
		arrParam(3) = Trim(frm1.txtExportTypeNm.value)	
		arrParam(4) = "MAJOR_CD= " & FilterVar(gstrExportTypeMajor, "''", "S") & ""	
		arrParam(5) = "수출구분"					
		
		arrField(0) = "Minor_CD"						
		arrField(1) = "Minor_NM"						

		arrHeader(0) = "수출구분"					
		arrHeader(1) = "수출구분명"					
	End Select
    
	arrParam(3) = ""
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	Select Case iType
	    Case 1
	    	frm1.txtApplicantCd.focus
	    Case 2
	    	frm1.txtSalesGrpCd.focus
	    case 3
	    	frm1.txtEdType.focus
	    case 4
	    	frm1.txtEpType.focus
	    case 5
	    	frm1.txtExportType.focus
	    case else
	End Select	

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPopup(arrRet,iType)
	End If	
	
End Function
'========================================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderByPopup("A")
End Sub
'========================================================================================================
Sub OpenOrderByPopup(ByVal pSpdNo)
	Dim arrRet
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Sub
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Sub
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Sub
'--------------------------------------------------------------------------------------------------------- 
Function SetConPopup(Byval arrRet,ByVal iType)

	Select Case iType
	Case 1
		frm1.txtApplicantCd.value = arrRet(0) 
		frm1.txtApplicantNm.value = arrRet(1)
		frm1.txtApplicantCd.focus
	Case 2
		frm1.txtSalesGrpCd.value = arrRet(0) 
		frm1.txtSalesGrpNm.value = arrRet(1)   
		frm1.txtSalesGrpCd.focus
	case 3
		frm1.txtEdType.value = arrRet(0) 
		frm1.txtEdTypeNm.value = arrRet(1) 
		frm1.txtEdType.focus
	case 4
		frm1.txtEpType.value = arrRet(0) 
		frm1.txtEpTypeNm.value = arrRet(1) 
		frm1.txtEpType.focus
	case 5
		frm1.txtExportType.value = arrRet(0) 
		frm1.txtExportTypeNm.value = arrRet(1) 
		frm1.txtExportType.focus
	case else
	End Select

End Function
'--------------------------------------------------------------------------------------------------------- 
Sub CookiePage(Byval Kubun)
	Const CookieSplit = 4877
	If Kubun = 1 Then		
			frm1.vspdData.row = frm1.vspdData.ActiveRow
			frm1.vspdData.Col = GetKeyPos("A",1)		
			WriteCookie CookieSplit, frm1.vspdData.Text
	End IF
	
End Sub

'=========================================================================================================
Sub Form_Load()
    
	Call LoadInfTB19029	
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   
	
	 '----------  Coding part  -------------------------------------------------------------
	Call InitVariables			    
	Call SetDefaultVal	
	Call InitSpreadSheet()

    Call SetToolbar("11000000000011")							'⊙: 버튼 툴바 제어 
    
    frm1.txtApplicantCd.focus
    
End Sub
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    lgBlnFlgChgValue = True
    
End Sub
'==========================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("00000000001")
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData	
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	
    If Row <= 0 Then
       
       ggoSpread.Source = frm1.vspdData
       If lgSortKey = 1 Then
			ggoSpread.SSSort Col		'Sort in ascending
			lgSortKey = 2
	   Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in descending
			lgSortKey = 1
       End If
       
       Exit Sub
    End If   
    
    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)

End Sub
'==========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
End Sub
'==========================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then
      	Exit Sub
    End If
    If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	  
    	If lgPageNo <> "" Then                   
			If DbQuery = False Then
				Exit Sub
			End if
    	End If
	End If
    
End Sub
'=======================================================================================================
Sub txtFromDate_DblClick(Button)  
	If Button = 1 Then
		frm1.txtFromDate.Action = 7 
		Call SetFocusToDocument("M")
        frm1.txtFromDate.Focus
	End If
End Sub
'=======================================================================================================
Sub txtToDate_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDate.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtToDate.Focus
	End If
End Sub
'==========================================================================================
Sub txtFromDate_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
Sub txtToDate_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'********************************************************************************************************* 
Function FncQuery() 

	Dim IntRetCD

    FncQuery = False                                                        
    
    Err.Clear                                                               

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
       						
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.ClearSpreadData

    Call InitVariables 														
    
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
	
	If ValidDateCheck(frm1.txtFromDate, frm1.txtToDate) = False Then Exit Function

    Call DbQuery																'☜: Query db data

    FncQuery = True																'⊙: Processing is OK

End Function
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function
'========================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                            
End Function
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    FncExit = True
End Function
'========================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim StrNextKey      

    DbQuery = False
    
    Err.Clear                                                  

			
	If   LayerShowHide(1) = False Then
         Exit Function 
    End If

	Dim strVal
    
    With frm1

		 If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
			strVal = strVal & "&txtApplicantCd=" & Trim(.HApplicantCd.value)
			strVal = strVal & "&txtSalesGrpCd=" & Trim(.HSalesGrpCd.value)
			strVal = strVal & "&txtEdType=" & Trim(.HEdType.value)
			strVal = strVal & "&txtFromDate=" & Trim(.HFromDate.value)
			strVal = strVal & "&txtToDate=" & Trim(.HToDate.value)
	    Else
		    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
			strVal = strVal & "&txtApplicantCd=" & Trim(.txtApplicantCd.value)
			strVal = strVal & "&txtSalesGrpCd=" & Trim(.txtSalesGrpCd.value)
			strVal = strVal & "&txtEdType=" & Trim(.txtEdType.value)
			strVal = strVal & "&txtFromDate=" & Trim(.txtFromDate.Text)
			strVal = strVal & "&txtToDate=" & Trim(.txtToDate.Text)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
       End if   
			<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------%>	
			strVal = strVal & "&lgPageNo="       & lgPageNo                			
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	
	Call RunMyBizASP(MyBizASP, strVal)										
        
    End With
    
    DbQuery = True

End Function
'========================================================================================
Function DbQueryOk()														

	lgIntFlgMode = parent.OPMD_UMODE	

    Call SetToolbar("11000000000111")							'⊙: 버튼 툴바 제어 
    
    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    Else
       frm1.txtApplicantCd.focus	
    End if      
	
	lgBlnFlgChgValue = False
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>>&nbsp;<% ' 상위 여백 %></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>통관현황조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenCCHdrRef">통관상세정보</A> | <A href="vbscript:OpenCCDtlRef">통관내역정보</A> | <A href="vbscript:OpenCCLanRef">통관란정보</A></TD>					
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
					<FIELDSET CLASS="CLSFLD"><TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
							<TD CLASS=TD5 NOWRAP>수입자</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtApplicantCd" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMLCBp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup 1">&nbsp;<INPUT NAME="txtApplicantNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
							<TD CLASS=TD5 NOWRAP>영업그룹</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrpCd"  TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMLCSaleGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup 2">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>신고구분</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEdType" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMLCBp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup 3">&nbsp;<INPUT NAME="txtEdTypeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
							<TD CLASS=TD5 NOWRAP>작성일</TD>
							<TD CLASS=TD6 NOWRAP>
								<script language =javascript src='./js/s4211ma8_fpDateTime1_txtFromDate.js'></script>
								&nbsp;~&nbsp;
								<script language =javascript src='./js/s4211ma8_fpDateTime2_txtToDate.js'></script>
							</TD>
						</TR>
				</TABLE></TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
			<TR>
				<TD WIDTH=100% HEIGHT=* valign=top><TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD HEIGHT="100%">
							<script language =javascript src='./js/s4211ma8_I741745114_vspdData.js'></script>
						</TD>
					</TR></TABLE>
				</TD>
			</TR>
		</TABLE></TD>
	</TR>
	<TR>
		<td <%=HEIGHT_TYPE_01%>></td>
	</TR>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
			<TR>
				<TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">통관등록</a></TD>
				<TD WIDTH="50"></TD>
			</TR></TABLE>
      </TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">

<INPUT TYPE=HIDDEN NAME="HApplicantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="HSalesGrpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="HEdType" tag="24">
<INPUT TYPE=HIDDEN NAME="HEpType" tag="24">
<INPUT TYPE=HIDDEN NAME="HExportType" tag="24">
<INPUT TYPE=HIDDEN NAME="HFromDate" tag="24">
<INPUT TYPE=HIDDEN NAME="HToDate" tag="24">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1" ></iframe>
</DIV>

</BODY>
</HTML>


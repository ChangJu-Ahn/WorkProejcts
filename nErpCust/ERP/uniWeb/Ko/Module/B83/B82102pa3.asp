
<%@ LANGUAGE="VBSCRIPT" %>
<%'******************************************************************************************************
'*  1. Module Name			: 접수검토사항         												        *
'*  2. Function Name		: 																			*
'*  3. Program ID			:                        										*
'*  4. Program Name			: Reference Popup GI for Order List											*
'*  5. Program Desc			: Reference Popup															*
'*  7. Modified date(First)	: 																*
'*  8. Modified date(Last)	: 																*
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)		:																			*	
'* 11. Comment 		:																					*
'********************************************************************************************************%>

<HTML>
<HEAD>
<!--'####################################################################################################
'#						1. 선 언 부																		#
'#####################################################################################################-->

<!--'********************************************  1.1 Inc 선언  ****************************************
'*	Description : Inc. Include																			*
'*****************************************************************************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<!--'============================================  1.1.1 Style Sheet  ===================================
'=====================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--'============================================  1.1.2 공통 Include  ==================================
'=====================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script LANGUAGE="VBScript">

Option Explicit

<!-- #Include file="../../inc/lgVariables.inc" -->

Dim IsOpenPop 
Dim arrReturn
Dim arrParent
Dim arrParam                         
Dim arrField
Dim PopupParent
                    
arrParent = window.dialogArguments

Set PopupParent = arrParent(0)

arrParam = arrParent(1)
arrField = arrParent(2)

Dim StartDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)

top.document.title = PopupParent.gActivePRAspName

'========================================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()
 
End Sub

'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)				=
'========================================================================================================

Function InitVariables()

    Redim arrReturn(0)
     Self.Returnvalue   = arrReturn
     
End Function
	
'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Description : 화면 초기화(수량 Field나 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)		=
'========================================================================================================
Sub SetDefaultVal()
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	If arrParam(0) = "" OR arrParam(0) = "1900-01-01" Then
	   frm1.txtDt.Text = StartDate
	Else
	   frm1.txtDt.Text = arrParam(0)
	End If
    frm1.txtGrade.value    = arrParam(1)
    frm1.txtDesc.value     = Replace(arrParam(2) , chr(7), chr(13)&chr(10))
    frm1.txtPerSon.value   = arrParam(3)
    frm1.txtPerSonNm.value = arrParam(4)
    
    '품질에서 접수했으면 막음...        
    If arrParam(7) <> "" OR (arrParam(8) = "E" OR arrParam(8) = "S" OR arrParam(8) = "T" OR arrParam(8) = "D") Then
       frm1.btnRun1.disabled = True
       If arrParam(9) = "X" OR arrParam(8) = "T" Then
          frm1.btnRun2.disabled = True
       End If   
       Call ggoOper.SetReqAttr(frm1.txtDt, "Q")
       Call ggoOper.SetReqAttr(frm1.txtGrade, "Q")
       Call ggoOper.SetReqAttr(frm1.txtDesc, "Q")
       Call ggoOper.SetReqAttr(frm1.txtPerSon, "Q")
       Call ggoOper.SetReqAttr(frm1.txtPerSonNm, "Q")
    Else
       If arrParam(8) = "R" OR arrParam(1) = "" Then
          frm1.btnRun2.disabled = True 
       End If   
    End If
       
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "P", "NOCOOKIE", "RA")%>
End Sub

'========================================================================================================
' Name : InitComboBox()     
' Desc : Initialize combo value
'========================================================================================================
Sub InitComboBox()
     If arrParam(10) = "N" Then
        Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'Y1008' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
     ElseIf arrParam(10) = "C" Then
        Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'Y1007' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
     End If   
     Call SetCombo2(frm1.txtGrade, lgF0, lgF1, Chr(11))
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream()
   '------ Developer Coding part (Start ) --------------------------------------------------------------                    
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    
End Sub
'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    
End Sub
'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    with frm1
    ggoSpread.Source = gActiveSpdSheet
    frm1.vspdData.Redraw = False
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()
	Call ggoSpread.ReOrderingSpreadData()
	end with 
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.3 개발자 정의 함수  ++++++++++++++++++++++++++++++++++++++
'+	개발자 정의 Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : OkClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function OkClick()

    Redim arrReturn(UBound(arrField))
    
    If Not chkField(Document, "2") Then                                          '☜: Check contents area
       Exit Function
    End If
    
	arrReturn(0) = frm1.txtDt.Text
    arrReturn(1) = frm1.txtGrade.value
    arrReturn(2) = Replace(frm1.txtDesc.value , chr(13)&chr(10) , chr(7))
    arrReturn(3) = frm1.txtPerSon.value
    arrReturn(4) = frm1.txtPerSonNm.value
    
	    if len(arrReturn(2)) > 100 then
	 Call  popupparent.DisplayMsgBox("127928", vbOKOnly, "100글자", "X") '초과할 수 없습니다. 
    else
		Self.Returnvalue = arrReturn
		Self.Close()
    end if
    
End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()

    If DisplayMsgBox("900018", PopUpParent.VB_YES_NO,"x","x") = vbNo Then
       Exit Function
    End If
    
    Redim arrReturn(UBound(arrField))
        
	arrReturn(0) = ""
    arrReturn(1) = ""
    arrReturn(2) = ""
    arrReturn(3) = ""
    arrReturn(4) = ""
    
	Self.Returnvalue = arrReturn
	Self.Close()
	
End Function

'=========================================  2.3.3 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CloseClick()

    Redim arrReturn(0)
    
    Self.Returnvalue = arrReturn
    Self.Close()
	
End Function

'=========================================  2.3.3 Mouse Pointer 처리 함수 ===============================
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
            case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분				=
'========================================================================================================

Sub Form_Load()
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call MM_preloadImages("../../CShared/image/Query.gif","../../CShared/image/OK.gif","../../CShared/image/Cancel.gif")
    Call InitVariables
    Call ggoOper.LockField(Document, "N")                       '⊙: Lock  Suitable  Field
    Call InitComboBox()    
    Call SetDefaultVal() 
End Sub

'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'======================================================================================================
'        Name : OpenPopup()
'        Description : 
'=======================================================================================================
Function OpenPopup(Byval arPopUp)

        Dim arrRet
        Dim Param(7), Field(8), Header(8)

        If IsOpenPop = True  Then  
           Exit Function
        End If   

        IsOpenPop = True
        
        Select Case arPopUp
               Case 1 '
                    Param(0) = frm1.txtPerSon.Alt
                    Param(1) = "B_CIS_ROUTING_USER A, Z_USR_MAST_REC B"
                    Param(2) = Trim(frm1.txtPerSon.value)
                    Param(4) = "A.USER_ID = B.USR_ID AND A.ITEM_ACCT = '" & arrParam(5) & "' AND A.ITEM_KIND = '" & arrParam(6) & "' AND A.ITEM_R = 'Y'"
                    Param(5) = frm1.txtPerSon.Alt

                    Field(0) = "A.USER_ID"
                    Field(1) = "B.USR_NM"
    
                    Header(0) = frm1.txtPerSon.Alt
                    Header(1) = frm1.txtPerSonNm.Alt
               frm1.txtPerSon.focus()
               Case Else
                    IsOpenPop = False
                    Exit Function
      End Select
        
      arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(Param, Field, Header), _
                "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

      IsOpenPop = False
                
      If arrRet(0) = "" Then
         Exit Function
      Else
         Call SubSetPopup(arrRet,arPopUp)
      End If        
        
End Function

'======================================================================================================
'        Name : SubSetPopup()
'        Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetPopup(Byval arrRet, Byval arPopUp)

    lgBlnFlgChgValue = True
    
    With Frm1
        Select Case arPopUp
               Case 1 
                    .txtPerSon.value   = arrRet(0)
                    .txtPerSonNm.value = arrRet(1)
               Case Else
                    Exit Sub
              End Select              
              
        End With
End Sub

Sub txtDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtDt.Action = 7
		Frm1.txtDt.Focus
	End If
End Sub

'========================================================================================
' Function Name : txtPerSon_OnChange
' Function Desc : 
'========================================================================================
Function txtPerSon_OnChange()
    Dim iDx
    Dim IntRetCd
 
    If frm1.txtPerSon.value = "" Then
        frm1.txtPerSonnm.value = ""
    ELSE    
		IntRetCd =  CommonQueryRs(" USR_NM "," Z_USR_MAST_REC "," USR_id="&filterVar(frm1.txtPerSon.value,"''","S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)      
		  If IntRetCd = false Then
			 frm1.txtPerSonnm.value=""
        Else
            frm1.txtPerSonnm.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<%
'########################################################################################################
'#						6. Tag 부																		#
'########################################################################################################
%>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR>
		<TD HEIGHT=40>
			<FIELDSET CLASS="CLSFLD">
				<TABLE WIDTH=100% CELLSPACING=0>					
					<TR>
					   <TD CLASS=TD5 NOWRAP>검토자</TD>
                       <TD CLASS=TD6 ColSpan=3><INPUT NAME="txtPerSon" ALT="검토자" TYPE="Text" SiZE=10 MAXLENGTH=10   tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPumpType" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenPopup('1')"> 
                                               <INPUT NAME="txtPerSonNm" ALT="검토자명" TYPE="Text" SiZE=25   tag="24XXXU"></TD>					   
					</TR> 
					<TR>
					   <TD CLASS=TD5 NOWRAP>검토결과</TD>
					   <TD CLASS=TD6 NOWRAP><SELECT NAME="txtGrade"  CLASS=cboNormal TAG="22" ALT="검토결과"><OPTION VALUE=""></OPTION></SELECT></TD>
					</TR> 
					<TR>   
					   <TD CLASS=TD5 NOWRAP>검토일자</TD>
				       <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/b82102pa3_fpDateTime1_txtDt.js'></script></TD>				            				                
                    </TR> 					
					<TR>	
						<TD CLASS=TD6 HEIGHT="40%" NOWRAP COLSPAN =4 ><B>*검토내역</B></TD>
					</TR>
					<TR>	
						<TD CLASS=TD6 ColSpan=3><TEXTAREA  NAME="txtDesc" tag="22XXXU" rows=4 cols=60  ALT="검토내역"></TEXTAREA>
                        			            <INPUT TYPE=HIDDEN NAME="htxtDesc" SIZE=800 tag="X4" TABINDEX=-1>
						</TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR><TD HEIGHT=10>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=10>&nbsp;</TD>
				<TD WIDTH=70% NOWRAP></TD>
				<TD WIDTH=30% ALIGN=RIGHT><BUTTON NAME="btnRun1" CLASS="CLSSBTN" ONCLICK="OkClick()">확인</BUTTON>&nbsp;
				                          <BUTTON NAME="btnRun2" CLASS="CLSSBTN" ONCLICK="CancelClick()">취소</BUTTON>&nbsp;
										  <BUTTON NAME="btnRun3" CLASS="CLSSBTN" ONCLICK="CloseClick()">종료</BUTTON>
				<TD WIDTH=10>&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</FORM>
</BODY>
</HTML>


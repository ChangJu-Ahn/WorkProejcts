<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Common
*  2. Function Name        : Common
*  3. Program ID           : ADO group sort popup
*  4. Program Name         : ADO group sort popup
*  5. Program Desc         : ADO group sort popup
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2001/04/18
*  9. Modifier (First)     :
* 10. Modifier (Last)      :
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../inc/IncServer.asp" -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/AdoQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Ccm.vbs">      </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Common.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Event.vbs">    </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../inc/IncImage.js">  </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

'============================================  1.2.2 Global 변수 선언  ==================================
'========================================================================================================
    Dim iCount
    Dim jCount
    Dim rBuffer
	Dim arrReturn					<% '--- Return Parameter Group %>
	Dim arrParent
	Dim arrParam
	Dim arrTitle
	Dim lgSortFieldCD1	
	Dim lgSortTitleNm
	Dim lgWarningMessage
    Dim iMessage 
		
    arrParent = window.dialogArguments

    lgSortFieldCD1  = arrParent(0)
    lgSortTitleNm   = arrParent(1)
	arrParam        = arrParent(2)
    arrTitle        = arrParent(3)
	     
    top.document.title = arrTitle(0)
    
    lgWarningMessage = "이미 선택된 값입니다. 다시 선택하세요"
	iMessage         = "최소한 한개는 선택 하십시요"
	
'==========================================  2.1.1 InitVariables()  =====================================
'=	Name : InitVariables()																				=
'=	Desc : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)			        =
'========================================================================================================
Function InitVariables()
	Redim arrReturn(C_MaxSelList * 2)
	arrReturn(0) = 0 
    Self.Returnvalue = arrReturn
		
    lblTitle.innerHTML = arrTitle(0) & "순서"
 
End Function
'==========================================  2.2.1 SetDefaultVal()  =====================================
'=	Name : SetDefaultVal()																				=
'=	Desc : 화면 초기화(수량 Field나 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)	            =
'========================================================================================================
Sub SetDefaultVal()
    Dim ii,jj,tt
	    
    Frm1.cboOrderBy1.length = UBound(lgSortFieldCD1) + 1
    Frm1.cboOrderBy2.length = UBound(lgSortFieldCD1) + 1
    Frm1.cboOrderBy3.length = UBound(lgSortFieldCD1) + 1
    Frm1.cboOrderBy4.length = UBound(lgSortFieldCD1) + 1
    Frm1.cboOrderBy5.length = UBound(lgSortFieldCD1) + 1
    Frm1.cboOrderBy6.length = UBound(lgSortFieldCD1) + 1
         
    Frm1.cboOrderBy1.options(0).value = ""
    Frm1.cboOrderBy1.options(0).text  = ""
         
    For ii = 0 to UBound(lgSortFieldCD1) - 1
        Frm1.cboOrderBy1.options(ii+1).value = lgSortFieldCD1(ii)
        Frm1.cboOrderBy1.options(ii+1).text  = lgSortTitleNm (ii)
    Next    
         
         frm1.cboOrderBy2.options(0).value = ""
         frm1.cboOrderBy2.options(0).text  = ""
         
         For ii = 0 to UBound(lgSortFieldCD1) - 1
             frm1.cboOrderBy2.options(ii+1).value = lgSortFieldCD1(ii)
             frm1.cboOrderBy2.options(ii+1).text  = lgSortTitleNm (ii)
	     Next    
	     
         frm1.cboOrderBy3.options(0).value = ""
         frm1.cboOrderBy3.options(0).text  = ""
         
         For ii = 0 to UBound(lgSortFieldCD1) - 1
             frm1.cboOrderBy3.options(ii+1).value = lgSortFieldCD1(ii)
             frm1.cboOrderBy3.options(ii+1).text  = lgSortTitleNm (ii)
	     Next    

         frm1.cboOrderBy4.options(0).value = ""
         frm1.cboOrderBy4.options(0).text  = ""
         
         For ii = 0 to UBound(lgSortFieldCD1) - 1
             frm1.cboOrderBy4.options(ii+1).value = lgSortFieldCD1(ii)
             frm1.cboOrderBy4.options(ii+1).text  = lgSortTitleNm (ii)
	     Next    
	     
         frm1.cboOrderBy5.options(0).value = ""
         frm1.cboOrderBy5.options(0).text  = ""
         
         For ii = 0 to UBound(lgSortFieldCD1) - 1
             frm1.cboOrderBy5.options(ii+1).value = lgSortFieldCD1(ii)
             frm1.cboOrderBy5.options(ii+1).text  = lgSortTitleNm (ii)
	     Next    
	     
         frm1.cboOrderBy6.options(0).value = ""
         frm1.cboOrderBy6.options(0).text  = ""
         
         For ii = 0 to UBound(lgSortFieldCD1) - 1
             frm1.cboOrderBy6.options(ii+1).value = lgSortFieldCD1(ii)
             frm1.cboOrderBy6.options(ii+1).text  = lgSortTitleNm (ii)
	     Next    
End Sub
	
'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================
Function OKClick()

		Dim GroupCol
		Dim ChkArr
		Dim iTemp
		Dim iDx
		Dim iDx1

		Redim arrReturn(C_MaxSelList * 2)
		Redim ChkArr(5)

		ChkArr(0) = Trim(Frm1.cboOrderBy1.value)
		ChkArr(1) = Trim(Frm1.cboOrderBy2.value)
		ChkArr(2) = Trim(Frm1.cboOrderBy3.value)
		ChkArr(3) = Trim(Frm1.cboOrderBy4.value)
		ChkArr(4) = Trim(Frm1.cboOrderBy5.value)
		ChkArr(5) = Trim(Frm1.cboOrderBy6.value)
		
		For iDx = 0 To UBound(ChkArr)
            If ChkArr(iDx) > "" Then
               iDx = UBound(ChkArr) + 2
               Exit For
            End If 
        Next    
        
        If iDx < UBound(ChkArr) + 2 Then
   	       Msgbox iMessage, vbExclamation, gLogoName & "-[Warning]"                 'Check If Selected list
		   Exit Function
		End If   


		For iDx = 0 To UBound(ChkArr)
		    iTemp = Trim(ChkArr(iDx))
		    If iTemp > "" Then
               For iDx1 = iDx + 1 To UBound(ChkArr)
                  If iTemp = ChkArr(iDx1) Then
                     Msgbox lgWarningMessage, vbExclamation, gLogoName & "-[Warning]"
                     Select Case iDx1
                        Case 1 
                              Frm1.cboOrderBy2.selectedIndex = 0
                              Frm1.cboOrderBy2.focus
                        Case 2 
                              Frm1.cboOrderBy3.selectedIndex = 0
                              Frm1.cboOrderBy3.focus
                        Case 3 
                              Frm1.cboOrderBy4.selectedIndex = 0
                              Frm1.cboOrderBy4.focus
                        Case 4 
                              Frm1.cboOrderBy5.selectedIndex = 0
                              Frm1.cboOrderBy5.focus
                        Case 5 
                              Frm1.cboOrderBy6.selectedIndex = 0
                              Frm1.cboOrderBy6.focus
                     End Select          
                     Exit Function
                  End If 
               Next    
            End If
        Next    		
		
		GroupCol = 0

		If Len(Trim(frm1.cboOrderBy1.value)) Then 
			GroupCol = GroupCol + 1
			arrReturn(GroupCol) = frm1.cboOrderBy1.value
			GroupCol = GroupCol + 1
			if frm1.rdoSortMethod1_A.checked = true then
				arrReturn(GroupCol) = frm1.rdoSortMethod1_A.value
			else
				arrReturn(GroupCol) = frm1.rdoSortMethod1_D.value
			end if
		End If		

		If Len(Trim(frm1.cboOrderBy2.value)) Then 
			GroupCol = GroupCol + 1
			arrReturn(GroupCol) = frm1.cboOrderBy2.value
			GroupCol = GroupCol + 1
			if frm1.rdoSortMethod2_A.checked = true then
				arrReturn(GroupCol) = frm1.rdoSortMethod2_A.value
			else
				arrReturn(GroupCol) = frm1.rdoSortMethod2_D.value
			end if
		End If

		If Len(Trim(frm1.cboOrderBy3.value)) Then 
			GroupCol = GroupCol + 1
			arrReturn(GroupCol) = frm1.cboOrderBy3.value		
			GroupCol = GroupCol + 1
			if frm1.rdoSortMethod3_A.checked = true then
				arrReturn(GroupCol) = frm1.rdoSortMethod3_A.value
			else
				arrReturn(GroupCol) = frm1.rdoSortMethod3_D.value
			end if
		End If

		If Len(Trim(frm1.cboOrderBy4.value)) Then 
			GroupCol = GroupCol + 1
			arrReturn(GroupCol) = frm1.cboOrderBy4.value		
			GroupCol = GroupCol + 1
			if frm1.rdoSortMethod4_A.checked = true then
				arrReturn(GroupCol) = frm1.rdoSortMethod4_A.value
			else
				arrReturn(GroupCol) = frm1.rdoSortMethod4_D.value
			end if
		End If


		If Len(Trim(frm1.cboOrderBy5.value)) Then 
			GroupCol = GroupCol + 1
			arrReturn(GroupCol) = frm1.cboOrderBy5.value		
			GroupCol = GroupCol + 1
			if frm1.rdoSortMethod5_A.checked = true then
				arrReturn(GroupCol) = frm1.rdoSortMethod5_A.value
			else
				arrReturn(GroupCol) = frm1.rdoSortMethod5_D.value
			end if
		End If

		If Len(Trim(frm1.cboOrderBy6.value)) Then 
			GroupCol = GroupCol + 1
			arrReturn(GroupCol) = frm1.cboOrderBy6.value		
			GroupCol = GroupCol + 1
			if frm1.rdoSortMethod6_A.checked = true then
				arrReturn(GroupCol) = frm1.rdoSortMethod6_A.value
			else
				arrReturn(GroupCol) = frm1.rdoSortMethod6_D.value
			end if
		End If


		arrReturn(0) = GroupCol

		Self.Returnvalue = arrReturn

		Self.Close()

	End Function
	
	
'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
	Function CancelClick()

		Self.Close()
	End Function
	
'=========================================  3.1.1 Form_Load()  ==========================================
'=	Name : Form_Load()																					=
'=	Description : Window Load시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분				=
'========================================================================================================
	Sub Form_Load()		

		Call SetDefaultVal()
		Call InitVariables
		Call MM_preloadImages("../image/Query.gif","../image/OK.gif","../image/Cancel.gif")
		DBQuery()
	End Sub
'=========================================  3.1.2 Form_QueryUnload()  ===================================
'   Event Name : Form_QueryUnload																		=
'   Event Desc :																						=
'========================================================================================================
	Sub Form_QueryUnload(Cancel, UnloadMode)

	End Sub

Function Document_onKeyUp()

	Dim objEl, KeyCode

	Set objEl = window.event.srcElement
	
	KeyCode = window.event.keycode

	If KeyCode = 27 Then
       Call CancelClick()
    ElseIf KeyCode = 13 Then
         Select Case  Left(objEl.getAttribute("tag"),1)
             Case "1"
                  Call cboOrderBy1_OnClick()  
             Case "2"
                  Call cboOrderBy2_OnClick()  
             Case "3"
                  Call cboOrderBy3_OnClick()  
             Case "4"
                  Call cboOrderBy4_OnClick()  
             Case "5"
                  Call cboOrderBy5_OnClick()  
             Case "6"
                  Call cboOrderBy6_OnClick()  
         End Select                  
    End If
End Function

'==========================================================================================================
' Function Name : cboSum1_OnClick
' Function Desc : Click시 동일한값이 있는지 판별 
'==========================================================================================================
Sub cboOrderBy1_OnClick()
	If Trim(Frm1.cboOrderBy1.value) = "" Then Exit Sub

	Select Case frm1.cboOrderBy1.value
	Case frm1.cboOrderBy2.value,frm1.cboOrderBy3.value,frm1.cboOrderBy4.value,frm1.cboOrderBy5.value,frm1.cboOrderBy6.value
        Msgbox lgWarningMessage, vbExclamation, gLogoName & "-[Warning]"
		frm1.cboOrderBy1.selectedIndex = 0
		frm1.cboOrderBy1.focus
	End Select

End Sub

'==========================================================================================================
' Function Name : cboSum2_OnClick
' Function Desc : Click시 동일한값이 있는지 판별 
'==========================================================================================================
Sub cboOrderBy2_OnClick()
	If Trim(Frm1.cboOrderBy2.value) = "" Then Exit Sub

	Select Case Frm1.cboOrderBy2.value
	    Case Frm1.cboOrderBy1.value, Frm1.cboOrderBy3.value, Frm1.cboOrderBy4.value, Frm1.cboOrderBy5.value, Frm1.cboOrderBy6.value

             Msgbox lgWarningMessage, vbExclamation, gLogoName & "-[Warning]"
             Frm1.cboOrderBy2.selectedIndex = 0
             Frm1.cboOrderBy2.focus
	End Select

End Sub
'==========================================================================================================
' Function Name : cboSum3_OnClick
' Function Desc : Click시 동일한값이 있는지 판별 
'==========================================================================================================
Sub cboOrderBy3_OnClick()
	If Trim(Frm1.cboOrderBy3.value) = "" Then Exit Sub

	Select Case frm1.cboOrderBy3.value
	Case frm1.cboOrderBy1.value,frm1.cboOrderBy2.value,frm1.cboOrderBy4.value,frm1.cboOrderBy5.value,frm1.cboOrderBy6.value
        Msgbox lgWarningMessage, vbExclamation, gLogoName & "-[Warning]"
		frm1.cboOrderBy3.selectedIndex = 0
		frm1.cboOrderBy3.focus
	End Select

End Sub

'==========================================================================================================
' Function Name : cboSum4_OnClick
' Function Desc : Click시 동일한값이 있는지 판별 
'==========================================================================================================
Sub cboOrderBy4_OnClick()
	If Trim(Frm1.cboOrderBy4.value) = "" Then Exit Sub

	Select Case frm1.cboOrderBy4.value
	Case frm1.cboOrderBy1.value,frm1.cboOrderBy2.value,frm1.cboOrderBy3.value,frm1.cboOrderBy5.value,frm1.cboOrderBy6.value
        Msgbox lgWarningMessage, vbExclamation, gLogoName & "-[Warning]"
		frm1.cboOrderBy4.selectedIndex = 0
		frm1.cboOrderBy4.focus
	End Select

End Sub

'==========================================================================================================
' Function Name : cboSum5_OnClick
' Function Desc : Click시 동일한값이 있는지 판별 
'==========================================================================================================
Sub cboOrderBy5_OnClick()
	If Trim(Frm1.cboOrderBy5.value) = "" Then Exit Sub

	Select Case frm1.cboOrderBy5.value
	Case frm1.cboOrderBy1.value,frm1.cboOrderBy2.value,frm1.cboOrderBy3.value,frm1.cboOrderBy4.value,frm1.cboOrderBy6.value
        Msgbox lgWarningMessage, vbExclamation, gLogoName & "-[Warning]"
		frm1.cboOrderBy5.selectedIndex = 0
		frm1.cboOrderBy5.focus
	End Select

End Sub
'*****************************************  3.3.3 cboSum3_OnClick()  **************************************
' Function Name : cboSum5_OnClick
' Function Desc : Click시 동일한값이 있는지 판별 
'**********************************************************************************************************
Sub cboOrderBy6_OnClick()
	If Trim(Frm1.cboOrderBy6.value) = "" Then Exit Sub

	Select Case frm1.cboOrderBy6.value
	Case frm1.cboOrderBy1.value,frm1.cboOrderBy2.value,frm1.cboOrderBy3.value,frm1.cboOrderBy4.value,frm1.cboOrderBy5.value
        Msgbox lgWarningMessage, vbExclamation, gLogoName & "-[Warning]"
		frm1.cboOrderBy6.selectedIndex = 0
		frm1.cboOrderBy6.focus
	End Select

End Sub

'########################################################################################################
'#						5. Interface 부																	#
'########################################################################################################
Function DBQuery()

		Err.Clear															<%'☜: Protect system from crashing%>

		DBQuery = False														<%'⊙: Processing is NG%>

        DBQueryOK 
		DBQuery = True	
End Function

'*******************************************  5.1 ADOQueryOk()  *******************************************
' Function Name : ADOQueryOk
' Function Desc : ADOQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'**********************************************************************************************************
Function DBQueryOK()

	If Trim(arrParam(0)) <> "" Then
		frm1.cboOrderBy1.value = arrParam(0)
		If Trim(arrParam(1)) = "1" then
			frm1.rdoSortMethod1_A.checked = true
		Else
			frm1.rdoSortMethod1_D.checked = true
		End If
	End If
	If Trim(arrParam(2)) <> "" Then
		frm1.cboOrderBy2.value = arrParam(2)
		If Trim(arrParam(3)) = "1" Then
			frm1.rdoSortMethod2_A.checked = True
		Else
			frm1.rdoSortMethod2_D.checked = True
		End If
	End if
	If Trim(arrParam(4)) <> "" Then
		frm1.cboOrderBy3.value = arrParam(4)
		If Trim(arrParam(5)) = "1" Then
			frm1.rdoSortMethod3_A.checked = True
		Else
			frm1.rdoSortMethod3_D.checked = True
		End If
	End if
	
	If Trim(arrParam(6)) <> "" Then
		frm1.cboOrderBy4.value = arrParam(6)
		If Trim(arrParam(7)) = "1" Then
			frm1.rdoSortMethod4_A.checked = True
		Else
			frm1.rdoSortMethod4_D.checked = True
		End If
	End if
	
	If Trim(arrParam(8)) <> "" Then
		frm1.cboOrderBy5.value = arrParam(8)
		If Trim(arrParam(9)) = "1" Then
			frm1.rdoSortMethod5_A.checked = True
		Else
			frm1.rdoSortMethod5_D.checked = True
		End If
	End if
	
	If Trim(arrParam(10)) <> "" Then
		frm1.cboOrderBy6.value = arrParam(10)
		If Trim(arrParam(11)) = "1" Then
			frm1.rdoSortMethod6_A.checked = True
		Else
			frm1.rdoSortMethod6_D.checked = True
		End If
	End if
	
End Function


</SCRIPT>
<!-- #Include file="../inc/UNI2KCMCom.inc" -->	
</HEAD>
<%'======================================================================================================
'#						6. Tag 부																		#
'=======================================================================================================%>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP">
<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD HEIGHT=5></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../image/table/seltab_up_bg.gif"><img src="../image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><SPAN CLASS="normal" ID="lblTitle">&nbsp;</SPAN></font></td>
								<td background="../image/table/seltab_up_bg.gif" align="right"><img src="../image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
	<TD WIDTH=100% CLASS="Tab11">
		<TABLE CLASS="BasicTB" CELLSPACING=0>
		<TR HEIGHT=100%>
			<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET>
				<TABLE WIDTH=100% CELLSPACING=0 border=0>
				<TR >
					<TD CLASS="TD5" NOWRAP ><LI>기준 1</LI></TD>
					<TD CLASS="TD6" NOWRAP><SELECT NAME="cboOrderBy1" STYLE="WIDTH: 110px" TAG="1"><OPTION selected></SELECT>
						<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoSortMethod1"  VALUE="1" CHECKED ID="rdoSortMethod1_A" ><LABEL FOR="rdoSortMethod1">오름차순</LABEL>
						<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoSortMethod1"  VALUE="2" ID="rdoSortMethod1_D"><LABEL FOR="rdoSortMethod1">내림차순</LABEL>
				</TR>
				<TR>
					<TD CLASS="TD5" NOWRAP><LI>기준 2</LI></TD>
					<TD CLASS="TD6" NOWRAP><SELECT NAME="cboOrderBy2" STYLE="WIDTH: 110px" TAG="2"><OPTION selected></SELECT>
						<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoSortMethod2"  VALUE="1" CHECKED ID="rdoSortMethod2_A" ><LABEL FOR="rdoSortMethod2">오름차순</LABEL>
						<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoSortMethod2"  VALUE="2" ID="rdoSortMethod2_D"><LABEL FOR="rdoSortMethod2">내림차순</LABEL>
					</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" NOWRAP><LI>기준 3</LI></TD>
					<TD CLASS="TD6" NOWRAP><SELECT NAME="cboOrderBy3" STYLE="WIDTH: 110px" TAG="3"><OPTION selected></SELECT>
						<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoSortMethod3"  VALUE="1" CHECKED ID="rdoSortMethod3_A" ><LABEL FOR="rdoSortMethod3">오름차순</LABEL>
						<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoSortMethod3"  VALUE="2" ID="rdoSortMethod3_D"><LABEL FOR="rdoSortMethod3">내림차순</LABEL>
					</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" NOWRAP><LI>기준 4</LI></TD>
					<TD CLASS="TD6" NOWRAP><SELECT NAME="cboOrderBy4" STYLE="WIDTH: 110px" TAG="4"><OPTION selected></SELECT>
						<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoSortMethod4"  VALUE="1" CHECKED ID="rdoSortMethod4_A" ><LABEL FOR="rdoSortMethod4">오름차순</LABEL>
						<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoSortMethod4"  VALUE="2" ID="rdoSortMethod4_D"><LABEL FOR="rdoSortMethod4">내림차순</LABEL>
					</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" NOWRAP><LI>기준 5</LI></TD>
					<TD CLASS="TD6" NOWRAP><SELECT NAME="cboOrderBy5" STYLE="WIDTH: 110px" TAG="5"><OPTION selected></SELECT>
						<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoSortMethod5"  VALUE="1" CHECKED ID="rdoSortMethod5_A" ><LABEL FOR="rdoSortMethod5">오름차순</LABEL>
						<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoSortMethod5"  VALUE="2" ID="rdoSortMethod5_D"><LABEL FOR="rdoSortMethod5">내림차순</LABEL>
					</TD>
				</TR>
				<TR>
					<TD CLASS="TD5" NOWRAP><LI>기준 6</LI></TD>
					<TD CLASS="TD6" NOWRAP><SELECT NAME="cboOrderBy6" STYLE="WIDTH: 110px" TAG="6"><OPTION selected></SELECT>
						<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoSortMethod6"  VALUE="1" CHECKED ID="rdoSortMethod6_A" ><LABEL FOR="rdoSortMethod6">오름차순</LABEL>
						<INPUT TYPE=RADIO CLASS=RADIO NAME="rdoSortMethod6"  VALUE="2" ID="rdoSortMethod6_D"><LABEL FOR="rdoSortMethod6">내림차순</LABEL>
					</TD>
				</TR>
				</TABLE>
			</FIELDSET>
			</TD>
		</TR>
		</TABLE>
	</TD>
	</TR>
	<TR><TD HEIGHT=20>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/OK.gif',1)"></IMG>&nbsp;&nbsp;
				<IMG SRC="../image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 framespacing=0 SCROLLING=no></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="PRG_ID" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
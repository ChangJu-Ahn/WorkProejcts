<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : Single Sample
*  3. Program ID           : H1a02ma1
*  4. Program Name         : H1a02ma1
*  5. Program Desc         : Single Sample
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2003/06/10
*  9. Modifier (First)     : Hwang Jeong Won
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
=======================================================================================================-->

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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "h2002mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_JUMP_ID = "H2001ma1" 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %> 
End Sub
'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
	On Error Resume Next

	Const CookieSplit = 4877						
	Dim strTemp

	If flgs = 1 Then
		WriteCookie CookieSplit , frm1.txtEmp_no.Value
	ElseIf flgs = 0 Then

		strTemp = ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function
			
		frm1.txtEmp_no.value =  strTemp

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		WriteCookie CookieSplit , ""
		
		Call MainQuery()
			
	End If
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   If pOpt = "Q" Then
      lgKeyStream = Frm1.txtEmp_No.Value & parent.gColSep       'You Must append one character(parent.gColSep)
   Else
      lgKeyStream = Frm1.txtEmp_No.Value & parent.gColSep         'You Must append one character(parent.gColSep)
   End If   
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : developer describe this line Initialize ComboBox
'========================================================================================================
Sub InitComboBox()
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
 		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec) 'condition
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
    Call ggoOper.FormatDate(frm1.txtEntr_dt, gDateFormat, 1)  

    Call FuncGetAuth(gStrRequestMenuID , parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
	Call SetToolbar("1110100000000111")							 					'⊙: Set ToolBar
	Call InitVariables
	
	frm1.txtEmp_No.focus

    Call InitComboBox
	Call CookiePage (0)                                                             '☜: Check Cookie
			
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    Dim RetStatus
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt

    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
    frm1.txtPath.value = ""
    If  txtEmp_no_Onchange() then
        Exit Function
    End If
    Call InitVariables                                                           '⊙: Initializes local global variables

    Call MakeKeyStream("Q")
    
	Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery = False Then
        Call RestoreToolBar()
        Exit Function
    End If
       
    FncQuery = True                                                              '☜: Processing is OK

End Function
	
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call ggoOper.LockField(Document , "N")                                       '☜: Lock  Field

    Frm1.ImgItemImage.src = ""

	Call SetToolbar("1110100000000111")												'⊙: Set ToolBar
    Call InitVariables                                                        '⊙: Initializes local global variables
    
    Set gActiveElement = document.ActiveElement   
    
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgBox("900002","x","x","x")                                
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"x","x")                        '☜: Do you want to delete? 
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    Call MakeKeyStream("D")
    
	Call DisableToolBar(parent.TBC_DELETE)
    If DbDelete = False Then
        Call RestoreToolBar()
        Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    
    FncDelete = True                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    dim file_end
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = False Then 
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
 
     If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If   
    If Not chkField(Document, "2") Then                                          '☜: Check contents area
       Exit Function
    End If

	if check_img_path(frm1.txtPath.value ) = false then
		IntRetCD = DisplayMsgBox("800488","x","x","x")                           '☜:There is no picture
		frm1.txtPath.focus
		exit function
	end if
   
    Call MakeKeyStream("S")

	Call DisableToolBar(parent.TBC_SAVE)
    If DbSave = False Then
        Call RestoreToolBar()
        Exit Function
    End If    
    
    FncSave = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : developer describe this line Called by MainSave in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
    FncCopy = True                                                            '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
	On Error Resume Next                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : developer describe this line Called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
	On Error Resume Next                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
	On Error Resume Next                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrev
' Desc : developer describe this line Called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 

    Dim strVal
    Dim IntRetCD

    FncPrev = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status
    
    FncPrev = True                                                               '☜: Processing is OK

End Function
'========================================================================================================
' Name : FncNext
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncNext() 
    Dim strVal
    Dim IntRetCD

    FncNext = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status
    
    FncNext = True                                                               '☜: Processing is OK
	
End Function
'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(parent.C_SINGLE, True)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

	frm1.txtPath.value = ""
	
	Call CommonQueryRs(" COUNT(*) "," HAA070T "," emp_no= " & FilterVar( Frm1.txtEmp_no.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

   	if   Replace(lgF0, Chr(11), "") > 0  then

		strVal = "../../ComASP/CPictRead.asp" & "?txtKeyValue=" & Frm1.txtEmp_No.value '☜: query key
		strVal = strVal     & "&txtDKeyValue=" & "default"                            '☜: default value
		strVal = strVal     & "&txtTable="     & "HAA070T"                            '☜: Table Name
		strVal = strVal     & "&txtField="     & "Photo"	                          '☜: Field
		strVal = strVal     & "&txtKey="       & "Emp_no"	                          '☜: Key
	else
		strVal = "../../../CShared/image/default_picture.jpg"
	end if
	
    Frm1.ImgItemImage.src = strVal
    call DbQueryOk()

    DbQuery = True                                                               '☜: Processing is NG

End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
		
	lgIntFlgMode = parent.OPMD_UMODE
	With Frm1
		.txtMode.value        = parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
		
    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0003                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '☜: Direction
		
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
    Frm1.txtEmp_no.focus
	Call SetToolbar("1111100000010111")
    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
	Frm1.txtPath.focus
End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Call InitVariables
    Call MainQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call MainNew()	
End Function

'========================================================================================================
' Name : OpenEmpName()
' Desc : developer describe this line 
'========================================================================================================
Function OpenEmpName(iWhere, strData)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	Select Case iWhere
	    Case 0
	        arrParam(0) = ""
	        arrParam(1) = frm1.txtName.value
	    Case 1
	        arrParam(0) = frm1.txtEmp_no.value
	        arrParam(1) = ""
	    Case 2
            arrParam(0) = strData
            arrParam(1) = ""
	End Select

    arrParam(2) = lgUsrIntCd

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	    if iWhere = 2 then
                frm1VspdData.Col = C_VALUE_EMP_NO
				frm1.VspdData.action =0                
	    else
		    frm1.txtEmp_no.focus
        end if
		Exit Function
	Else
		Call SetEmpName(arrRet, iWhere)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmpName(arrRet, iWhere)
	With frm1
	    if iWhere = 2 then
            ggoSpread.Source = Frm1.vspdData
	        With .VspdData
    		    .Row = .ActiveRow
                .Col = C_VALUE_NAME
                .Text = arrRet(1)
                .Col = C_VALUE_EMP_NO
                .Text = arrRet(0)
				.action =0                
	        End With
	    else
		    .txtEmp_no.value = arrRet(0)
		    .txtName.value = arrRet(1)
		    Call ggoOper.ClearField(Document, "2")					 '☜: Clear Contents  Field
		    Set gActiveElement = document.ActiveElement
		    Call txtEmp_no_Onchange()
		    .txtEmp_no.focus
        end if
	End With
End Sub

'==============================================================================
' 현재 경로가 수정되면 
'==============================================================================
Sub txtPath_OnChange()
	Dim strFileExt
	
	lgBlnFlgChgValue = True	
	document.all.ImgItemImage.src= frm1.txtPath.value
	
End Sub
'========================================================================================================
'   Event Name : txtEmp_no_Onchange             
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal

	frm1.txtName.value = ""
	frm1.txtDept_nm.value = ""
	frm1.txtRoll_pstn.value = ""
	frm1.txtEntr_dt.value = ""
	frm1.txtPay_grd.value = ""
'	Frm1.ImgItemImage.src = ""

    If  frm1.txtEmp_no.value = "" Then
        frm1.txtEmp_no.value = ""
    Else
        Frm1.ImgItemImage.src = ""
		frm1.txtPath.value = ""

	    IntRetCd =  FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
               
	    if  IntRetCd < 0 then

	    	strVal = "../../../CShared/image/default_picture.jpg"
		    Frm1.ImgItemImage.src = strVal  
	    
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
            Call ggoOper.ClearField(Document, "2")
            call InitVariables()
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
        Else
            frm1.txtName.value = strName
            frm1.txtDept_nm.value = strDept_nm
            frm1.txtRoll_pstn.value = strRoll_pstn
            frm1.txtPay_grd.value = strPay_grd1 & "-" & strPay_grd2
            frm1.txtEntr_dt.Text = UNIDateClientFormat(strEntr_dt)   
            
			Call CommonQueryRs(" COUNT(*) "," HAA070T "," emp_no= " & FilterVar( Frm1.txtEmp_no.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

			if   Replace(lgF0, Chr(11), "") > 0  then
				strVal = "../../ComASP/CPictRead.asp" & "?txtKeyValue=" & Frm1.txtEmp_no.value '☜: query key
				strVal = strVal     & "&txtDKeyValue=" & "default"                            '☜: default value
				strVal = strVal     & "&txtTable="     & "HAA070T"                            '☜: Table Name
				strVal = strVal     & "&txtField="     & "Photo"	                          '☜: Field
				strVal = strVal     & "&txtKey="       & "Emp_no"	                          '☜: Key
			else
				strVal = "../../../CShared/image/default_picture.jpg"
			end if

		    Frm1.ImgItemImage.src = strVal  
            
        End if 
        
    End if  
    
End Function 

function check_img_path(path)

	if (ggoSaveFile.fileExists(path) = 0)  then
		Check_img_path = true
	else
		Check_img_path = false
	end if
end function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST" ENCTYPE="MULTIPART/FORM-DATA">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>인사마스터사진등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
			    	        <TR>
			    	    		<TD CLASS="TD5" NOWRAP>사원</TD>
			    	    		<TD CLASS="TD6"><INPUT NAME="txtEmp_no" ALT="사원" TYPE="Text" MAXLENGTH=13 SiZE=13 tag=12XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmpName 1, ''"></TD>
			    	        	<TD CLASS="TD5" NOWRAP>성명</TD>
			    	    		<TD CLASS="TD6"><INPUT NAME="txtName" ALT="성명" TYPE="Text" MAXLENGTH=30 SiZE=20 tag=14></TD>
			            	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>부서명</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtDept_nm" ALT="부서명" TYPE="Text" SiZE=15 tag=14></TD>
			            		<TD CLASS="TD5" NOWRAP>직위</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtRoll_pstn" ALT="직위" TYPE="Text" SiZE=15 tag=14></TD>
			            	</TR>
			            	<TR>
							    <TD CLASS="TD5" NOWRAP>입사일</TD>
							    <TD CLASS=TD6><script language =javascript src='./js/h2002ma1_txtEntr_dt_txtEntr_dt.js'></script></TD>
			            		<TD CLASS="TD5" NOWRAP>급호</TD>
			            		<TD CLASS="TD6"><INPUT NAME="txtPay_grd" ALT="급호" TYPE="Text" SiZE=15 tag=14></TD>
			            	</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR HEIGHT=50>
								<TD CLASS=TD5 NOWRAP>경로</TD>
								<TD CLASS=TD656 NOWRAP><INPUT TYPE=FILE NAME="txtPath" SIZE=40 MAXLENGTH=100 tag=21 ALT="경로" value=""></TD>
							</TR>
							<TR VALIGN=TOP>
								<TD CLASS=TD5 NOWRAP>사진</TD>
								<TD CLASS=TD656 NOWRAP colspan=3>
								    <IMG src="../../../CShared/image/default_picture.jpg" ID= "ImgItemImage"  WIDTH=100 HEIGHT=150>
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
	<TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
	                <TD WIDTH=10>&nbsp;</TD>
	         		<TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">인사마스타</a></TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>


<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/incServer.asp"  -->
<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" href="../../inc/CommStyleSheet.css">

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/adoQuery.vbs"></SCRIPT>
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<%
    Dim menu_id
    Dim lang_cd
    
    menu_id = Trim(Request("menu_id"))
    lang_cd = Trim(Request("lang_cd"))
%>


<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance


'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "e1803mb1.asp"						           '☆: Biz Logic ASP Name
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
'--------------------------------------------------------------------------------------------------------
'  Constants for SpreadSheet #1
'--------------------------------------------------------------------------------------------------------

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" --> 
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================


Dim menu_id
Dim lang_cd
dim gLogoName

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================
'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    if  pOpt = "Q" then
        lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep       'You Must append one character(gColSep)
        lgKeyStream = lgKeyStream & "" & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtmenu_id.Value) & gColSep 
        lgKeyStream = lgKeyStream & Trim(frm1.txtlang_cd.Value) & gColSep 
    else
        lgKeyStream = Trim(parent.txtEmp_no.Value) & gColSep       'You Must append one character(gColSep)
        lgKeyStream = lgKeyStream & "" & gColSep
        lgKeyStream = lgKeyStream & Trim(frm1.txtmenu_id.Value) & gColSep 
        lgKeyStream = lgKeyStream & Trim(frm1.txtlang_cd.Value) & gColSep 
    end if
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
'	Call CommonQueryRs(" rTrim(LANG_CD),LANG_NM "," B_LANGUAGE "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
'    iCodeArr = lgF0
 '   iNameArr = lgF1
	frm1.txtlang_cd.options.length =0 
     iCodeArr = gLang & chr(11) 
    iNameArr = gLang & chr(11) 
	Call SetCombo2(frm1.txtlang_cd, iCodeArr, iNameArr, Chr(11))		
	frm1.txtlang_cd.value = gLang    
	Call CommonQueryRs(" DISTINCT menu_id, MENU_NAME,orders "," E11000T "," MENU_LEVEL=" & FilterVar("1", "''", "S") & " AND LANG_CD = " & FilterVar(gLang, "''", "S") & " order by orders", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) 
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtref_menu_id, iCodeArr, iNameArr, Chr(11))    

	Call CommonQueryRs(" DISTINCT menu_level "," E11000T "," ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = "2" & chr(11) & "1" & chr(11)
    iNameArr = "프로그램" & chr(11) & "메뉴" & chr(11)
    Call SetCombo2(frm1.txtmenu_level, iCodeArr, iNameArr, Chr(11))
   

    iCodeArr = "Y" & chr(11) & "N" & chr(11)
    iNameArr = "Y" & chr(11) & "N" & chr(11)
    Call SetCombo2(frm1.txtpro_use_flag, iCodeArr, iNameArr, Chr(11))    

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0120", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtpro_auth, iCodeArr, iNameArr, Chr(11))    

'ProtectTag(frm1.txtprog_order)
  
         
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

Sub ClearComboBox()

'frm1.txtlang_cd.options.length =0  
frm1.txtref_menu_id.options.length =0 
frm1.txtmenu_level.options.length =0 
frm1.txtpro_use_flag.options.length =0 
frm1.txtpro_auth.options.length =0 
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear  
    parent.document.All("nextprev").style.VISIBILITY = "hidden"
    'parent.document.All("emp_select").style.VISIBILITY = "hidden"
    menu_id = "<%=menu_id%>"
    lang_cd = "<%=lang_cd%>"
    lgIntFlgMode = OPMD_CMODE   'insert mode
    Call InitComboBox()

    Call LockField(Document)	
	Call LayerShowHide(0)

    Call SetToolBar("01110")

    if  menu_id <> "" then
        frm1.txtmenu_id.value = menu_id
'        frm1.txtlang_cd.value = Lang_cd
        Call parent.Click_OpenFrame(Replace(UCase(BIZ_PGM_ID),"MB","MA"))
        Call DbQuery(1)
    else
'        frm1.txtlang_cd.value = parent.txtLang.value
 '       frm1.txtlang_cd.focus()
    end if
    call show_menu_order()    
  
 End Sub
'========================================================================================
' Function Name : Window_onUnLoad
' Function Desc : 페이지 전환이나 화면이 닫힐 경우 실행해야 될 로직 처리 
'========================================================================================
Sub Form_unLoad()
End Sub

Function DbQuery(ppage)

    Dim strVal
    dim where_stm
    dim obj
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG
    'If Grid1.ChkChange() Then Exit Function
    'Call ClearField(document,2)
    Call LayerShowHide(1)
    Call MakeKeyStream("Q")
    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
    DbQuery = True                                                               '☜: Processing is NG
End Function

Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status

    lgIntFlgMode = OPMD_UMODE   'update mode
'    ProtectTag(frm1.txtlang_cd)
'    frm1.txtlang_cd.disabled = true
    ProtectTag(frm1.txtmenu_id)
    ProtectTag(frm1.txtmenu_level)    
    if frm1.txtmenu_level.value = 1 then
		ProtectTag(frm1.txtref_menu_id)
	end if
    frm1.txtmenu_name.focus()
    
call show_menu_order()
	

End Function

Function DbQueryFail()
    Err.Clear
	call FncNew()    
    lgIntFlgMode = OPMD_CMODE   'insert mode
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
	Dim strDate
	dim where_stm
	dim obj

    Err.Clear    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
   With Frm1

		if Trim(.txtmenu_id.value) = "" then
			Call  DisplayMsgBox("800094","X","X","X")
            .txtmenu_id.focus()
            exit function		
		else
			if lgIntFlgMode<> OPMD_UMODE then
				call commonqueryRS(" Menu_id ", " E11000T " , " menu_id =  " & FilterVar(frm1.txtmenu_id.value, "''", "S") & "", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)			
				if not(lgF0="") then
					Call DisplayMsgBox("800492","X",Trim(.txtmenu_id.value),"메뉴ID")   
					.txtmenu_id.focus()
					exit function	
				end if
			end if
        end if
        
        if Trim(.txtmenu_name.value) = "" then
			Call  DisplayMsgBox("800094","X","X","X")
            .txtmenu_name.focus()
            exit function
        end if
        
        if Trim(.txthref.value) = "" then
			Call  DisplayMsgBox("800094","X","X","X")
            .txthref.focus()
            exit function
        end if
        
        if  .txtmenu_level.value = "0" then
            .txtpro_type.value = "AS"
        elseif  .txtmenu_level.value = "1" then
            .txtpro_type.value = "MM"
        else
            .txtpro_type.value = "PP"
        end if
        
        if Trim(.txtmenu_order.value) = "" or Trim(.txtmenu_order.value)="개수초과" then
    		Call DisplayMsgBox("970028","X","메뉴개수","X")  
    		.txtmenu_id.focus()
            exit function    		      
    	end if

  End With
  

      
	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	DbSave = False														         '☜: Processing is NG
		
	Call LayerShowHide(1)
    Call MakeKeyStream("Q")

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	With Frm1
		.txtMode.Value        = "UID_M0002"                                        '☜: Save
		.txtFlgMode.Value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With
'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
    DbSave  = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    'Call InitComboBox()
     frm1.txtmenu_order.options.length =0 
	call DbQuery(1)
	'call show_menu_order()
	
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal

    Err.Clear                                                                    '☜: Clear err status
		
	Dim IntRetCD
	
    Err.Clear                                                                    '☜: Clear err status

	if Trim(frm1.txtmenu_id.value) = "" then
		Call  DisplayMsgBox("800094","X","X","X")
        frm1.txtmenu_id.focus()
        exit function
    end if
'	intRetCD = commonqueryRS(" Menu_id ", " E11000T " , " menu_id = '" & trim(frm1.txtmenu_id.value) & "' and lang_cd='" & frm1.txtlang_cd.value & "'", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	intRetCD = commonqueryRS(" Menu_id ", " E11000T " , " menu_id =  " & FilterVar(frm1.txtmenu_id.value, "''", "S") & "", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
	if replace(lgF0,chr(11),"") <> Trim(frm1.txtmenu_id.value)  then
		Call DisplayMsgBox("800480", "X", "X","X" ) 
'		frm1.txtmenu_id.value = ""
'		frm1.txtmenu_id.focus()
		call FncNew()
		Exit function
	End if
	
	IntRetCD =  DisplayMsgBox("900003", VB_YES_NO,"x","x")
	If IntRetCD = vbNo Then
		Exit Function
	End If

	DbDelete = False			                                                 '☜: Processing is NG
		
	Call LayerShowHide(1)
    Call MakeKeyStream("Q")
	With Frm1
		.txtMode.value        = "UID_M0003"                                        '☜: delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	DbDelete = True                                                             '⊙: Processing is NG
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'Call InitVariables()

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call FncNew()	
	
End Function

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Call ClearField(document,2)
	
    'ReleaseTag(frm1.txtlang_cd)                                                   '⊙: Initializes local global variables
    
    'ReleaseTag(frm1.txtmenu_id)                                                   '⊙: Initializes local global variables
    
     
    'parent.document.All("nextprev").style.VISIBILITY = "hidden"
    'parent.document.All("emp_select").style.VISIBILITY = "hidden"
    
    lgIntFlgMode = OPMD_CMODE   'insert mode
  
    Call LockField(Document)	
  	frm1.txthref.disabled = false
'    frm1.txtlang_cd.value = parent.txtLang.value
    frm1.txtmenu_level.selectedIndex    = 0
    frm1.txtpro_use_flag.selectedIndex  = 0
    frm1.txtpro_auth.selectedIndex      = 0
    frm1.txtref_menu_id.selectedIndex   = 0
    frm1.txtmenu_order.selectedIndex   = 0
    frm1.txtmenu_order.options.length =0  	
'    frm1.txtlang_cd.focus()
    
 	order_list_id.innerHTML = ""
	order_list_name.innerHTML = ""
	order_list_order.innerHTML = ""	   
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
  call show_menu_order()
call ClearComboBox  
call InitComboBox()  
    FncNew = True											 '☜: Processing is OK

  
End Function

Sub SubPrint(objFrame)
    Set objActiveEl = document.activeElement
    objFrame.focus()
    objFrame.print()
    objActiveEl.focus
    Set objActiveEl = nothing
End Sub

Sub txtmenu_level_OnChange()

    IF  frm1.txtmenu_level.value = "2" then
        ReleaseTag(frm1.txtref_menu_id)
        frm1.txthref.value = ""
        frm1.txtref_menu_id.disabled = false
   frm1.txtref_menu_id.selectedIndex    = 0
    else

		ProtectTag(frm1.txtref_menu_id)
		frm1.txtref_menu_id.value = ""
		frm1.txtref_menu_id.disabled = true
		frm1.txthref.value = "#"
    end if
	
	order_list_id.innerHTML = ""
	order_list_name.innerHTML = ""
	order_list_order.innerHTML = ""	
	frm1.txtmenu_order.options.length =0
	call show_menu_order()     

End Sub

Sub txtref_menu_id_OnChange()
	dim obj		
	
	frm1.txtmenu_order.options.length =0   
	call show_menu_order()

	IF  frm1.txtmenu_level.value = "2" and frm1.txtref_menu_id.value = frm1.txtoriginal_ref_id.value   then 
		Set obj = Document.CreateElement("OPTION")	
		obj.Text = frm1.txtoriginal_order.value
		obj.Value = frm1.txtoriginal_order.value
		frm1.txtmenu_order.Add(obj)
		frm1.txtmenu_order.selectedIndex   = 1
	end if


End Sub

Function show_menu_order()
	
	dim i
	dim j
	dim ok
	dim iCodeArr
	dim pCodeArr
	dim pNameArr
	dim where_stm
	dim max_order

	Err.Clear  
	IF  frm1.txtmenu_level.value = "1" then	
		where_stm = " MENU_LEVEL=" & FilterVar("1", "''", "S") & " AND LANG_CD = " & FilterVar(gLang, "''", "S") & " order by orders"
		max_order=12

  	ELSEIF  frm1.txtmenu_level.value = "2" then	
		where_stm = " MENU_LEVEL=" & FilterVar("2", "''", "S") & " AND LANG_CD = " & FilterVar(gLang, "''", "S") & " and ref_menu_id= " & FilterVar( frm1.txtref_menu_id.value, "''", "S") & " order by orders"
		max_order=10
	END IF			
	if CommonQueryRs(" left(menu_id,20),left(MENU_NAME,20),orders "," E11000T ",where_stm, lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  then
			
			order_list_id.innerHTML = replace(lgF0 , chr(11),"<br>")
			order_list_name.innerHTML = replace(lgF1 , chr(11),"<br>")
			order_list_order.innerHTML = replace(lgF2 , chr(11),"<br>")
		
		    iCodeArr = lgF2
			iCodeArr = Split(iCodeArr,chr(11))
			pCodeArr = ""
			pNameArr = ""
			For i = 1 To iCodeArr(UBound(iCodeArr)-1)+1
				ok= true
				For j = 0 To UBound(iCodeArr)
				if cstr(i) = iCodeArr(j)  then
						ok= false
				end if
				Next
			
				if ok then
					if i < max_order+1 then
						pCodeArr = pCodeArr & cstr(i) & chr(11) 
					else 
						pCodeArr = pCodeArr & "개수초과" & chr(11) 
					end if
				end if
			Next    
		else
			pCodeArr = "1" & chr(11)
			order_list_id.innerHTML = ""
			order_list_name.innerHTML = ""
			order_list_order.innerHTML = ""
			
		end if 
		pNameArr=pCodeArr
		Call SetCombo2(frm1.txtmenu_order, pCodeArr,pNameArr, Chr(11))     

End Function





'========================================================================================================

'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

<!--
'########################################################################################################
'#						6. TAG 부																		#
'######################################################################################################## 
-->
</SCRIPT>
<!-- #Include file="../../inc/uniSimsClassID.inc" --> 
</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 width=749 border=0>
        <TR>
            <TD width=13></TD>
            <TD>
                <TABLE cellSpacing=0 cellPadding=0 width=722 border=0 bgcolor=#ffffff>
                    <TR>
                        <TD colspan=4>
                            <TABLE cellSpacing=1 cellPadding=0 width=100% border=0 bgcolor=#ffffff>
		                        <TR height=20>
		                            <TD></TD>
		                            <TD></TD>
		                            <TD></TD>
		                            <TD></TD>
                                </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>언어</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3>
		                                <SELECT NAME="txtlang_cd" STYLE="WIDTH: 100px" TAG="22"></SELECT>
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>메뉴ID</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3>
		                                <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtmenu_id" TYPE="Text" MAXLENGTH=25 SiZE=30 tag="22">
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>메뉴명</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3>
		                                <INPUT CLASS="SINPUTTEST_STYLE" NAME="txtmenu_name" TYPE="Text" MAXLENGTH=24 SiZE=24 tag="22">
		                            </TD>
                                </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>프로그램</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3>
		                                <INPUT CLASS="SINPUTTEST_STYLE" NAME="txthref" TYPE="Text" MAXLENGTH=25 SiZE=30 tag="26">
                                    </TD>      
		                        </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>메뉴타입</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3>
		                                <SELECT NAME="txtmenu_level" STYLE="WIDTH: 100px" TAG="22"></SELECT>
		                                <INPUT NAME="txtpro_type" TYPE=hidden>
                                    </TD>      
		                        </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>사용여부</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3>
		                                <SELECT NAME="txtpro_use_flag" STYLE="WIDTH: 100px" TAG="22"></SELECT>
                                    </TD>      
		                        </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>레벨</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3>
		                                <SELECT NAME="txtpro_auth" STYLE="WIDTH: 100px" TAG="22"></SELECT>
                                    </TD>      
		                        </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>상위메뉴</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3>
		                                <SELECT NAME="txtref_menu_id" STYLE="WIDTH: 100px" TAG="22"></SELECT>
                                    </TD>      
		                        </TR>
                                <TR>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>메뉴순서</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3>
		                                <SELECT NAME="txtmenu_order" STYLE="WIDTH: 100px" TAG="22"></SELECT>
                                    </TD>      
		                        </TR>
		                        <TR height=63>
		                            <TD CLASS="TDFAMILY_TITLE" NOWRAP>메뉴순서리스트</TD>
		                            <TD CLASS="TDFAMILY2" colspan=3>
		                                <table>
		                                <tr><td width=150 bgcolor=#e6e3fa>메뉴ID</td><td width=250 bgcolor=#e6e3fa>메뉴명</td><td width=80 bgcolor=#e6e3fa>메뉴순서</td></tr>
										<tr><td><SPAN CLASS="normal" ID="order_list_id" style="BACKGROUND-COLOR: #E9EDF9;BORDER:0" >&nbsp;</SPAN></td>
										    <td><SPAN CLASS="normal" ID="order_list_name" style="BACKGROUND-COLOR: #E9EDF9;BORDER:0" >&nbsp;</SPAN></td>
										    <td><SPAN CLASS="normal" ID="order_list_order" style="BACKGROUND-COLOR: #E9EDF9;BORDER:0" >&nbsp;</SPAN></td>
										</tr>
										</table>
									</TD>
                                </TR>
		                        <TR height=10>
		                            <TD></TD>
		                            <TD></TD>
		                            <TD></TD>
		                            <TD></TD>
                                </TR>
                            </TABLE>
                        </TD>
                    </TR>
                </TABLE>
            </TD>
            <TD width=14></TD>
        </TR>
    </TABLE>

    <TABLE cellSpacing=0 cellPadding=0 width=700 border=0 bgcolor=#ffffff>
        <TR><TD HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH="100%" HEIGHT=0 FRAMEBORDER=0 SCROLLING=yes noresize framespacing=0></IFRAME></TD></TR>
    </TABLE>
    <INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtoriginal_ref_id"        TAG="24">
    <INPUT TYPE=HIDDEN NAME="txtoriginal_order"        TAG="24">
</FORM>	

</BODY>
</HTML>

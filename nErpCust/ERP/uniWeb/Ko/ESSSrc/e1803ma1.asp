<%@ LANGUAGE="VBSCRIPT" %>
<% Response.Expires = -1%>

<HTML>
<HEAD>
<TITLE><%=Request("strTitle")%></TITLE>

<!-- #Include file="../ESSinc/incServer.asp"  -->

<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/incEvent.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../ESSinc/adoQuery.vbs"></SCRIPT>
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<%
    Dim menu_id
    Dim lang_cd
    
    menu_id = Trim(Request("menu_id"))
    lang_cd = Trim(Request("lang_cd"))
%>

<Script Language="VBScript">
Option Explicit  

Const BIZ_PGM_ID      = "e1803mb1.asp"						           '☆: Biz Logic ASP Name

<!-- #Include file="../ESSinc/lgvariables.inc" --> 

Dim menu_id
Dim lang_cd
dim gLogoName

'========================================================================================================
' Function Name : MakeKeyStream
'========================================================================================================
Sub MakeKeyStream(pOpt)
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
End Sub        

'========================================================================================================
' Name : InitComboBox()
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr

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
End Sub

'========================================================================================================
' Name : ClearComboBox()
'========================================================================================================
Sub ClearComboBox()
	frm1.txtref_menu_id.options.length =0 
	frm1.txtmenu_level.options.length =0 
	frm1.txtpro_use_flag.options.length =0 
	frm1.txtpro_auth.options.length =0 
End Sub

'========================================================================================================
' Name : Form_Load
'========================================================================================================
Sub Form_Load()

    Err.Clear  
    parent.document.All("nextprev").style.VISIBILITY = "hidden"
    menu_id = "<%=menu_id%>"
    lang_cd = "<%=lang_cd%>"
    lgIntFlgMode = OPMD_CMODE   'insert mode
    Call InitComboBox()

    Call LockField(Document)	
	Call LayerShowHide(0)

    Call SetToolBar("01110")

    if  menu_id <> "" then
        frm1.txtmenu_id.value = menu_id
        Call parent.Click_OpenFrame(Replace(UCase(BIZ_PGM_ID),"MB","MA"))
        Call DbQuery(1)
    end if
    call show_menu_order()    
  
 End Sub
'========================================================================================
' Function Name : Form_unLoad
'========================================================================================
Sub Form_unLoad()
End Sub

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery(ppage)

    Dim strVal
    dim where_stm
    dim obj
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    Call LayerShowHide(1)
    Call MakeKeyStream("Q")
    strVal = BIZ_PGM_ID & "?txtMode="            & "UID_M0001"                   '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
    DbQuery = True                                                               '☜: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()
    Err.Clear                                                                    '☜: Clear err status

    lgIntFlgMode = OPMD_UMODE   'update mode
    ProtectTag(frm1.txtmenu_id)
    ProtectTag(frm1.txtmenu_level)  
'	frm1.txtmenu_level.disabled  = true
  
    if frm1.txtmenu_level.value = 1 then
		ProtectTag(frm1.txtref_menu_id)
	end if
    frm1.txtmenu_name.focus()
	call show_menu_order()
End Function

'========================================================================================
' Function Name : DbQueryFail
'========================================================================================
Function DbQueryFail()
    Err.Clear
	call FncNew()    
    lgIntFlgMode = OPMD_CMODE   'insert mode
End Function

'========================================================================================================
' Name : DbSave
'========================================================================================================
Function DbSave()
	Dim strVal
	Dim strDate
	dim where_stm
	dim obj

    Err.Clear    
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

	DbSave = False														         '☜: Processing is NG
		
	Call LayerShowHide(1)
    Call MakeKeyStream("Q")

	With Frm1
		.txtMode.Value        = "UID_M0002"                                        '☜: Save
		.txtFlgMode.Value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

	ReleaseTag(frm1.txtmenu_level)
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
    DbSave  = True                                                               '☜: Processing is NG
    
 
End Function

'========================================================================================================
' Function Name : DbSaveOk
'========================================================================================================
Function DbSaveOk()
     frm1.txtmenu_order.options.length =0 
     
	call DbQuery(1)
End Function

'========================================================================================================
' Name : DbDelete
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
	intRetCD = commonqueryRS(" Menu_id ", " E11000T " , " menu_id =  " & FilterVar(frm1.txtmenu_id.value, "''", "S") & "", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	
	if replace(lgF0,chr(11),"") <> Trim(frm1.txtmenu_id.value)  then
		Call DisplayMsgBox("800480", "X", "X","X" ) 
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

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

	DbDelete = True                                                             '⊙: Processing is NG
End Function

'========================================================================================================
' Function Name : DbDeleteOk
'========================================================================================================
Function DbDeleteOk()
	Call FncNew()	
End Function

'========================================================================================================
' Name : FncNew
'========================================================================================================
Function FncNew()
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Call ClearField(document,2)
    
    lgIntFlgMode = OPMD_CMODE   'insert mode
  
    Call LockField(Document)	
  	frm1.txthref.disabled = false
    frm1.txtmenu_level.selectedIndex    = 0
    frm1.txtpro_use_flag.selectedIndex  = 0
    frm1.txtpro_auth.selectedIndex      = 0
    frm1.txtref_menu_id.selectedIndex   = 0
    frm1.txtmenu_order.selectedIndex   = 0
    frm1.txtmenu_order.options.length =0  	
    
 	order_list_id.innerHTML = ""
	order_list_name.innerHTML = ""
	order_list_order.innerHTML = ""	   

	call show_menu_order()
	call ClearComboBox  
	call InitComboBox()  
    FncNew = True											 '☜: Processing is OK
End Function

'========================================================================================================
' Name : show_menu_order
'========================================================================================================
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
'                        5.5 Tag Event
'========================================================================================================
Sub txtmenu_level_OnChange()

    IF  frm1.txtmenu_level.value = "2" then
        ReleaseTag(frm1.txtref_menu_id)
        frm1.txthref.value = ""
        frm1.txtref_menu_id.disabled = false
		frm1.txtref_menu_id.selectedIndex = 0
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
</SCRIPT>
<!-- #Include file="../ESSinc/uniSimsClassID.inc" --> 
</HEAD>

<BODY topmargin=0 leftmargin=0 marginheight=0 marginwidth=0>
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
    <TABLE cellSpacing=0 cellPadding=0 border=0 width=716>
        <TR>
           <td height="10"></td>
        </TR>
        <TR>
            <TD>
                <TABLE width=100% cellSpacing=1 cellPadding=0 border=0 bgcolor=#DDDDDD>
                    <TR>
		                <TD CLASS="ctrow01" NOWRAP>언어</TD>
		                <TD CLASS="ctrow06" colspan=3>
		                    <SELECT CLASS="form01" NAME="txtlang_cd" STYLE="WIDTH: 100px" TAG="22"></SELECT>
		                </TD>
                    </TR>
                    <TR>
		                <TD CLASS="ctrow01" NOWRAP>메뉴ID</TD>
		                <TD CLASS="ctrow06" colspan=3>
		                    <INPUT CLASS="form01" NAME="txtmenu_id" TYPE="Text" MAXLENGTH=25 SiZE=30 tag="22">
		                </TD>
                    </TR>
                    <TR>
		                <TD CLASS="ctrow01" NOWRAP>메뉴명</TD>
		                <TD CLASS="ctrow06" colspan=3>
		                    <INPUT CLASS="form01" NAME="txtmenu_name" TYPE="Text" MAXLENGTH=24 SiZE=24 tag="22">
		                </TD>
                    </TR>
                    <TR>
		                <TD CLASS="ctrow01" NOWRAP>프로그램</TD>
		                <TD CLASS="ctrow06" colspan=3>
		                    <INPUT CLASS="form01" NAME="txthref" TYPE="Text" MAXLENGTH=25 SiZE=30 tag="22">
                        </TD>      
		            </TR>
                    <TR>
		                <TD CLASS="ctrow01" NOWRAP>메뉴타입</TD>
		                <TD CLASS="ctrow06" colspan=3>
		                    <SELECT CLASS="form01"  NAME="txtmenu_level" STYLE="WIDTH: 100px" tag="22"></SELECT>
		                    <INPUT CLASS="form01"  NAME="txtpro_type" TYPE=hidden>
                        </TD>      
		            </TR>
                    <TR>
		                <TD CLASS="ctrow01" NOWRAP>사용여부</TD>
		                <TD CLASS="ctrow06" colspan=3>
		                    <SELECT CLASS="form01" NAME="txtpro_use_flag" STYLE="WIDTH: 100px" tag="22"></SELECT>
                        </TD>      
		            </TR>
                    <TR>
		                <TD CLASS="ctrow01" NOWRAP>레벨</TD>
		                <TD CLASS="ctrow06" colspan=3>
		                    <SELECT CLASS="form01" NAME="txtpro_auth" STYLE="WIDTH: 100px" tag="22"></SELECT>
                        </TD>      
		            </TR>
                    <TR>
		                <TD CLASS="ctrow01" NOWRAP>상위메뉴</TD>
		                <TD CLASS="ctrow06" colspan=3>
		                    <SELECT CLASS="form01" NAME="txtref_menu_id" STYLE="WIDTH: 100px" tag="22"></SELECT>
                        </TD>      
		            </TR>
                    <TR>
		                <TD CLASS="ctrow01" NOWRAP>메뉴순서</TD>
		                <TD CLASS="ctrow06" colspan=3>
		                    <SELECT CLASS="form01" NAME="txtmenu_order" STYLE="WIDTH: 100px" tag="22"></SELECT>
                        </TD>      
		            </TR>
		            <TR height=63>
		                <TD CLASS="ctrow01" NOWRAP>메뉴순서리스트</TD>
		                <TD CLASS="ctrow06" colspan=3>
		                    <table cellSpacing=1 cellPadding=0 border=0 bgcolor=#FFFFFF>
		                    <tr><td height="5"></td>
		                    </tr>
		                    <tr><td CLASS="TDFAMILY_TITLE1" width=150>메뉴ID</td>
		                        <td CLASS="TDFAMILY_TITLE1" width=250>메뉴명</td>
		                        <td CLASS="TDFAMILY_TITLE1" width=80>메뉴순서</td>
		                    </tr>
							<tr><td CLASS="listrow02"><SPAN CLASS="listrow01" ID="order_list_id" ></SPAN></td>
							    <td CLASS="listrow02"><SPAN CLASS="listrow01" ID="order_list_name"></SPAN></td>
							    <td CLASS="listrow02"><SPAN CLASS="listrow01" ID="order_list_order"></SPAN></td>
							</tr>
		                    <tr><td height="5"></td>
		                    </tr>
							</table>
						</TD>
                    </TR>
                </TABLE>
            </TD>
        </TR>
        <TR>
           <td height="10"></td>
        </TR>
    </TABLE>

    <TABLE cellSpacing=0 cellPadding=0 border=0>
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

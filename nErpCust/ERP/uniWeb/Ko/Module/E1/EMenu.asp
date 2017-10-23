<%@ Language=VBScript %>
<HTML>
<%
Response.Buffer = True
Response.Expires = -1
%>
		<!-- #Include file="../../inc/Adovbs.inc"  -->
		<!-- #Include file="../../inc/incServerAdoDb.asp" -->
		<!-- #Include file="../../inc/incServer.asp" -->
		<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
		<!-- #Include file="../../inc/incSvrFuncSims.inc" -->

<% 
Dim	Name
Dim	dept_nm
Dim	entr_dt
Dim internal_cd
Dim nat_cd

    Call SubOpenDB(lgObjConn)                                                             '☜: Make a DB Connection
	
	if gEmpNo = "unierp" then
		Name = "unierp"
	else
		lgStrSQL = " SELECT Emp_no, NAME, dept_nm, pay_grd2, entr_dt, internal_cd, nat_cd "
		lgStrSQL = lgStrSQL & " FROM haa010t where emp_no= " & FilterVar(gEmpNo, "''", "S") & ""
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = true Then
			Name	   = lgObjRs("NAME")
			dept_nm    = lgObjRs("dept_nm")
			entr_dt    = lgObjRs("entr_dt")
			internal_cd    = lgObjRs("internal_cd")
			nat_cd     = lgObjRs("nat_cd")
		End IF  
		Call SubCloseRs(lgObjRs)		
	end if

    Call SubCloseDB(lgObjConn) 

'=====================MENU BAR 설정변수=====================
Const GCOL			= ":"
Const GROW			= ";"
Const MENUBAR		= "MENUBAR"
Const TOPMENUBAR	= "TOPMENUBAR"
Const TOPSUBBAR		= "TOPSUBBAR"
Const LEFTMENUBAR	= "LEFTMENUBAR"
Const SUBNAME		= "_SUB"
Const LEFTNAME		= "_LEFT"
Const LEFTID		= "_LEFTID"
Const BARNAME		= "_BAR"
Const MENUAREA		= "-1"
Const MAINCLASS		= "MAINMENU"
Const SUBCLASS		= "SUBMENU"
Const LEFTMAINCLASS	= "LEFTMAIN"
Const LEFTSUBCLASS	= "LEFTSUB"


%>
<!-- #Include file="../../inc/incSvrVarSims.inc"  -->
<!-- #Include file="../../inc/incSvrFuncSims.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" href="../../inc/MenuStyleSheet.css">
			<SCRIPT LANGUAGE="VBScript" SRC="../../inc/ccm.vbs"></SCRIPT>
			<SCRIPT LANGUAGE="VBScript" SRC="../../inc/variables.vbs"></SCRIPT>
			<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCookie.vbs"></SCRIPT>
			<SCRIPT LANGUAGE="VBScript" SRC="../../inc/operation.vbs"></SCRIPT>
			<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCommFunc.vbs"></SCRIPT>
			<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEvent.vbs">   </SCRIPT>			
<Script language="vbscript">
Option Explicit
'=====================MENU BAR 설정변수=====================
Const GCOL					= "<%=GCOL%>"
Const GROW					= "<%=GROW%>"
Const MENUBAR				= "<%=MENUBAR%>"
Const TOPMENUBAR			= "<%=TOPMENUBAR%>"
Const TOPSUBBAR				= "<%=TOPSUBBAR%>"
Const LEFTMENUBAR			= "<%=LEFTMENUBAR%>"
Const SUBNAME				= "<%=SUBNAME%>"
Const LEFTNAME				= "<%=LEFTNAME%>"
Const LEFTID				= "<%=LEFTID%>"
Const BARNAME				= "<%=BARNAME%>"
Const MENUAREA				= "<%=MENUAREA%>"
Const MAINCLASS				= "<%=MAINCLASS%>"
Const SUBCLASS				= "<%=SUBCLASS%>"
Const LEFTMAINCLASS			= "<%=LEFTMAINCLASS%>"
Const LEFTSUBCLASS			= "<%=LEFTSUBCLASS%>"
Const LEFTMENUWIDTH			= 200

Const VIEWMENUCNT			= 7
Const SKIPMENUCNT			= 1
Const MENUXSPACE			= 10
Const MENUYSPACE			= 15
Const MCLASS				= "TOPMAIN"
Const SCLASS				= "TOPSUB"
Const LCLASS				= "LEFTMENU"

Const MENUTOP				= "MENUTOP"
Const MENUEND				= "MENUEND"

Const CLICK_COLOR			= "#E07000"
Const OVER_COLOR			= "#717173"
Const SUB_OUT_COLOR			= "#065D88"
Const LEFT_OUT_COLOR		= "#017FA8"
Const MAIN_COLOR			= "#0242AC"

Const OVER_CURSOR			= "hand"
Const OUT_CURSOR			= "auto"
Const MOVER_CURSOR			= "default"

Const LEFTKEY				= 37
Const RIGHTKEY				= 39
Const UPKEY					= 38
Const DOWNKEY				= 40
Const TABKEY				= 9
Const ESCKEY				= 27

Class Menu

	Dim ID	
	Dim URL
	Dim GROUP
	Dim NEXTFLAG
	Dim OPENFLAG
	Dim CLICKFLAG
	Dim MOVERFLAG
	Dim DISPLAYFLAG
	Dim TOPEND
	Dim MTitle
	Dim PROTYPE
	
End Class

Dim TempMain,TempSub,TempLeft
Dim TopMain,TopSub,LeftMenu
Dim COpenSub,COpenLeft,CurrURL
Dim oldGroup
Dim lgFncLogoff	'로그오프function실행 여부 

COpenSub = ""
COpenLeft = ""
CurrURL = "A1103MA.asp"

'========================= Window Event =======================
'==============================================================
'Function: Window_onLoad()
'==============================================================

Function Window_onLoad()

	On Error Resume Next
	Err.Clear 

    if  Trim(txtemp_no.value) = "" Then
        document.location = "../../unisims.asp"
    End If

    window.document.body.scroll = "no"

	Call Menu_Init(TempMain,MCLASS)		
	Call Menu_Init(TempSub,SCLASS)	
	Call Menu_Init(TempLeft,LCLASS)
	Call FncHomeMenu()
    Call Menu_Display(MCLASS,TopMain)

    document.All("nextprev").style.VISIBILITY = "hidden"
'------------첫화면에 공지사항 settting
	document.all("divHomeMenu").style.VISIBILITY = "hidden"
	document.all("DivPgmMenu").style.VISIBILITY = "visible"
	Call formmenu_onLoad(inPagevalue) 
	document.All("formmenu").src = "ESSBoard_list.asp"
	txtTitle.value="공지사항"
    document.title = gLogoName & " - " & LeftMenu(InIDx).MTitle & " [ " & "<%=NAME%>" & " ]"	
'-----------------------------
	Window_onLoad = True
End Function
'==============================================================
'Function: Window_unLoad()
'==============================================================
Function Window_onUnLoad()
	Dim i
	On Error Resume Next
	Err.Clear 

	If IsArray(TopMain) Then
		For i = 0 To Ubound(TopMain)
			Set TopMain(i) = nothing
		Next
	End If
	If IsArray(TopSub) Then
		For i = 0 To Ubound(TopSub)
			Set TopSub(i) = nothing
		Next
	End If
	If IsArray(LeftMenu) Then
		For i = 0 To Ubound(LeftMenu)
			Set LeftMenu(i) = nothing
		Next
	End If
	If lgFncLogoff = False Then FncLogoff(1)
	Window_unLoad = True
End Function
'==============================================================
'Function: Document_onMouseOver()
'==============================================================
Function Document_onMouseOver()
	Dim CuEvObj
	On Error Resume Next
	Err.Clear 
	
	Set CuEvObj = window.event.srcElement
	Call Menu_Operation(CuEvObj)
	Set CuEvObj = nothing
	
	Document_onMouseOver = True
End Function
'==============================================================
'Function: Menu_Analysis(CuEvObj)
'==============================================================
Sub Menu_Analysis(CuEvObj,IDx,InList,InClass)

	If Not IsNull(CuEvObj.getAttribute("LEVEL")) Then	
		If Not IsNull(CuEvObj.id) Then
			IDx = Menu_Search(TopMain,CuEvObj.id,"MENUIDX")
			If IDx <> -1 Then
				InList = TopMain
				InClass = MCLASS				
				Exit Sub
			End If
			IDx = Menu_Search(TopSub,CuEvObj.id,"MENUIDX")
			If IDx <> -1 Then
				InList = TopSub
				InClass = SCLASS				
				Exit Sub
			End If
			IDx = Menu_Search(LeftMenu,CuEvObj.id,"MENUIDX")
			If IDx <> -1 Then
				InList = LeftMenu
				InClass = LCLASS				
				Exit Sub
			End If
			If Not IsArray(InList) Then
				InList = False
			End If			
		Else
			IDx = -1
		End If		
	Else
		IDx = -1
		InClass = False		
	End If	
End Sub
'==============================================================
'Function: Menu_Operation(CuEvObj)
'==============================================================
Function Menu_Operation(CuEvObj)
	Dim IDx,InList,InClass
	On Error Resume Next
	Err.Clear 
	Call Menu_Analysis(CuEvObj,IDx,InList,InClass)	
	If IDx <> -1 Then
		Call Close_Menu(InList,IDx,InClass)
		Call MouseOver_Menu(IDx,InList,InClass)
	End If
	
	If InClass = False And InClass <> "" Then
		Call Close_Menu(TopMain,"",MCLASS)
		Call Close_Menu(TopSub,"",SCLASS)
		Call Close_Menu(LeftMenu,"",LCLASS)
	End If
	Menu_Operation = True
	
End Function
'==============================================================
'Function: Search_NextID(InArr)
'==============================================================
Function Search_NextID(InList,InArr,InClass)
	Dim IDx,TempIDx,i,OutIDx
	On Error Resume Next
	Err.Clear 

	IDx = -1
	If IsArray(InArr) Then		
		For i = 0 To Ubound(InArr)
			If InList(InArr(i)).OPENFLAG = True And InList(InArr(i)).MOVERFLAG = True And InList(InArr(TempIDx)).CLICKFLAG = False Then			
				TempIDx = i
			End If			
		Next
		If TempIDx <> "" Then		
			If InList(InArr(TempIDx + 1)).CLICKFLAG = True Then
				Temp = Temp + 1
			End If			
			If TempIDx > Ubound(InArr) Or TempIDx < Lbound(InArr) Then
				IDx = Lbound(InArr)
			Else
				IDx = TempIDx + 1
			End If			
		End If
		If IDx <> -1 Then	
			Search_NextID = InArr(IDx)			
		Else	
			Search_NextID = IDx			
		End If
		Exit Function
	End If
	Search_NextID = IDx
End Function
'==============================================================
'Function: DOWNKDY_handler(CuEvObj)
'==============================================================
Function DOWNKDY_handler(CuEvObj)
	Dim IDx,InList,InClass,TempArr,CurrObj,OpenObj
	Dim i
	On Error Resume Next
	Err.Clear 

	Call Menu_Analysis(CuEvObj,IDx,InList,InClass)
	If IDx <> -1 Then	
		Select Case InClass
		Case MCLASS			
			If InList(IDx).PROTYPE = "MM" Then
				If UCase(document.all(InList(IDx).ID & SUBNAME).style.visibility) = "VISIBLE" Then
					TempArr = Menu_Return(TopSub,InList(IDx).ID & SUBNAME,"GROUP")
					
					IDx = Search_NextID(TopSub,TempArr,InClass)
					If IDx	<> -1 Then
						Set CurrObj = document.all(TopSub(IDx).ID)						
						Call Menu_Operation(CurrObj)
						Set CurrObj = nothing
					End If
				Else				
					Call Menu_Operation(CuEvObj)
				End If
			End If		
		Case LCLASS
				TempArr = Menu_Return(InList,InList(IDx).GROUP,"GROUP")
				IDx = Search_NextID(InList,TempArr,InClass)				
				If IDx <> -1 Then
					Set OpenObj = document.all(InList(IDx).ID)
					Set CurrObj = document.all(replace(InList(IDx).ID,LEFTID,""))
					Call Menu_Operation(CurrObj)					
					Set OpenObj = nothing
					Set CurrObj = nothing
				End If
		End Select
	End If
		
	DOWNKDY_handler = True
End Function
'==============================================================
'Function: RIGHTKEY_handler(CuEvObj)
'==============================================================
Function RIGHTKEY_handler(CuEvObj)
	Dim IDx,TempArr,CurrObj
	Dim i,OFlag
	On Error Resume Next
	Err.Clear 
	
	OFlag = False
	RIGHTKEY_handler = True
End Function
'==============================================================
'Function: Document_onKeyDown()
'==============================================================
Function Document_onKeyDown()
	Dim CuEvObj,KeyCode

	On Error Resume Next
	Err.Clear 
	
	Set CuEvObj = window.event.srcElement		
	KeyCode = window.event.keycode

	Select Case KeyCode
		Case DOWNKEY
			Call DOWNKDY_handler(CuEvObj)
		Case UPKEY
		Case LEFTKEY
		Case RIGHTKEY
			Call RIGHTKEY_handler(CuEvObj)
		Case TABKEY
		Case ESCKEY
		Case 13		' Enter Key: Used as Query in Condition
			If Left(CuEvObj.getAttribute("tag"),1) = "1" Then
				Call formmenu.DbQuery(1)
			end if
	End Select		
	
	Document_onKeyDown	= True	
End Function

'==============================================================
'Function: Document_onClick()
'==============================================================
Function Document_onClick()
Dim StrURL,CuEvObj,IDx
	On Error Resume Next
	Err.Clear 
	
	Set CuEvObj = window.event.srcElement
	IDx = Menu_Search(TopMain,CuEvObj.id,"MENUIDX")
	
	If IDx <> -1 Then
		If (TopMain(IDx).PROTYPE = "AS" Or TopMain(IDx).PROTYPE = "AE") And TopMain(IDx).MOVERFLAG = True Then		
			Call SlipMenu(TopMain,TopMain(IDx).PROTYPE)
			Call Menu_Display(MCLASS,TopMain)			
		End If
		If TopMain(IDX).PROTYPE = "MP" Then
			Call Click_Menu(IDx,TopMain,MCLASS)
		End If
	End If	
	IDx = Menu_Search(TopSub,CuEvObj.id,"MENUIDX")
	If IDx <> -1 Then
		Call Click_Menu(IDx,TopSub,SCLASS)	
	End If
	
	IDx = Menu_Search(LeftMenu,CuEvObj.id,"MENUIDX")
	If IDx <> -1 Then
		Call Click_Menu(IDx,LeftMenu,LCLASS)		
	End If
	
	Set CuEvObj = nothing
	Document_onClick = True
End Function
Function txtEmp_no2_Onchange()
    On Error Resume Next
    Err.Clear
	call formmenu.txtEmp_no2_Onchange()
End Function

sub menu_move(strType)

	Call SlipMenu(TopMain,strType)
	Call Menu_Display(MCLASS,TopMain)
end sub

'==============================================================
'Function: formmenu_onLoad()
'==============================================================
Function formmenu_onLoad(inPagevalue)
Dim IDx
	On Error Resume Next
	Err.Clear 	
	
        CurrURL = UCase(inPagevalue)
		IDx = Menu_Search(TopSub,CurrURL,"URLIDX")		
		If IDx <> -1 Then		
		   Call Click_Menu(IDx,TopSub,SCLASS)
		End If
		
	formmenu_onLoad = True
End Function
'========================= String 처리 ========================
'==============================================================
'Function: Menu_Search(InList,InCom,InType)
'==============================================================
Function Menu_Search(InList,InCom,InType)
Dim i
	On Error Resume Next
	Err.Clear
	If IsArray(InList ) Then
		For i = 0 To Ubound(InList)		
			Select Case InType
			Case "MENUIDX"			
				If InList(i).ID = InCom Then
					Menu_Search = i
					Exit Function
				End If
			Case "GROUPIDX"
				If InList(i).GROUP = InCom Then
					Menu_Search = i
					Exit Function
				End If
			Case "CLICKIDX"
				If InList(i).CLICKFLAG = InCom Then
					Menu_Search = i
					Exit Function
				End If
			Case "URLIDX"
				If InList(i).URL = InCom Then
					Menu_Search = i
					Exit Function
				End If
			Case "OPENIDX"
				If InList(i).OPENFLAG = InCom Then
					Menu_Search = i
					Exit Function
				End If
			End Select
		Next
	End If
	Menu_Search = -1
End Function
'==============================================================
'Function: Menu_Count(InList,InComp,InType)
'==============================================================
Function Menu_Count(InList,InComp,InType)
Dim i,Cnt
	On Error Resume Next
	Err.Clear
	
	Cnt = 0	
	If IsArray(InList) Then
		For i = 0 To Ubound(InList)
			Select Case InType
				Case "DISPLAYFLAG"
					If InLIst(i).DISPLAYFLAG = InComp Then
						Cnt = Cnt + 1
					End If
				Case "GROUP"
					If InLIst(i).GROUP = InComp Then
						Cnt = Cnt + 1
					End If
			End Select
		Next
	Else
		Cnt = -1
	End If
	
	Menu_Count = (Cnt - 1)
End Function
'==============================================================
'Function: Menu_Return(InList,InComp,InType)
'==============================================================
Function Menu_Return(InList,InComp,InType)
Dim i,j,Cnt,TempArr
	On Error Resume Next
	Err.Clear
	
	Cnt = Menu_Count(InList,InComp,InType)
	
	If Cnt <> -1 Then
		ReDim TempArr(Cnt)
		j = 0
		For i = 0 To Ubound(InList)
			Select Case InType
				Case "DISPLAYFLAG"				
					If InList(i).DISPLAYFLAG = InComp Then					
						TempArr(j) = InList(i).ID
						j = j + 1
					End If
				Case "GROUP"				
					If InList(i).GROUP = InComp Then
						TempArr(j) = i						
						j = j + 1
					End If
			End Select
		Next	
	End If
	
	Menu_Return = TempArr
End Function
'==============================================================
'Function: Str_Split(InSrt,InComp)
'==============================================================
Function Str_Split(InStr,InComp)
Dim OutArr,OutStr
	On Error Resume Next
	Err.Clear 
	
	If Len(InStr) > 0 And Len(InComp) > 0 Then					
		OutStr = Left(InStr,Len(InStr)-Len(InComp))	
		If OutStr <> "" Then
			OutArr = Split(OutStr,InComp)
		End If
	End If
	
Str_Split = OutArr
End Function

'========================= MENU INIT 처리 ========================
'==============================================================
'Function: InitMenu(InMenu,InClass)
'==============================================================
Sub Menu_Init(InMenu,InClass)
Dim TempArr,i
	On Error Resume Next
	Err.Clear 

	TempArr = Str_Split(InMenu,GROW)
	If IsArray(TempArr) Then
		Select Case InClass
		Case MCLASS		
			ReDim TopMain(Ubound(TempArr))
			Call MenuSet_Init(TopMain,TempArr,InClass)
		Case SCLASS
			ReDim TopSub(Ubound(TempArr))
			Call MenuSet_Init(TopSub,TempArr,InClass)			
		Case LCLASS
			ReDim LeftMenu(Ubound(TempArr))
			Call MenuSet_Init(LeftMenu,TempArr,InClass)
		End Select			
	End If
End Sub
'========================= Menu 처리 ========================
'==============================================================
'Function: SlipMenu(InList,InProType)
'==============================================================
Function SlipMenu(InList,InProType)
Dim i,TempArr,NIDx,PIDx,LIDx,FIDx
	On Error Resume Next
	Err.Clear 
	
	If  SKIPMENUCNT - 1 >= 0 Then
		TempArr = Menu_Return(InList,True,"DISPLAYFLAG")	
		PIDx = Menu_Search(InList,TempArr(Lbound(TempArr) + 1),"MENUIDX")
		NIDx = Menu_Search(InList,TempArr(Ubound(TempArr) - 1),"MENUIDX")
		FIDx = Menu_Search(InList,TempArr(Lbound(TempArr)),"MENUIDX")
		LIDx = Menu_Search(InList,TempArr(Ubound(TempArr)),"MENUIDX")
		If IsArray(TempArr) Then
			Select Case InProType
			Case "AS"			
				If InList(LIDx).MOVERFLAG = False Then
					InList(LIDx).MOVERFLAG = True
				End If				
				For i = 0 To SKIPMENUCNT - 1
					If PIDx - 1 - i >= Lbound(InList) + 1 Then
						InList(PIDx - 1 - i).DISPLAYFLAG = True
						InList(NIDx - i).DISPLAYFLAG = False
						If PIDx - 1 - i = Lbound(InList) + 1 Then
							If InList(FIDx).MOVERFLAG = True Then
								InList(FIDx).MOVERFLAG = False
								Call MouseOut_Menu(FIDx,InList,MCLASS)
							End If
						End If
					End If				
				Next
			Case "AE"
				If InList(FIDx).MOVERFLAG = False Then
					InList(FIDx).MOVERFLAG = True
				End If				
				For i = 0 To SKIPMENUCNT - 1
					If NIDx + 1 + i <= Ubound(InList) - 1 Then
						InList(PIDx + i).DISPLAYFLAG = False						
						InList(NIDx + 1 + i).DISPLAYFLAG = True						
						If NIDx + 1 + i = Ubound(InList) - 1 Then
							If InList(LIDx).MOVERFLAG = True Then
								InList(LIDx).MOVERFLAG = False
								Call MouseOut_Menu(LIDx,InList,MCLASS)
							End If
						End If
					End If				
				Next
			End Select
		End IF
	End If
	
	SlipMenu = True
End Function
'==============================================================
'Function : Close_Menu(InList,InIDx,InClass)
'==============================================================
Function Close_Menu(InList,InIDx,InClass)
Dim i,TempArr,SubID,IDx
	On Error Resume Next
	Err.Clear
	
	If IsArray(InList) Then
		If InIDx <> "" Then
			If InList(InIDx).PROTYPE = "AS" Or InList(InIDx).PROTYPE = "AE" Then		
				Call MouseOut_Menu(InIDx,InList,InClass)
			End If
		End If
		
		For i = 0 To Ubound(InList)		
			If InList(i).OPENFLAG = True And InList(i).MOVERFLAG = True Then
				Call MouseOut_Menu(i,InList,InClass)
			End If
		Next
	End If
	
	Close_Menu = True
End Function
'==============================================================
'Function : Click_Menu(InIDx,InList,InClass)
'==============================================================
Function Click_Menu(InIDx,InList,InClass)
Dim CurrObj,IDx,TempLeftID
	On Error Resume Next
	Err.Clear 

		IDx = Menu_Search(TopMain,True,"CLICKIDX")		
		Call Click_CloseMenu(IDx,TopMain,MCLASS)	
		IDx = Menu_Search(TopSub,True,"CLICKIDX")	
		Call Click_CloseMenu(IDx,TopSub,SCLASS)
		IDx = Menu_Search(LeftMenu,True,"CLICKIDX")	
		Call Click_CloseMenu(IDx,LeftMenu,LCLASS)

		If InList(InIDx).CLICKFLAG = False Then						
			Call Click_OpenMenu(InIDx,InList,InClass)
		End If
		Select Case InClass		
	
		Case SCLASS	
			IDx = Menu_Search(LeftMenu,InList(InIDx).ID & LEFTID,"MENUIDX")
			If IDx <> -1 Then						
				Call Click_OpenMenu(IDx,LeftMenu,LCLASS)
				TempLeftID = Replace(TopSub(InIDx).GROUP,SUBNAME,LEFTNAME)
				Call Menu_Display(LCLASS,TempLeftID)
			End If	
		Case LCLASS							
			IDx = Menu_Search(TopSub,Replace(InList(InIDx).ID,LEFTID,""),"MENUIDX")
			If IDx <> -1 Then			
				Call Click_OpenMenu(IDx,TopSub,SCLASS)						
			End If
		End Select
		CurrURL = InList(InIDx).URL
		Set CurrObj = nothing
	Click_Menu = True
End Function
'==============================================================
'Function : Click_OpenFrame(inPagevalue)
'==============================================================
Function Click_OpenFrame(inPagevalue)
Dim CurrObj,InIDx,IDx
on Error Resume Next
Err.Clear 

    CurrURL = UCase(inPagevalue)
	InIDx = Menu_Search(LeftMenu,CurrURL,"URLIDX")		
	If InIDx <> -1 Then		
		    IDx = Menu_Search(TopMain,True,"CLICKIDX")		
		    Call Click_CloseMenu(IDx,TopMain,MCLASS)	
		    IDx = Menu_Search(TopSub,True,"CLICKIDX")	
		    Call Click_CloseMenu(IDx,TopSub,SCLASS)
		    IDx = Menu_Search(LeftMenu,True,"CLICKIDX")	
		    Call Click_CloseMenu(IDx,LeftMenu,LCLASS)

		    If LeftMenu(InIDx).CLICKFLAG = False Then						
	            Set CurrObj = document.all(LeftMenu(InIDx).ID)
	            CurrObj.style.color	= CLICK_COLOR
	            LeftMenu(InIDx).CLICKFLAG   = True
	            LeftMenu(InIDx).OPENFLAG	= True
	            LeftMenu(InIDx).MOVERFLAG   = False
	            Set CurrObj = nothing				            
		    End If
		    IDx = Menu_Search(TopSub,Replace(LeftMenu(InIDx).ID,LEFTID,""),"MENUIDX")		    
		    If IDx <> -1 Then			
		    	Set CurrObj = document.all(TopSub(IDx).ID)
	            CurrObj.style.color	= CLICK_COLOR
	            TopSub(IDx).CLICKFLAG   = True
	            TopSub(IDx).OPENFLAG    = True
	            TopSub(IDx).MOVERFLAG   = False
	            Set CurrObj = nothing			
		    End If			

		txtTitle.value = LeftMenu(InIDx).MTitle
    	document.title = gLogoName & " - " & LeftMenu(InIDx).MTitle & " [ " & "<%=NAME%>" & " ]"
	End If	
	
Click_OpenFrame = True
End Function
'==============================================================
'Function : Click_OpenMenu(InIDx,InList,InClass)
'==============================================================
Function Click_OpenMenu(InIDx,InList,InClass)
	Dim CurrObj
	
	Set CurrObj = document.all(InList(InIDx).ID)
	CurrObj.style.color	= CLICK_COLOR
	InList(InIDx).CLICKFLAG = True
	InList(InIDx).OPENFLAG	= True
	InList(InIDx).MOVERFLAG = False
	txtTitle.value = InList(InIDx).MTitle
	document.title = gLogoName & " - " & InList(InIDx).MTitle & " [ " & "<%=NAME%>" & " ]"
	Set CurrObj = nothing	
    Call SetToolBar("0000")
	if oldGroup<>"" then
		if oldGroup <> mid(InList(InIDx).GROUP,1,instr(1,InList(InIDx).GROUP,"_")-1) then
			txtEmp_no2.value = txtemp_no.value
			txtName2.value   = txtname.value
		end if
	end if
	oldGroup = mid(InList(InIDx).GROUP,1,instr(1,InList(InIDx).GROUP,"_")-1)
	
	Call FncPgmMenu1(InList(InIDx).URL,InIDx,InList)
	Click_OpenMenu = True
End Function
'==============================================================
'Function : Click_CloseMenu(InIDx,InList,InClass)
'==============================================================
Function Click_CloseMenu(InIDx,InList,InClass)
Dim CurrObj
	If InIDx <> -1 Then
		Set CurrObj = document.all(InList(InIDx).ID)
		Select Case InClass
			Case MCLASS
				CurrObj.style.color	= MAIN_COLOR
			Case SCLASS
				CurrObj.style.color	= SUB_OUT_COLOR
			Case LCLASS
				CurrObj.style.color	= LEFT_OUT_COLOR
		End Select
		InList(InIDx).CLICKFLAG	= False
		InList(InIDx).OPENFLAG	= False
		InList(InIDx).MOVERFLAG	= True
		Set CurrObj = nothing
	End If

Click_CloseMenu = True
End Function
'==============================================================
'Function : MouseOver_Menu(InIDx,InList,InClass)
'==============================================================
Function MouseOver_Menu(InIDx,InList,InClass)
Dim CurrObj,SubID,IDx,TempObj
	On Error Resume Next
	Err.Clear 	
	
	Set CurrObj	= document.all(InList(InIDx).ID)
	Select Case inClass
		Case MCLASS
			SubID = InList(InIDx).ID & SUBNAME
			IDx = Menu_Search(TopSub,SubID,"GROUPIDX")
			If IDx <> -1 Then
				Call Menu_Display(SCLASS,SubID)
			Else
				Call Menu_Display(SCLASS,"")
			End If
		Case SCLASS			
			If InList(InIDx).MOVERFLAG Then
				CurrObj.style.color				= OVER_COLOR
			Else
			End If			
			Set TempObj = Document.all(Replace(InList(InIDx).GROUP,SUBNAME,LEFTNAME))
			If UCase(TempObj.style.visibility)	= "VISIBLE" Then
				IDx = Menu_Search(LeftMenu,InList(InIDx).ID & LEFTID,"MENUIDX")				
				If IDx <> -1 Then
					Call MouseOver_Menu(IDx,LeftMenu,LCLASS)
				End If
			End If
			Set TempObj =  Nothing
		Case LCLASS			
			If InList(InIDx).MOVERFLAG Then
				CurrObj.style.color				= OVER_COLOR
			End If
	End Select
	If InList(InIDx).MOVERFLAG Then
		CurrObj.style.cursor					= OVER_CURSOR		
	Else
	End If
	InList(InIDx).OPENFLAG = True	
	Set CurrObj	= nothing
	
	MouseOver_Menu = True
End Function
'==============================================================
'Function : MouseOver_Menu(InIDx,InList,InClass)
'==============================================================
Function MouseOut_Menu(InIDx,InList,InClass)
Dim CurrObj,SubID,IDx,TempLeftID,TempObj
	On Error Resume Next
	Err.Clear 
	
	Set CurrObj	= document.all(InList(InIDx).ID)
	Select Case InClass
		Case MCLASS			
			Call Menu_Display(SCLASS,"")			
		Case SCLASS			
			If InList(InIDx).MOVERFLAG Then
				CurrObj.style.color		= SUB_OUT_COLOR
			Else
			End If
			Set TempObj = Document.all(Replace(InList(InIDx).GROUP,SUBNAME,LEFTNAME))
			If UCase(TempObj.style.visibility)	= "VISIBLE" Then
				IDx = Menu_Search(LeftMenu,InList(InIDx).ID & LEFTID,"MENUIDX")
				If IDx <> -1 Then
					Call MouseOut_Menu(IDx,LeftMenu,LCLASS)
				End If
			End If
			Set TempObj =  Nothing
		Case LCLASS
			If InList(InIDx).MOVERFLAG Then
				CurrObj.style.color			= LEFT_OUT_COLOR
			End If
	End Select	
	If InList(InIDx).MOVERFLAG Then
		CurrObj.style.cursor			= OUT_CURSOR		
	Else
	End If
	InList(InIDx).OPENFLAG = False	
	Set CurrObj	= nothing
	
	MouseOut_Menu = True	
End Function
'==============================================================
'Function : MenuSet_Init(InMenu,InList,InClass)
'==============================================================
Sub MenuSet_Init(InMenu,InList,InClass)
Dim i,j,TempArr,PMenu,NMenu,MFlag,GObj
	On Error Resume Next
	Err.Clear 
	
	j = 0
	PMenu = ""
	NMenu = ""
	MFlag = False	
	For i = 0 To Ubound(InList)	
		TempArr = Str_Split(InList(i),GCOL)		
		If IsArray(TempArr) Then
			Set InMenu(i) = New Menu			
			InMenu(i).URL		= UCase(TempArr(2))
			InMenu(i).OPENFLAG	= False
			InMenu(i).PROTYPE   = TempArr(5)
			InMenu(i).CLICKFLAG	= False
			InMenu(i).MTitle = TempArr(1)
			Select Case TempArr(5)
			Case "MM"
				InMenu(i).NEXTFLAG	= True
			Case Else
				InMenu(i).NEXTFLAG	= False
			End Select
			
			Select Case InClass
			Case MCLASS	
				InMenu(i).ID		= TempArr(0)
				InMenu(i).GROUP	= TOPMENUBAR
				Select Case TempArr(5)
				Case "AS"
					If Ubound(InList) - 1 > VIEWMENUCNT Then
						InMenu(i).MOVERFLAG		= False
						InMenu(i).DISPLAYFLAG	= True						
					Else
						InMenu(i).MOVERFLAG		= False
						InMenu(i).DISPLAYFLAG	= False						
					End If
					InMenu(i).TOPEND	= -1
				Case "AE"
					If Ubound(InList) - 1 > VIEWMENUCNT Then
						InMenu(i).MOVERFLAG		= True
						InMenu(i).DISPLAYFLAG	= True
					Else
						InMenu(i).MOVERFLAG		= False
						InMenu(i).DISPLAYFLAG	= False
					End If
					InMenu(i).TOPEND	= -2
				Case Else
					If Ubound(InList) - 1 > VIEWMENUCNT Then
						InMenu(i).MOVERFLAG = True
						If i <= VIEWMENUCNT Then
							InMenu(i).DISPLAYFLAG = True							
						Else
							InMenu(i).DISPLAYFLAG = False							
						End If
						
						Select Case i
						Case Lbound(InList) + 1
							InMenu(i).TOPEND	= MENUTOP
						Case VIEWMENUCNT
							InMenu(i).TOPEND	= MENUEND
						Case Else
							InMenu(i).TOPEND	= i - 1
						End Select
					Else
						InMenu(i).MOVERFLAG		= True
						InMenu(i).DISPLAYFLAG	= True
						Select Case i
						Case Lbound(InList)
							InMenu(i).TOPEND	= MENUTOP
						Case Ubound(InList)
							InMenu(i).TOPEND	= MENUEND
						Case Else
							InMenu(i).TOPEND	= i
						End Select
					End If
				End Select
			Case SCLASS
				InMenu(i).ID		= TempArr(0)
				Set GObj = document.all(InMenu(i).ID).parentElement
				InMenu(i).GROUP			= GObj.ID
				InMenu(i).MOVERFLAG		= True
				InMenu(i).DISPLAYFLAG	= False				
				PMenu = TempArr(4)
				If PMenu <> NMenu Then
					j = 0
					InMenu(i).TOPEND		= MENUTOP
					If NMenu <> "" Then
						InMenu(i-1).TOPEND	= MENUEND
					End If
					NMenu = PMenu
				Else
					If i = Ubound(InList) Then
						InMenu(i-1).TOPEND	= MENUEND
					Else
						InMenu(i).TOPEND	= j
					End If
				End If
				j = j + 1
				Set GObj = nothing
			Case LCLASS				
				InMenu(i).ID		= TempArr(0) & LEFTID				
				Set GObj = document.all(InMenu(i).ID).parentElement							
				InMenu(i).GROUP			= GObj.ID
				InMenu(i).DISPLAYFLAG	= False
				Select Case TempArr(5)
				Case "MM"
					InMenu(i).MOVERFLAG = False
					InMenu(i).TOPEND	= -1					
					If i - 1 >= 0 Then
						InMenu(i-1).TOPEND	= MENUEND
					End If					
					j = 0
				Case "MP"
					InMenu(i).MOVERFLAG = True
					InMenu(i).TOPEND	= MENUEND
					j = 0
				Case Else
					InMenu(i).MOVERFLAG = True
					If j = 1 Then
						InMenu(i).TOPEND	= MENUTOP
					ElseIf i = Ubound(InList) Then
						InMenu(i).TOPEND	= MENUEND
					Else
						InMenu(i).TOPEND	= j - 1
					End If					
				End Select
				j = j + 1
				Set GObj = nothing
			End Select			
		End If
	Next	
End Sub
'==============================================================
'Function : Menu_Display(InClass,InList)
'==============================================================
Function Menu_Display(InClass,InList)
Dim i,IDx,HIDx,TempArr1,TempArr2,TempObj,PrevObj,LastObj,FirstObj
Dim ParentObj,GroupObj

	On Error Resume Next
	Err.Clear 

	Select Case InClass
	Case MCLASS		
		TempArr1 = Menu_Return(InList,True,"DISPLAYFLAG")
		If IsArray(TempArr1) Then
			For i = 0 To Ubound(TempArr1)			
				Set TempObj = document.all(TempArr1(i))
				If i = 0 Then
					If TempObj.offsetLeft < MENUYSPACE Then
						TempObj.style.Left		= TempObj.offsetLeft + MENUYSPACE
					Else
						TempObj.style.Left		= TempObj.offsetLeft
					End If
					Call GUBUN_Init(TempObj.nextSibling,"Visible","VISIBLE")									
					Call GUBUN_Init(TempObj.nextSibling,TempObj.offsetLeft + TempObj.offsetWidth,"LEFT")
					TempObj.style.Height	= TempObj.parentElement.offsetHeight					
				Else
					Set PrevObj = document.all(TempArr1(i - 1))					
					TempObj.style.Left		= PrevObj.offsetLeft + PrevObj.offsetWidth + MENUXSPACE 
					Call GUBUN_Init(TempObj.nextSibling,"Visible","VISIBLE")
					Call GUBUN_Init(TempObj.nextSibling,TempObj.offsetLeft + TempObj.offsetWidth,"LEFT")					
					TempObj.style.Height	= PrevObj.parentElement.offsetHeight
				End If						
				TempObj.style.visibility	= "Visible"
			Next
		End If
		If IsArray(TempArr1) Then
			IDx = Menu_Search(InList,TempArr1(Lbound(TempArr1)),"MENUIDX")
			If IDx <> -1 Then
				If InList(IDx).PROTYPE = "AS" Then
					Set FirstObj = document.all(InList(IDx).ID)
					Call GUBUN_Init(FirstObj.nextSibling,"Hidden","VISIBLE")					
					Set FirstObj = Nothing
				End If
			End IF		
			HIDx = Menu_Search(InList,TempArr1(Ubound(TempArr1)),"MENUIDX")
			If InList(HIDx).PROTYPE = "AE" Then
				Set LastObj = document.all(TempArr1(Ubound(TempArr1) - 1))
				Call GUBUN_Init(LastObj.nextSibling,"Hidden","VISIBLE")
				Set LastObj = Nothing				
			End If
			Set LastObj = document.all(InList(HIDx).ID)
			Call GUBUN_Init(LastObj.nextSibling,"Hidden","VISIBLE")
			If isObject(LastObj.nextSibling) Then
				LastObj.nextSibling.style.visibility	= "Hidden"
			End If
			Set LastObj = Nothing			
		End If
		
		TempArr2 = Menu_Return(InList,False,"DISPLAYFLAG")				
		If IsArray(TempArr2) Then		
			For i = 0 To Ubound(TempArr2)			
				Set TempObj = document.all(TempArr2(i))			
				TempObj.style.visibility	= "Hidden"
				Call GUBUN_Init(TempObj.nextSibling,"Hidden","VISIBLE")
				Set TempObj = Nothing
			Next
		End If		
	Case SCLASS		
		If  COpenSub <> "" And (COpenSub <> InList Or InList = "")  Then
			Set PrevObj = document.all(COpenSub)
			PrevObj.style.visibility	= "Hidden"
			Set PrevObj = nothing		
		End If
		If InList <> "" Then
			Set TempObj = document.all(InList)
			If UCase(TempObj.style.visibility) <> "VISIBLE" Then			
				Set GroupObj = document.all(InList).parentElement
				Set ParentObj = document.all(Replace(InList,SUBNAME,""))
				TempObj.style.Left = ParentObj.offsetLeft - TempObj.offsetWidth/2 + ParentObj.offSetWidth/2							
				If TempObj.offsetLeft < GroupObj.offsetLeft + MENUYSPACE Then				
					TempObj.style.Left = GroupObj.offsetLeft + MENUYSPACE
				ElseIf (TempObj.offsetLeft + TempObj.offsetWidth) > (GroupObj.offsetLeft + GroupObj.offsetWidth) Then
					TempObj.style.Left = TempObj.offsetLeft - ((TempObj.offsetLeft + TempObj.offsetWidth) - (GroupObj.offsetLeft + GroupObj.offsetWidth)) - MENUYSPACE
				End If
				TempObj.style.Top = ParentObj.offsetTop + MENUYSPACE - 5
				TempObj.style.visibility	= "Visible"
				Set GroupObj = nothing
				Set ParentObj = nothing		 
			End If
			Set TempObj = nothing
		End If
		COpenSub = InList		
	Case LCLASS	
		If  COpenLeft <> "" And (COpenLeft <> InList Or InList = "")  Then
			Set PrevObj = document.all(COpenLeft)
			PrevObj.style.visibility	= "Hidden"
			Set PrevObj = nothing		
		End If
		If InList <> "" Then
			Set TempObj = document.all(InList)
			If UCase(TempObj.style.visibility) <> "VISIBLE" Then
				Set ParentObj = TempObj.parentElement
				TempObj.style.Left = ParentObj.offsetLeft + MENUXSPACE
				TempObj.style.Top = ParentObj.offsetTop + MENUYSPACE
				TempObj.style.Width = LEFTMENUWIDTH - MENUXSPACE
				TempObj.style.visibility	= "Visible"				
				Set ParentObj = nothing
			End If
			Set TempObj = nothing
		End If
		COpenLeft = InList		
	End Select	
	Menu_Display = True
End Function
'==============================================================
'Function : GUBUN_Ini(InObj,InHav,InType)
'==============================================================
Function GUBUN_Init(InObj,InHav,InType)
	On Error Resume Next
	Err.Clear 
	
	If Not IsNull(InObj.className) Then
		Select Case InType
			Case "LEFT"				
				InObj.style.Left	= InHav
			Case "VISIBLE"				
				InObj.style.visibility	= InHav
		End Select
	End If
	GUBUN_Init = True
End Function

Function FncPgmMenu(inPagevalue)

	document.All("DivPgmMenu").style.POSITION = "absolute"
    document.All("divHomeMenu").style.VISIBILITY = "hidden"
    document.All("DivPgmMenu").style.VISIBILITY = "visible"
    Call formmenu_onLoad(inPagevalue)  
	document.All("formmenu").src = "./" & inPagevalue
End Function

Function FncPgmMenu1(inPagevalue,InIDx,InList)
	document.All("DivPgmMenu").style.POSITION = "absolute"
    document.All("divHomeMenu").style.VISIBILITY = "hidden"
    document.All("DivPgmMenu").style.VISIBILITY = "visible"    
	document.All("formmenu").src = "./" & inPagevalue & "?strTitle=" & InList(InIDx).MTitle
End Function

Function FncHomeMenu()
Dim IDx	

	document.All("divHomeMenu").style.POSITION = "absolute"
	document.all("divHomeMenu").style.VISIBILITY = "visible"	

    Call SetToolBar("0000")
    document.all("DivPgmMenu").style.VISIBILITY = "hidden"
    document.All("nextprev").style.VISIBILITY = "hidden"
    IDx = Menu_Search(TopMain,True,"CLICKIDX")	
    If IDx <> -1 Then	
		Call Click_CloseMenu(IDx,TopMain,MCLASS)
	End If
    IDx = Menu_Search(LeftMenu,True,"CLICKIDX")
    If IDx <> -1 Then	
		Call Click_CloseMenu(IDx,LeftMenu,LCLASS)
		Call Menu_Display(LCLASS,"")
	End If
	IDx = Menu_Search(TopSub,True,"CLICKIDX")
    If IDx <> -1 Then
		Call Click_CloseMenu(IDx,TopSub,SCLASS)
		Call Menu_Display(SCLASS,"")
	End If
    Call Menu_Init(TopMain,MCLASS)		
	Call Menu_Init(TopSub,SCLASS)	
	Call Menu_Init(LeftMenu,LCLASS)
	document.title = gLogoName & " [ " & "<%=NAME%>" & " ]"
End Function

Function FncLogoff(Where)
	Dim IntRetCD,strPath
	lgFncLogoff = false
	If Where=1 Then '브라우저를 강제 종료시킬때 
	Else
		intRetCD = msgbox("대사우서비스를 종료하시겠습니까?", vbOKCancel,"대사우서비스")
	    If IntRetCD<>1 Then Exit Function 
	    
	End If
   
    txtemp_no.value = ""
    txtname.value = ""
'    txtpassword.value = ""
    txtinternal_cd.value = ""
    txtnat_cd.value = ""
    txtDEPT_AUTH.value = ""
    txtPRO_AUTH.value = ""
    txtLang.value = ""
    
    'strPath = GetHomePath & "/unisims.asp"

    window.document.location = "./e1logoffmb1.asp?txtMode=UID_M0003"

    'window.document.location.href = strPath
    'Call SIMS.ExitProcess(trim(gLogoName), CStr(strPath))
   
	lgFncLogoff = True
End Function
'========================================================================================
' Function Name : GetUserPath
' Function Desc : 현재 디렉토리 패스 알아오기 
'========================================================================================
Function GetHomePath()
		Dim strLoc, iPos , iLoc, strPath
		strLoc = window.location.href
					iLoc = inStr(1, strLoc, "/")
  					iLoc = Cint(inStr(iLoc+1, strLoc, "/"))
  					iLoc = Cint(inStr(iLoc+1, strLoc, "/"))
  					iLoc = Cint(inStr(iLoc+1, strLoc, "/"))
  					iLoc = Cint(inStr(iLoc+1, strLoc, "/"))
            
                If iLoc > 0 Then
                   strLoc = Left(strLoc, iLoc - 1)
                End If
		
		iLoc = 1: iPos = 0
		Do Until iLoc <= 0						
			iLoc = inStr(iPos+1, strLoc, "/")
			If iLoc <> 0 Then iPos = iLoc
		Loop	
		GetHomePath = strLoc
End Function


Function FncPassword(pParm)

	Dim arrRet
	Dim arrParam(2)
	
    if pparm = 1 then
	    arrRet = window.showModalDialog("EchangePW.asp", Array(arrParam), _
	    	"dialogWidth=395px; dialogHeight=200px; center: Yes; help: No; resizable: No; status: No;")
    else		
	    arrRet = window.showModalDialog("EchangePWFirst.asp", Array(arrParam), _
	    	"dialogWidth=400px; dialogHeight=250px; center: Yes; help: No; resizable: No; status: No;")
    end if
End Function

Function FncHelp()
	dim from_GetProgId
	from_GetProgId= formmenu.GetProgId()
    If from_GetProgId = "" Then 
        window.open "../../Help/Ess/ess.htm","ESShelp","status=no,toolbar=no,menubar=no,height=430,width=845,center=yes,top=45,left=100"
    Else
        window.open "../../Help/Ess/ess.htm?path=esshelp/" & from_GetProgId & ".htm" ,"ESShelp","status=no,toolbar=no,menubar=no,height=430,width=845,center=yes,top=45,left=100"
    End If

End Function

Function FncQuery()

    call formmenu.DbQuery(1)

End Function

Function FncSave()

    call formmenu.DbSave()

End Function

Function FncAdd()

    call formmenu.FncNew()

End Function

Function FncDel()

    call formmenu.DbDelete()

End Function

Function FncNext()
	On Error Resume Next
    call formmenu.FncNext()
End Function

Function FncPrev()
	On Error Resume Next
    call formmenu.FncPrev()
End Function

Function FncPrint()

    formmenu.focus()
    call formmenu.Print()

End Function

'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line 
'========================================================================================================
Function OpenEmp(pEmpNo)
	Dim arrRet
	Dim arrParam(2)
	Dim iWhereFlg

	If OpenEmp = True Then Exit Function
	OpenEmp = True

	arrParam(0) = txtEmp_no2.value			' Code Condition
	arrParam(1) = txtName2.value			' Name Cindition
	
    If inStr(1,UCase(formmenu.document.location),"E16",1)>0 or inStr(1,UCase(formmenu.document.location),"E17",1)>0  Then
        iWhereFlg = True
      
        arrParam(2) = Trim(txtEmp_no.Value)     ' 근태관리 담당자일 경우 
	    arrRet = window.showModalDialog("E1EmpPopa3.asp", Array(arrParam), _
    		"dialogWidth=540px; dialogHeight=385px; center: Yes; help: No; resizable: No; status: No;")
    Else
        iWhereFlg = False
        
        arrParam(2) = Trim(txtinternal_cd.Value)' lgUsrIntCd
	    arrRet = window.showModalDialog("E1EmpPopa1.asp", Array(arrParam), _
	    	"dialogWidth=540px; dialogHeight=385px; center: Yes; help: No; resizable: No; status: No;")
	End If
		
	OpenEmp = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
	    txtemp_no2.value = arrRet(0)
	    txtname2.value = arrRet(1)
	    If iWhereFlg = True Then 
	        formmenu.document.frm1.txtemp_no.value = arrRet(0)
	        formmenu.document.frm1.txtName.value = arrRet(1)
	        formmenu.document.frm1.txtroll_pstn.value = arrRet(2)
	        formmenu.document.frm1.txtDept_nm.value = arrRet(3)
	    End If
	End If	
			
End Function

</Script>
<Script language="vbscript" Runat="Server">
'==============================================================
'Function: Str_MidLeft(InStr,InComp)
'==============================================================
Function Str_MidLeft(InStr,InComp)
Dim OutStr
	Err.Clear 
	
	If Len(InStr) > 0 And Len(InComp) > 0 Then	

	End If
	
Str_MidLeft = OutStr
End Function
'==============================================================
'Function: Str_Split(InSrt,InComp)
'==============================================================
Function Str_Split(InStr,InComp)
Dim OutArr,OutStr
	On Error Resume Next
	Err.Clear 
	
	If Len(InStr) > 0 And Len(InComp) > 0 Then
	    If	Right(InStr,Len(InComp)) = InCOmp Then
		    OutStr = Left(InStr,Len(InStr)-Len(InComp))		
		Else
		    OutStr = Left(InStr,Len(InStr))
		End If
		If OutStr <> "" Then
			OutArr = Split(OutStr,InComp)
		End If
	End If
	
Str_Split = OutArr
End Function
'==============================================================
'Function: Level_MenuReturn(InArr,IDx,InLevel)
'==============================================================
Function Level_MenuReturn(InArr,IDx,InLevel)
Dim i,j,TempArr,TempList,OutArr
	On Error Resume Next
	Err.Clear 
	
	If IsArray(InArr) Then
		For i = 0 To Ubound(InArr)
			TempArr = Str_Split(InArr(i),GCOL)
			If IsArray(TempArr) Then
				If TempArr(IDx) = InLevel Then
					For j = 0 To Ubound(TempArr)
						TempList = TempList & TempArr(j) & GCOL
					Next
					TempList = TempList & GROW
				End If
			End If			
		Next 
	End If	
	If TempList <> "" Then
		OutArr = Str_Split(TempList,GROW)			
	End If
	Level_MenuReturn = OutArr
End Function


</Script>

<script language="JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i]) &&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}

}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n]) &&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->

</script>
<!-- #Include file="../../inc/uniSimsClassID.inc" -->
</HEAD>
<body bgcolor="#ffffff" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" onLoad="javascript:MM_preloadImages('../../../CShared/image/uniSIMS/beforenext_15.gif','../../../CShared/image/uniSIMS/buttonover_22.jpg','../../../CShared/image/uniSIMS/buttonover_23.jpg','../../../CShared/image/uniSIMS/buttonover_24.jpg','../../../CShared/image/uniSIMS/buttonover_25.jpg')" >
<%
'=========================메뉴 String 생성=======================
Dim TempArr,TempArr1,TempList,MenuStr,i,j
Dim TopArrow,TopMain,TopSub,TopMenu,LeftMenu
Dim PMenu,NMenu
Dim StrMain,StrSub,TempArr2,k

	Call SubOpenDB(lgObjConn)
	lgStrSQL = "SELECT MENU_ID,MENU_NAME,HREF,MENU_LEVEL,REF_MENU_ID,PRO_TYPE FROM E11000T"
	lgStrSQL = lgStrSQL & " WHERE PRO_AUTH >=  " & FilterVar(gProAuth , "''", "S") & " "
	lgStrSQL = lgStrSQL & " AND LANG_CD =  " & FilterVar(gLang , "''", "S") & " "
	lgStrSQL = lgStrSQL & " AND PRO_USE_FLAG = " & FilterVar("Y", "''", "S") & " "
	lgStrSQL = lgStrSQL & " ORDER BY ref_menu_id,orders"				

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        Response.End
    Else
		Do While Not lgObjRs.EOF
		MenuStr = MenuStr & lgObjRs(0) & GCOL & lgObjRs(1) & GCOL & lgObjRs(2) & GCOL
		MenuStr = MenuStr & lgObjRs(3) & GCOL & lgObjRs(4) & GCOL & lgObjRs(5) & GCOL & GROW		
		lgObjRs.MoveNext
		Loop
	End If	
	Call SubCloseRs(lgObjRs)
	Call SubCloseDB(lgObjConn)
	
	TempArr		= Str_Split(MenuStr,GROW)	
	TopArrow	= Level_MenuReturn(TempArr,3,"0")
	TopMain		= Level_MenuReturn(TempArr,3,"1")
	
	If IsArray(TopMain) And IsArray(TopArrow) Then
		TempList = TempList & TopArrow (Lbound(TopArrow)) & GROW
		TempList = TempList & Join(TopMain,GROW)
		TempList = TempList & GROW & TopArrow (Ubound(TopArrow)) & GROW
	End If

	If TempList <> "" Then
		TopMenu		= Str_Split(TempList,GROW)
	End If
	
	TopSub		= Level_MenuReturn(TempArr,3,"2")
	
	If IsArray(TopMain) Then
		TempList = ""
		For i = 0 To Ubound(TopMain)
			TempArr		= Str_Split(TopMain(i),GCOL)
			TempList	= TempList & TopMain(i) & GROW	
			TempArr1	= Level_MenuReturn(TopSub,4,TempArr(0))			
			If IsArray(TempArr1) Then				
				TempList	= TempList & Join(TempArr1,GROW)
				TempList	= TempList & GROW
			End If	
			
		Next		
		If TempList <> "" Then
			LeftMenu	= Str_Split(TempList,GROW)
		End If
	End If
'==============================================================================
'** Lenth,mid,left function (Koean character is 2byte)
'==============================================================================
'**** calculate length of character (Koean character is 2byte)
Function Len2(AllText) 
    Dim nLen 
    Dim nCnt 
    Dim szEach 

    nLen = 0 
    AllText = Trim(AllText) 
    For nCnt = 1 To Len(AllText) 

            szEach = Mid(AllText,nCnt,1) 
            If 0 <= Asc(szEach) And Asc(szEach) <= 255 Then 
                    nLen = nLen + 1             '한글이 아닌 경우 
            Else 
                    nLen = nLen + 2             '한글인 경우 
            End If 
    Next 

    Len2 = nLen 
End Function 

'**** mid function (Koean character is 2byte)
Function Mid2(s, start, length) 
	Dim i, CharAt, VBLength, VBn1, VBn2, BLength, AddByte 
	VBn2=length 
	VBLength=Len(s) 
	BLength=0 
	for i=1 to VBLength 
		CharAt=mid(s, i, 1) 
		if asc(CharAt)>0 and asc(CharAt)<255 then 
			BLength=BLength + 1 
		else 
			BLength=BLength + 2 
		end if 
		If BLength>=start Then 
			Exit For 
		End If 
	next 

	VBn1=i 
	If VBn1<1 Then VBn1=1 
	BLength=0 
	for i=VBn1 to VBLength 
		CharAt=mid(s, i, 1) 
		if asc(CharAt)>0 and asc(CharAt)<255 then 
			BLength=BLength + 1 
		else 
			BLength=BLength + 2 
		end if 
		If BLength=length Then 
			VBn2=i+1 
			Exit For 
		ElseIf BLength>length Then 
			VBn2=i 
			Exit For 
		End If 
	next 
	Mid2=Mid(s, VBn1, VBn2-VBn1) 
End Function 

'**** Left function (Koean character is 2byte)
Function Left2(s, size) 
	Left2=Mid2(s, 1, size) 
End Function 	
'=======================메뉴 생성===========================
'=======================Top Main ===========================
%>
<table width="1000" border="0" CELLPADDING="0" CELLSPACING="0">
	<tr>
		<td width="119" height="2"><img src="../../../CShared/image/uniSIMS/logopart_01.gif" width="119" height="25"></td>
		<td width="24" height="2"><img src="../../../CShared/image/uniSIMS/logopart_02.gif" width="24" height="25"></td>
		<td rowspan="2" colspan="2" background="../../../CShared/image/uniSIMS/skyback_03.gif"></td>
		<td rowspan="2" width="401"><img src="../../../CShared/image/uniSIMS/sky_05.jpg" width="479" height="53"></td>
		<td rowspan="2" width="140"><img src="../../../CShared/image/uniSIMS/sky_06.jpg" width="140" height="53"></td>
	</tr>
	<tr>
		<td width="119">
			<script language =javascript src='./js/emenu_ShockwaveFlash1_N284585011.js'></script>
		</td>
		<td width="24"><img src="../../../CShared/image/uniSIMS/logopart_09.gif" width="24" height="28"></td>
	</tr>
</table>
<table width="1000" border="0" cellpadding="0" cellspacing="0">
	<TR>
		<TD height="14" BACKGROUND="../../../CShared/image/uniSIMS/topback_101.jpg" WIDTH="532"></TD>
		<td height="14" background="../../../CShared/image/uniSIMS/topback_101.jpg" valign="top" width="31"></td>
		<td colspan="2" height="14" background="../../../CShared/image/uniSIMS/topback_101.jpg"></td>
		<td align="right" width="68"><img src="../../../CShared/image/uniSIMS/top_middle_131.jpg" width="77" height="14"></td>
		<td align="right" width="63"><img src="../../../CShared/image/uniSIMS/top_middle_141.jpg" width="63" height="14"></td>
	</TR>
<%
	Response.Write "<TR>" & vbCrLf
	Response.Write "<TD height=25 BACKGROUND=../../../CShared/image/uniSIMS/topback_102.jpg WIDTH=532 LEVEL='" & MENUAREA & "'>" & vbCrLf
	Response.Write "<DIV ID='" & TOPMENUBAR & "' CLASS='" & TOPMENUBAR & "' LEVEL ='" & MENUAREA & "' NOWRAP>" & vbCrLf
	If IsArray(TopMenu) Then
		For i = 0 To Ubound(TopMenu)
			TempArr = Str_Split(TopMenu(i),GCOL)

			If IsArray(TempArr) Then
				If TempArr(5) = "AS" Then
					Response.Write "&nbsp;<SPAN><A HREF='" & TempArr(2) & "' ID='" & TempArr(0) & "' CLASS='" & MAINCLASS & "' LEVEL ='" & TempArr(3) & "'>"
					Response.Write "<img src=../../../CShared/image/arrow/left_arr.jpg border=0 onclick='vbscript:menu_move(""AS"")'>"	
				else
					If TempArr(5) = "AE" Then
						Response.Write "&nbsp;<SPAN><A HREF='" & TempArr(2) & "' ID='" & TempArr(0) & "' CLASS='" & MAINCLASS & "' LEVEL ='" & TempArr(3) & "'>"
						Response.Write "<img src=../../../CShared/image/arrow/right_arr.jpg border=0 onclick='vbscript:menu_move(""AE"")'>"	
					Else
						Response.Write "&nbsp;<SPAN><A HREF='" & TempArr(2) & "' ID='" & TempArr(0) & "' CLASS='" & MAINCLASS & "' MTitle='" & TempArr(1) & "' LEVEL ='" & TempArr(3) & "'>"

					end if
				End If
				Response.Write TempArr(1) 				
				Response.Write "</A><SPAN CLASS='GUBUN'>|</SPAN></SPAN>" & vbCrLf		
			End If
		Next
	
		Response.Write "</DIV>" & vbCrLf
	End If
	Response.Write "</TD>" & vbCrLf
    Response.Write "<td height=25 background=../../../CShared/image/uniSIMS/topback_102.jpg valign=top width=31></td>" & vbCrLf
    Response.Write "<td colspan=2 valign=top height=25 background=../../../CShared/image/uniSIMS/topback_102.jpg style='FONT-SIZE: 10pt; PADDING-TOP: 1px;'align = right>" & vbCrLf
    Response.Write "</td>" & vbCrLf
    Response.Write "<td align=right width=68><img src=../../../CShared/image/uniSIMS/top_middle_132.jpg width=77 height=25></td>" & vbCrLf
    Response.Write "<td align=right width=63><img src=../../../CShared/image/uniSIMS/top_middle_142.jpg width=63 height=25></td>" & vbCrLf
    Response.Write "</TR>" & vbCrLf



'=======================Top SUB  ===========================	

	Response.Write "<TR>" & vbCrLf
	Response.Write "<TD colspan=3 height=39 BACKGROUND=../../../CShared/image/uniSIMS/topback_171.gif LEVEL='" & MENUAREA & "' width=714 >" & vbCrLf
	If IsArray(TopSub) Then
		PMenu = ""
		NMenu = ""
		For i = 0 To Ubound(TopSub)		
			TempArr = Str_Split(TopSub(i),GCOL)
			If IsArray(TempArr) Then	
				PMenu = TempArr(4)			
				If PMenu <> NMenu Then
					If NMenu <> "" Then
						Response.Write "</DIV>" & vbCrLf
						NMenu = ""					
					End If
					Response.Write  "&nbsp;<DIV ID='" & PMenu & SUBNAME & "' CLASS='" & TOPSUBBAR & "' LEVEL ='" & MENUAREA & "' NOWRAP>" & vbCrLf
					NMenu = PMenu				
				End If
				Response.Write "&nbsp;<SPAN HREF=" & TempArr(2) & " ID='" & TempArr(0) & "' CLASS=" & SUBCLASS & " MTitle='" & TempArr(1) & "' LEVEL ='" & TempArr(3) & "'>" & vbCrLf
				Response.Write "::" & TempArr(1)
				Response.Write "</SPAN>" & vbCrLf
			End If
		Next
		Response.Write "</DIV>" & vbCrLf
	End If
	Response.Write "</TD>" & vbCrLf
    Response.Write "<td height=39 background=../../../CShared/image/uniSIMS/topback_171.gif valign=bottom width=155 align=right>"

    Response.Write "<a href='vbscript:FncHomeMenu()' onMouseOut=javascript:MM_swapImgRestore() "
    Response.Write " onMouseOver=javascript:MM_swapImage('home','','../../../CShared/image/uniSIMS/buttonover_22.jpg',1)>"
    Response.Write "<img name=home border=0 src=../../../CShared/image/uniSIMS/button_22.jpg width=35 height=32 alt='HOME'></a>"
    Response.Write "<a href='vbscript:FncLogoff(2)' onMouseOut=javascript:MM_swapImgRestore()"
    Response.Write " onMouseOver=javascript:MM_swapImage('logout','','../../../CShared/image/uniSIMS/buttonover_23.jpg',1)>"
    Response.Write "<img name=logout border=0 src=../../../CShared/image/uniSIMS/button_23.jpg width=36 height=32 alt='로그오프'></a>"
    Response.Write "<a href='vbscript:FncPassword(1)' onMouseOut=javascript:MM_swapImgRestore()"
    Response.Write " onMouseOver=javascript:MM_swapImage('password','','../../../CShared/image/uniSIMS/buttonover_24.jpg',1)>"
    Response.Write "<img name=password border=0 src=../../../CShared/image/uniSIMS/button_24.jpg width=36 height=32 alt='패스워드변경'></a>"
    Response.Write "<a href='vbscript:FncHelp()' onMouseOut=javascript:MM_swapImgRestore()"
    Response.Write " onMouseOver=javascript:MM_swapImage('admin','','../../../CShared/image/uniSIMS/buttonover_25.jpg',1)>"
    Response.Write "<img name=admin border=0 src=../../../CShared/image/uniSIMS/button_25.jpg width=36 height=32 alt='HELP'></a></td>" & vbCrLf
    Response.Write "<td valign=top align=right width=68><img src=../../../CShared/image/uniSIMS/top_middle_181.jpg width=77 height=39></td>" & vbCrLf
    Response.Write "<td valign=top align=right width=63 ><img src=../../../CShared/image/uniSIMS/top_middle_right_191.jpg width=63 height=39></td>" & vbCrLf
	Response.Write "</TR>" & vbCrLf

	Response.Write "<TR>" & vbCrLf
    Response.Write "<TD WIDTH=532></TD>" & vbCrLf
    Response.Write "<TD WIDTH=31></TD>" & vbCrLf
    Response.Write "<TD WIDTH=151></TD>" & vbCrLf
    Response.Write "<TD WIDTH=155></TD>" & vbCrLf
    Response.Write "<TD WIDTH=140 colspan=2><img src=../../../CShared/image/uniSIMS/top_bottom_right.jpg width=140 height=17></TD>" & vbCrLf

    Response.Write "</TR>" & vbCrLf
    Response.Write "</TABLE>" & vbCrLf
'=======================Left Menu  ===========================
			
	For i = 0 To Ubound(LeftMenu)
		TempArr = Str_Split(LeftMenu(i),GCOL)
		If IsArray(TempArr) Then
			If Trim(TempArr(4)) = "" Then
				StrMain = StrMain & TempArr(1) & GCOL
			End If
		End If
	Next
	For i = 0 To Ubound(LeftMenu)
		TempArr = Str_Split(LeftMenu(i),GCOL)
		If IsArray(TempArr) Then
			If Trim(TempArr(4)) <> "" Then 
				StrSub = StrSub & TempArr(1) & GCOL & TempArr(2) & GCOL
			ElseIf Trim(TempArr(4)) = "" And i <> 0 Then
				StrSub = StrSub & GROW
			End If	
		End If
	Next
	
	TempArr = Str_Split(StrMain,GCOL)
	TempArr1 = Str_Split(StrSub,GROW)
    i = 0
    j = 0
    %>
		</table>
		<DIV ID="divHomeMenu" style="VISIBILITY: visible; POSITION: absolute">
			<table width="1000" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td height="21" colspan="15"></td>
				</tr>
				<TR>
					<td width="55"><img src="../../../CShared/image/uniSIMS/wall_top_left_32.gif" width="55" height="42"></td>
					<td background="../../../CShared/image/uniSIMS/wall_topback.gif" width="38"></td>
					<%
            If IsArray(TempArr) Then
				If i <= Ubound(TempArr) Then
					Response.Write "<TD background=../../../CShared/image/uniSIMS/titlebox_33.gif CLASS=MENU width=94 valign=top height=42>"
					Response.Write TempArr(i)				
					i = i + 1
				Else
					Response.Write "<td background=../../../CShared/image/uniSIMS/wall_topback.gif width=94 valign=top height=42>"
				End If
			End If
    %>
					</TD>
<%     for h=33 to 41 step 2%>
					<td background="../../../CShared/image/uniSIMS/wall_topback.gif" width="50" ></td>
					<%
            If IsArray(TempArr) Then
				If i <= Ubound(TempArr) Then
					Response.Write "<TD background=../../../CShared/image/uniSIMS/titlebox_" & h & ".gif CLASS=MENU width=94 valign=top height=42>"
					Response.Write TempArr(i)				
					i = i + 1
				Else
					Response.Write "<td background=../../../CShared/image/uniSIMS/wall_topback.gif width=94 valign=top height=42>"
				End If
			End If
    %>
					</TD>
<%       Next %>
					<td background="../../../CShared/image/uniSIMS/wall_topback.gif" width="37"></td>
					<td width="63" align="right"><img src="../../../CShared/image/uniSIMS/wall_top_right_45.jpg" width="63" height="42"></td>
				</TR>
				<TR height="150">
					<TD background="../../../CShared/image/uniSIMS/left_wall_71.gif" width="55"></TD>
					<TD width="38"></TD>
<%     for h=1 to 6 %>					
					<TD valign="top" align="left" colspan="2">
						<%
			If IsArray(TempArr1) Then
				If j <= Ubound(TempArr1) Then
					TempArr2 = Str_Split(TempArr1(j),GCOL)						
					If IsArray(TempArr2) Then
						For k = 0 To Ubound(TempArr2) Step 2
							Response.Write "<A HREF=vbscript:FncPgmMenu(" & Chr(34) & TempArr2(K + 1) & Chr(34) & ")"
							Response.Write " ONMOUSEOVER=" & Chr(34) & "vbscript:Window.event.srcElement.ClassName='MENUOVER'" & Chr(34) 
							Response.Write " ONMOUSEOUT=" & Chr(34) & "vbscript:Window.event.srcElement.ClassName='MENU'" & Chr(34) 
							if Len2(TempArr2(k)) >18 then 
								Response.Write " CLASS='MENU'>" & Left2(TempArr2(k),14) & "..</A><BR>" & vbCrLf
							else
								Response.Write " CLASS='MENU'>" & TempArr2(k) & "</A><BR>" & vbCrLf
							END IF
						Next
					End If				
					j = j + 1
				End If
			End If
    %>
					</TD>
<%       Next %>					

					<TD background="../../../CShared/image/uniSIMS/right_wall_73.gif" width="64"></TD>
				</TR>
				<TR height="42">
					<TD background="../../../CShared/image/uniSIMS/left_wall_71.gif" width="55"></TD>
					<TD width="38"></TD>
					<%
            If IsArray(TempArr) Then
				If i <= Ubound(TempArr) Then
					Response.Write "<TD background=../../../CShared/image/uniSIMS/titlebox_33.gif CLASS=MENU width=94 valign=top height=42>"
					Response.Write TempArr(i)				
					i = i + 1
				Else
					Response.Write "<td width=94 valign=top height=42>"
				End If
			End If
    %>
					</TD>
<%     for h=33 to 41 step 2%>					
					<td width="50"></td>
					<%
            If IsArray(TempArr) Then
				If i <= Ubound(TempArr) Then
                Response.Write "<TD background=../../../CShared/image/uniSIMS/titlebox_" & h & ".gif CLASS=MENU width=94 valign=top height=42>"
					Response.Write TempArr(i)				
					i = i + 1
			    Else
                   Response.Write "<td width=94 valign=top height=42>"
				End If
			End If
    %>
					</TD>
<%       Next %>						
					<td width="50"></td>
					<TD background="../../../CShared/image/uniSIMS/right_wall_73.gif" width="64"></TD>
				</TR>
				<TR height="150">
					<TD background="../../../CShared/image/uniSIMS/left_wall_71.gif" width="55"></TD>
					<TD width="38"></TD>
<%     for h=1 to 6 %>						
					<TD width="94" valign="top" align="left" COLSPAN="2">
						<%
			If IsArray(TempArr1) Then
				If j <= Ubound(TempArr1) Then
					TempArr2 = Str_Split(TempArr1(j),GCOL)						
					If IsArray(TempArr2) Then
						For k = 0 To Ubound(TempArr2) Step 2
							Response.Write "<A HREF=vbscript:FncPgmMenu(" & Chr(34) & TempArr2(K + 1) & Chr(34) & ")"
							Response.Write " ONMOUSEOVER=" & Chr(34) & "vbscript:Window.event.srcElement.ClassName='MENUOVER'" & Chr(34) 
							Response.Write " ONMOUSEOUT=" & Chr(34) & "vbscript:Window.event.srcElement.ClassName='MENU'" & Chr(34) 
							if Len2(TempArr2(k)) >18 then 
								Response.Write " CLASS='MENU'>" & Left2(TempArr2(k),14) & "..</A><BR>" & vbCrLf
							else
								Response.Write " CLASS='MENU'>" & TempArr2(k) & "</A><BR>" & vbCrLf
							END IF
						Next
					End If				
					j = j + 1
				End If
			End If
    %>
					</TD>
<%       Next %>	
					<TD background="../../../CShared/image/uniSIMS/right_wall_73.gif" width="64"></TD>
				</TR>
				<tr>
					<td width="55" height="13" valign="top" background="../../../CShared/image/uniSIMS/left_down_wall.gif"></td>
					<td background="../../../CShared/image/uniSIMS/wall_down_back_53.gif" colspan="13" height="12"></td>
					<td rowspan="2" valign="top" align="right" width="64" background="../../../CShared/image/uniSIMS/wall_down_right_52.jpg"></td>
				</tr>
				<tr>
					<td colspan="14">&nbsp;</td>
				</tr>
			</table>
		</DIV>
		<DIV ID="divPgmMenu" style="VISIBILITY: hidden; POSITION: absolute; zindex: 0">
			<TABLE cellSpacing="0" cellPadding="0" BORDER="0" bgcolor="#ffffff" width="1024">
				<TR height="400">
					<TD valign="top" align="middle" width="200">
						<%		
	If IsArray(LeftMenu) Then
		For i = 0 To Ubound(LeftMenu)
			TempArr = Str_Split(LeftMenu(i),GCOL)
			If IsArray(TempArr) Then
				If Mid(TempArr(5),1,1) = "M" Then
					If i <> 0 Then
						Response.Write "</DIV>"	 & vbCrLf
					End If
					Response.Write  "<DIV ID='" & TempArr(0) & LEFTNAME & "' CLASS='" & LEFTMENUBAR & "' LEVEL ='" & MENUAREA & "'>" & vbCrLf
				End If
				If TempArr(5) = "MM" Then
					Response.Write "<TABLE cellSpacing=0 cellPadding=0 BORDER=0 bgcolor=#ffffff><TR><TD background='../../../CShared/image/uniSIMS/leftmenu1.jpg' width=24 height=20></TD>" & vbCrLf
					Response.Write "<TD background='../../../CShared/image/uniSIMS/leftmenu2.jpg' width=130 height=20 align=center>" & vbCrLf
					Response.Write "<DIV CLASS='" & LEFTMAINCLASS & "'>"
					Response.Write TempArr(1)
					Response.Write "</DIV>" & vbCrLf
					Response.Write "</TD></TR></TABLE>" & vbCrLf
				ElseIf TempArr(5) = "MP" Or TempArr(5) = "PP" Then
					Response.Write "<A HREF='#' ID='" & TempArr(0) & LEFTID & "' CLASS='" & LEFTSUBCLASS & "' MTitle='" & TempArr(1) & "' LEVEL ='" & TempArr(3) & "'>"
					Response.Write "::" & TempArr(1)
					Response.Write "</A><BR>" & vbCrLf
				End If
			End If
		Next

		Response.Write "</DIV>" & vbCrLf
	End IF
%>
					</TD>
					<TD vAlign="top" align="left" width="800" colspan="3">
						<TABLE width="100%" cellSpacing="0" cellPadding="0" BORDER="0" bgcolor="#ffffff">
							<TR height="26" valign="top">
								<TD colspan="3">
									<TABLE width="100%" height="100%" cellSpacing="0" cellPadding="0">
										<TR valign="center">
											<TD width="13" background="../../../CShared/image/uniSIMS/pgmtitle1.jpg"></TD>
											<TD width="200" background="../../../CShared/image/uniSIMS/pgmtitle.jpg" align="middle"><INPUT type="text" NAME="txtTitle" readonly style='BORDER-RIGHT: medium none; BORDER-TOP: medium none; FONT-WEIGHT: bolder; FONT-SIZE: 12pt; BORDER-LEFT: medium none; COLOR: white; BORDER-BOTTOM: medium none; BACKGROUND-COLOR: transparent; TEXT-ALIGN: center'></TD>
											<TD width="573" background="../../../CShared/image/uniSIMS/pgmtitle2.jpg"></TD>
											<TD width="14" background="../../../CShared/image/uniSIMS/pgmtitle3.jpg"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR height="21" valign="top">
								<TD width="13"></TD>
								<TD ALIGN="right">
									<DIV id="nextprev" style='VISIBILITY:hidden;HEIGHT:21px'>
										<span class='TXTCHK'>사번:</span><INPUT type="text" class="inputbox" NAME="txtEmp_no2" tag="1" MAXLENGTH="13" SiZE="13" style='FONT-SIZE: 9pt'><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" border="0" TYPE="BUTTON" onclick="VBScript:Call OpenEmp(txtemp_no.value)">
										<span class='TXTCHK'>성명:</span><INPUT type="text" class="inputbox" NAME="txtName2" MAXLENGTH="15" SiZE="15" style='FONT-SIZE: 9pt'>
										<A ONCLICK="VBSCRIPT:CALL FncPrev()" onMouseOver="javascript: this.style.cursor='hand'">
											<IMG src="../../../CShared/image/uniSIMS/prev.jpg" alt='이전' border="0"></A> <A ONCLICK="VBSCRIPT:CALL FncNext()" onMouseOver="javascript: this.style.cursor='hand'">
											<IMG src="../../../CShared/image/uniSIMS/next.jpg" alt='다음' border="0"></A>
									</DIV>
								</TD>
								<TD width="14"></TD>
							</TR>
							<TR height="13">
								<TD background="../../../CShared/image/uniSIMS/body1left.jpg" width="13"></TD>
								<TD background="../../../CShared/image/uniSIMS/body1.jpg"></TD>
								<TD background="../../../CShared/image/uniSIMS/body1right.jpg" width="14"></TD>
							</TR>
							<TR height="7">
								<TD background="../../../CShared/image/uniSIMS/bodyleft.jpg" width="13"></TD>
								<TD background="../../../CShared/image/uniSIMS/body.jpg"></TD>
								<TD background="../../../CShared/image/uniSIMS/bodyright.jpg" width="14"></TD>
							</TR>
							<TR height="320">
								<TD background="../../../CShared/image/uniSIMS/bodyleft.jpg" width="13"></TD>
								<TD width="773">
									<IFRAME id="formmenu" NAME="formmenu" src="" WIDTH="100%" HEIGHT="100%" FRAMEBORDER="0" framespacing="0" SCROLLING="auto"></IFRAME>
								</TD>
								<TD background="../../../CShared/image/uniSIMS/bodyright.jpg" width="14"></TD>
							</TR>
							<TR height="7">
								<TD background="../../../CShared/image/uniSIMS/bodyleft.jpg" width="13"></TD>
								<TD background="../../../CShared/image/uniSIMS/body.jpg"></TD>
								<TD background="../../../CShared/image/uniSIMS/bodyright.jpg" width="14"></TD>
							</TR>
							<TR height="5">
								<TD background="../../../CShared/image/uniSIMS/body2left.jpg" width="13"></TD>
								<TD background="../../../CShared/image/uniSIMS/body2.jpg"></TD>
								<TD background="../../../CShared/image/uniSIMS/body2right.jpg" width="14"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR height="40" valign="bottom">
					<TD align="left" width="200"></TD>
					<TD valign="center" align="left" width="300">
						<DIV ID="MousePT" NAME="MousePT">
							<iframe name="MouseWindow" FRAMEBORDER="0" SCROLLING="no" noresize framespacing="0" width="220" height="41" src="../../inc/cursor.htm"></iframe>
						</DIV>
					</TD>
					<TD vAlign="bottom" width="500" align="right">
						<INPUT type="image" SRC="../../../CShared/image/uniSIMS/ret1.jpg" WIDTH="28" HEIGHT="27" border="0" OnClick="vbscript: FncQuery()" name="SUBMIT" alt='조회' onMouseOver="javascript:this.src='../../../CShared/image/uniSIMS/ret2.jpg';" onMouseOut="javascript:this.src='../../../CShared/image/uniSIMS/ret1.jpg';">
						<INPUT type="image" SRC="../../../CShared/image/uniSIMS/add1.jpg" WIDTH="28" HEIGHT="27" border="0" OnClick="vbscript: FncAdd()" name="add" alt='추가' onMouseOver="javascript:this.src='../../../CShared/image/uniSIMS/add2.jpg';" onMouseOut="javascript:this.src='../../../CShared/image/uniSIMS/add1.jpg';">
						<INPUT type="image" SRC="../../../CShared/image/uniSIMS/del1.jpg" WIDTH="28" HEIGHT="27" border="0" OnClick="vbscript: FncDel()" name="del" alt='삭제' onMouseOver="javascript:this.src='../../../CShared/image/uniSIMS/del2.jpg';" onMouseOut="javascript:this.src='../../../CShared/image/uniSIMS/del1.jpg';">
						<INPUT type="image" SRC="../../../CShared/image/uniSIMS/save1.jpg" WIDTH="28" HEIGHT="27" border="0" OnClick="vbscript: FncSave()" name="save" alt='저장' onMouseOver="javascript:this.src='../../../CShared/image/uniSIMS/save2.jpg';" onMouseOut="javascript:this.src='../../../CShared/image/uniSIMS/save1.jpg';">
						<INPUT type="image" SRC="../../../CShared/image/uniSIMS/print1.jpg" WIDTH="28" HEIGHT="27" border="0" OnClick="vbscript: FncPrint()" name="prt" alt='출력' onMouseOver="javascript:this.src='../../../CShared/image/uniSIMS/print2.jpg';" onMouseOut="javascript:this.src='../../../CShared/image/uniSIMS/print1.jpg';">
					</TD>
					<TD vAlign="center" width="14" align="right">
					</TD>
				</TR>
			</TABLE>
		</DIV>
		<INPUT type=hidden NAME="txtemp_no" value="<%=gEmpNo%>"> <INPUT type=hidden NAME="txtname" value="<%=name%>">
		<INPUT type=hidden NAME="txtinternal_cd" value="<%=internal_cd%>"> <INPUT type=hidden NAME="txtnat_cd" value="<%=nat_cd%>">
		<INPUT type=hidden NAME="txtDEPT_AUTH" value="<%=gDeptAuth%>"> <INPUT type=hidden NAME="txtPRO_AUTH" value="<%=gProAuth%>">
		<INPUT type=hidden NAME="txtLang" value="<%=gLang%>"> <INPUT type=hidden NAME="txtYearEnd" value="<%=gLastYearEnd%>">
		<INPUT type=hidden NAME="txtdept_nm" value="<%=dept_nm%>"> 
		<DIV style="DISPLAY:none">
			<script language =javascript src='./js/emenu_SIMS_SIMS.js'></script>
		</DIV>
		<IFRAME id="logoff" name="logoff" Style="DISPLAY:none"></IFRAME>
		<script Language="vbscript">
TempMain	= "<%=Join(TopMenu,GROW) & GROW%>"
TempSub		= "<%=Join(TopSub,GROW) & GROW%>"
TempLeft	= "<%=Join(LeftMenu,GROW) & GROW%>"
		</script>
	</body>
</HTML>

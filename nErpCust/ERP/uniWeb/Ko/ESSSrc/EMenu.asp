<%@ Language=VBScript %>
<%

   Response.Buffer = True
   Response.Expires = -1


   If Request.Cookies("unierp")("gjdoiwp") = "" Then
      Response.Redirect "../Scam.asp"
   End If
   
%>
<HTML>
<!-- #Include file="../ESSinc/Adovbs.inc"  -->
<!-- #Include file="../ESSinc/incServerAdoDb.asp" -->
<!-- #Include file="../ESSinc/incServer.asp" -->
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->

<% 
Dim	Name
Dim	dept_nm
Dim	entr_dt
Dim internal_cd
Dim nat_cdf

    Call SubOpenDB(lgObjConn)                                            '☜: Make a DB Connection
	
	if gEmpNo = "unierp" then
		Name = "unierp"
	else
		lgStrSQL = " SELECT Emp_no, NAME, dept_nm, pay_grd2, entr_dt, internal_cd, nat_cd "
		lgStrSQL = lgStrSQL & " FROM haa010t where emp_no= " & FilterVar(gEmpNo, "''", "S") & ""

		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = true Then
			Name	   = lgObjRs("NAME")
			dept_nm    = lgObjRs("dept_nm")
			entr_dt    = lgObjRs("entr_dt")
			internal_cd= lgObjRs("internal_cd")
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
<!-- #Include file="../ESSinc/incSvrVarSims.inc"  -->
<!-- #Include file="../ESSinc/incSvrFuncSims.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" href="../ESSinc/common.css">

<SCRIPT LANGUAGE="VBScript" SRC="../ESSinc/ccm.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../ESSinc/variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../ESSinc/incCookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../ESSinc/operation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../ESSinc/incCommFunc.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../ESSinc/incEvent.vbs"></SCRIPT>			
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
Const LEFTMENUWIDTH			= 195

Const VIEWMENUCNT			= 9
Const SKIPMENUCNT			= 1
Const MENUXSPACE			= 0
Const MENUYSPACE			= 0
Const MCLASS				= "TOPMAIN"
Const SCLASS				= "TOPSUB"
Const LCLASS				= "LEFTMENU"

Const MENUTOP				= "MENUTOP"
Const MENUEND				= "MENUEND"

Const CLICK_COLOR			= "#E07000"
Const OVER_COLOR			= "#017FA8"
Const OVER_COLOR_SCLASS		= "#F3FF62"
Const SUB_OUT_COLOR			= "WHITE"
Const LEFT_OUT_COLOR		= "#717173"
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
CurrURL = ""

'========================= Window Event =======================
'==============================================================
'Function: Window_onLoad()
'==============================================================

Function Window_onLoad()

	Dim IDx
	
	On Error Resume Next
	Err.Clear 

    if  Trim(txtemp_no.value) = "" Then
        document.location = "../default.asp"
    End If

    window.document.body.scroll = "no"

	Call Menu_Init(TempMain,MCLASS)		
	Call Menu_Init(TempSub,SCLASS)	
	Call Menu_Init(TempLeft,LCLASS)
	Call FncHomeMenu()
    Call Menu_Display(MCLASS,TopMain)

    document.All("nextprev").style.VISIBILITY = "hidden"
'------------첫화면에 공지사항 setting
	document.all("divHomeMenu").style.VISIBILITY = "hidden"
	document.all("DivPgmMenu").style.VISIBILITY = "visible"
 
	Call formmenu_onLoad(inPagevalue) 
	document.All("formmenu").src = "ESSBoard_list.asp"
	txtTitle.value="공지사항"
    document.title = gLogoName & " - " & LeftMenu(InIDx).MTitle & " [ " & "<%=NAME%>" & " ]"	
    
    Call Menu_Display(LCLASS,"E18_LEFTID")

	IDx = Menu_Search(LeftMenu,"E1807MA1_LEFTID","MENUIDX")
	If IDx <> -1 Then
		Call Click_OpenMenu(IDx,LeftMenu,LCLASS)
		Call Click_Menu(IDx,LeftMenu,LCLASS)		
	End If

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
	If UCase(CuEvObj.tagName) = "TD" Then Exit Function
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
'    Call SetToolBar("0000")
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
				CurrObj.style.color	= OVER_COLOR_SCLASS
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
				CurrObj.style.color	= OVER_COLOR
			End If
	End Select
	If InList(InIDx).MOVERFLAG Then
		CurrObj.style.cursor= OVER_CURSOR		
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
				Set GObj = document.all(InMenu(i).ID)'.parentElement
				InMenu(i).GROUP			= GObj.GROUP
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

'    Call SetToolBar("0000")
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
    
    
    window.document.location = "./e1logoffmb1.asp?txtMode=UID_M0003"
   
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
	    	"dialogWidth=406px; dialogHeight=233px; center: Yes; help: No; resizable: No; status: No;")
    else		
	    arrRet = window.showModalDialog("EchangePWFirst.asp", Array(arrParam), _
	    	"dialogWidth=406px; dialogHeight=238px; center: Yes; help: No; resizable: No; status: No;")
    end if
End Function

Function FncHelp()

	dim from_GetProgId

	from_GetProgId= formmenu.GetProgId()
	
    If from_GetProgId = "" Then 
        window.open "../ESSHelp/ess_help.asp","ESShelp","resizable=yes,scrollbars=yes,status=no,toolbar=no,menubar=no,height=600,width=845,center=yes,top=45,left=100"
    Else
        window.open "../ESSHelp/ess_help.asp?path=" & from_GetProgId & ".doc", "ESShelp","resizable=yes,scrollbars=yes,status=no,toolbar=no,menubar=no,height=600,width=845,center=yes,top=45,left=100"
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
    		"dialogWidth=546px; dialogHeight=387px; center: Yes; help: No; resizable: No; status: No;")
    Else
        iWhereFlg = False
        
        arrParam(2) = Trim(txtinternal_cd.Value)' lgUsrIntCd
	    arrRet = window.showModalDialog("E1EmpPopa1.asp", Array(arrParam), _
	    	"dialogWidth=546px; dialogHeight=387px; center: Yes; help: No; resizable: No; status: No;")
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
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_showHideLayers() { //v6.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) if ((obj=MM_findObj(args[i]))!=null) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
}

//-->

</script>
<HEAD>
<!-- #Include file="../ESSinc/uniSimsClassID.inc" -->
</HEAD>
<body leftmargin="0" topmargin="0" marginwidth="0">
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
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="943" valign="top">
	  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>       <!-- 1. 로고 및 아이콘 -- start -->
          <td height="43">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td><a href="#"><img src="../ESSimage/logo.gif" width="132" height="43" border="0"></a></td> 
                <td align="right" valign="bottom">
<%
    Response.Write "<a href='vbscript:FncHomeMenu()' onMouseOut=javascript:MM_swapImgRestore() "
    Response.Write " onMouseOver=javascript:MM_swapImage('home','','../../CShared/ESSimage/ut_menu_01.gif',1)>"
    Response.Write "<img name=home border=0 src=../../CShared/ESSimage/ut_menu_01.gif alt='HOME'></a>"
    Response.Write "<a href='vbscript:FncLogoff(2)' onMouseOut=javascript:MM_swapImgRestore()"
    Response.Write " onMouseOver=javascript:MM_swapImage('logout','','../../CShared/ESSimage/ut_menu_02.gif',1)>"
    Response.Write "<img name=logout border=0 src=../../CShared/ESSimage/ut_menu_02.gif alt='로그오프'></a>"
    Response.Write "<a href='vbscript:FncPassword(1)' onMouseOut=javascript:MM_swapImgRestore()"
    Response.Write " onMouseOver=javascript:MM_swapImage('password','','../../CShared/ESSimage/ut_menu_03.gif',1)>"
    Response.Write "<img name=password border=0 src=../../CShared/ESSimage/ut_menu_03.gif alt='패스워드변경'></a>"
    Response.Write "<a href='vbscript:FncHelp()' onMouseOut=javascript:MM_swapImgRestore()"
    Response.Write " onMouseOver=javascript:MM_swapImage('admin','','../../CShared/ESSimage/ut_menu_04.gif',1)>"
    Response.Write "<img name=admin border=0 src=../../CShared/ESSimage/ut_menu_04.gif alt='HELP'></a></td>" & vbCrLf
	Response.Write "</tr>" & vbCrLf
	Response.Write "</table></td>" & vbCrLf
	Response.Write "</tr>" & vbCrLf   ' 1. 로고 및 아이콘 -- end
        
    '------------------ Top Menu S ------------------------->        
	Response.Write "<TR>" & vbCrLf    ' 2. 상위 메뉴 -- start
	Response.Write "  <td height=34>" & vbCrLf
	Response.Write "    <DIV ID='" & TOPMENUBAR & "' CLASS='" & TOPMENUBAR & "' LEVEL ='" & MENUAREA & "' NOWRAP>" & vbCrLf
	Response.Write "       <table width=100% border=0 cellspacing=0 cellpadding=0>" & vbCrLf
	Response.Write "         <TR class=TOPMENUBAR>" & vbCrLf    

	If IsArray(TopMenu) Then
		For i = 0 To Ubound(TopMenu)
			TempArr = Str_Split(TopMenu(i),GCOL)

			If IsArray(TempArr) Then
				If TempArr(5) = "AS" Then
				else
					If TempArr(5) = "AE" Then
					Else
						Response.Write "<td width=100 background=../../CShared/ESSimage/main_menu_bg.gif>" 				
						Response.Write "<A HREF='" & TempArr(2) & "' ID='" & TempArr(0) & "' CLASS='" & MAINCLASS & "' MTitle='" & TempArr(1) & "' LEVEL ='" & TempArr(3) & "' >"
						Response.Write TempArr(1) & vbCrLf	
						Response.Write "</A></td>" & vbCrLf		
					end if
				End If
			End If
		Next
		
		If Ubound(TopMenu) = 9 Then
			Response.Write "<td width=100 background=../../CShared/ESSimage/main_menu_bg.gif></td>" & vbCrLf
		ElseIf Ubound(TopMenu) < 9 Then
			For i = Ubound(TopMenu) To 9-1
				Response.Write "<td width=100 background=../../CShared/ESSimage/main_menu_bg.gif></td>" & vbCrLf
			Next
			Response.Write "<td width=100 background=../../CShared/ESSimage/main_menu_bg_02.gif></td>" & vbCrLf
		End If
	End If
	Response.Write "</TR>" & vbCrLf
	Response.Write "</TABLE>" & vbCrLf
	Response.Write "</DIV>" & vbCrLf
	Response.Write "</TD>" & vbCrLf
    Response.Write "</TR>" & vbCrLf     ' 2. 상위 메뉴 -- end

'=======================Top SUB  ===========================	
	Response.Write "<TR>" & vbCrLf      ' 3. 상위 메뉴 서브 -- start
	Response.Write "<td height=87 background='../ESSimage/main_img_01.gif' LEVEL='" & MENUAREA & "'>" & vbCrLf

	If IsArray(TopSub) Then
		PMenu = ""
		NMenu = ""
		For i = 0 To Ubound(TopSub)		
			TempArr = Str_Split(TopSub(i),GCOL)
			If IsArray(TempArr) Then	
				PMenu = TempArr(4)			
				If PMenu <> NMenu Then
					If NMenu <> "" Then
						Response.Write "&nbsp;"
						Response.Write "    </td>" & vbCrLf
						Response.Write "  </tr>" & vbCrLf
						Response.Write "</table>" & vbCrLf
						Response.Write "</DIV>" & vbCrLf
						NMenu = ""					
					End If
					Response.Write "<DIV ID='" & PMenu & SUBNAME & "' CLASS='" & TOPSUBBAR & "' LEVEL ='" & MENUAREA & "'  NOWRAP>" & vbCrLf
					Response.Write "<table  border=0 cellpadding=0 cellspacing=0>" & vbCrLf
					Response.Write "  <tr>" & vbCrLf
					Response.Write "	<td width=10></td>" & vbCrLf
					Response.Write "	<td height=24 bgcolor=#78AC00 class='submenu' LEVEL ='" & MENUAREA & "'>" & vbCrLf
					Response.Write "&nbsp;"
					NMenu = PMenu
				Else			
					Response.Write "ㅣ" & vbCrLf
				End If
				Response.Write "<SPAN HREF=" & TempArr(2) & " ID='" & TempArr(0) & "' GROUP='" & PMenu & SUBNAME & "' MTitle='" & TempArr(1) & "' LEVEL ='" & TempArr(3) & "'>" & vbCrLf
				Response.Write TempArr(1)
				Response.Write "</SPAN>" & vbCrLf
			End If
		Next
		Response.Write "&nbsp;"
		Response.Write "    </td>" & vbCrLf
		Response.Write "  </tr>" & vbCrLf
		Response.Write "</table>" & vbCrLf
		Response.Write "</DIV>" & vbCrLf
	End If
	Response.Write "</TD>" & vbCrLf
	Response.Write "</TR>" & vbCrLf    
	' 3. 상위 메뉴 서브 -- end
    
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
		<TR>     <!-- 4. 해바라기 그림 및 홈 -- start -->
           <td width=100% height=* valign=top>
		   <DIV ID="divHomeMenu" style="VISIBILITY: visible; POSITION: absolute">
           <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
                <td width="195" valign="top">
					<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td height="72"><img src="../../CShared/ESSimage/main_img_02.gif"></td>
                    </tr>
                    <tr> 
                      <!-------------------- Left Menu S ------------------------->
                      <td valign="top" background="../../CShared/ESSimage/left_menu_bg_02.gif">&nbsp;</td>
                      <!---------------------- Left Menu E ------------------------->
                    </tr>
                    </table>
                    <p class="submenu"></p>
                </td>
                <td width=15></td>
                <!--------------------------------- Content S --------------------------------------->
                <td align=center valign=top>
                    <table border=0 cellspacing=0 cellpadding=0>
                    <tr> 
					  <%for h=1 to 4%>
                      <td width=165 valign=top>
                         <table width=100% border=0 cellspacing=0 cellpadding=0>
							<%   ' 홈에서 메인 메뉴 
							If IsArray(TempArr) Then
								If i <= Ubound(TempArr) Then
									Response.Write "<TD height=29 background=../../CShared/ESSimage/home_0" & h & ".gif CLASS=MENU width=94 valign=top height=42>"
									Response.Write "<font color=00375A>" & TempArr(i) & "</font>"		
									i = i + 1
									Response.Write "</TD>"		
									Response.Write "</TR>"		
									Response.Write "<TR>"
									Response.Write "   <td height=5 background=../../CShared/ESSimage/home_bg_01.gif></td>"		
									Response.Write "</TR>"
								End If
							End If
							%>						
						<%        ' 홈에서 서브 메뉴 
						If IsArray(TempArr1) Then
							If j <= Ubound(TempArr1) Then
								TempArr2 = Str_Split(TempArr1(j),GCOL)						
								If IsArray(TempArr2) Then
									For k = 0 To Ubound(TempArr2) Step 2
										Response.Write "<TR>" & vbCrLf									
										Response.Write "<TD background=../../CShared/ESSimage/home_bg_01.gif class=home02><img src=../../CShared/ESSimage/home_ic.gif width=8 height=10>" & vbCrLf 
										Response.Write "<A HREF=vbscript:FncPgmMenu(" & Chr(34) & TempArr2(K + 1) & Chr(34) & ")"
										Response.Write " ONMOUSEOVER=" & Chr(34) & "vbscript:Window.event.srcElement.ClassName='MENUOVER'" & Chr(34) 
										Response.Write " ONMOUSEOUT=" & Chr(34) & "vbscript:Window.event.srcElement.ClassName='MENU'" & Chr(34) 
										if Len2(TempArr2(k)) >26 then 
											Response.Write " CLASS='MENU'>" & Left2(TempArr2(k),22) & "..</A><BR>" & vbCrLf
										else
											Response.Write " CLASS='MENU'>" & TempArr2(k) & "</A><BR>" & vbCrLf
										END IF
										Response.Write "</TD>" & vbCrLf									
										Response.Write "</TR>" & vbCrLf									
									Next
								End If				
								j = j + 1
								Response.Write "<TR>" & vbCrLf									
								Response.Write "<TD><img src=../../CShared/ESSimage/home_bg_02.gif width=165 height=6></TD>" & vbCrLf									
								Response.Write "</TR>" & vbCrLf									
							End If
						End If
						%>
					  </table>
					  </td>
					  <td width=15></td>
					<%Next %>					
					</tr>
                    <tr> 
                      <td height=15></td>
                    </tr>
                    <tr>  <!-- 홈 다음 줄 -->
					  <%for h=5 to 8%>
                      <td width=165 valign=top><table width=100% border=0 cellspacing=0 cellpadding=0>
	                    <tr>
							<%   ' 홈에서 메인 메뉴 
							If IsArray(TempArr) Then
								If i <= Ubound(TempArr) Then
									Response.Write "<TD height=29 background=../../CShared/ESSimage/home_0" & h & ".gif CLASS=MENU width=94 valign=top height=42>"
									Response.Write "<font color=00375A>" & TempArr(i) & "</font>"		
									i = i + 1
									Response.Write "</TD>"		
									Response.Write "</TR>"		
									Response.Write "<TR>"		
									Response.Write "   <td height=5 background=../../CShared/ESSimage/home_bg_01.gif></td>"		
									Response.Write "</TR>"		
								End If
							End If
							%>						
						<%        ' 홈에서 서브 메뉴 
						If IsArray(TempArr1) Then
							If j <= Ubound(TempArr1) Then
								TempArr2 = Str_Split(TempArr1(j),GCOL)						
								If IsArray(TempArr2) Then
									For k = 0 To Ubound(TempArr2) Step 2
										Response.Write "<TR>" & vbCrLf									
										Response.Write "<TD width=165 background=../../CShared/ESSimage/home_bg_01.gif class=home02><img src=../../CShared/ESSimage/home_ic.gif width=8 height=10>" & vbCrLf 
										Response.Write "<A HREF=vbscript:FncPgmMenu(" & Chr(34) & TempArr2(K + 1) & Chr(34) & ")"
										Response.Write " ONMOUSEOVER=" & Chr(34) & "vbscript:Window.event.srcElement.ClassName='MENUOVER'" & Chr(34) 
										Response.Write " ONMOUSEOUT=" & Chr(34) & "vbscript:Window.event.srcElement.ClassName='MENU'" & Chr(34) 
										if Len2(TempArr2(k)) >18 then 
											Response.Write " CLASS='MENU'>" & Left2(TempArr2(k),14) & "..</A><BR>" & vbCrLf
										else
											Response.Write " CLASS='MENU'>" & TempArr2(k) & "</A><BR>" & vbCrLf
										END IF
										Response.Write "</TD>" & vbCrLf									
										Response.Write "</TR>" & vbCrLf									
									Next
								End If				
								j = j + 1
								Response.Write "<TR>" & vbCrLf									
								Response.Write "<TD><img src=../../CShared/ESSimage/home_bg_02.gif width=165 height=6></TD>" & vbCrLf									
								Response.Write "</TR>" & vbCrLf									
							End If
						End If
						%>
					  </table>
					  </td>
					  <td width=10></td>
					<%Next %>					
					</tr>
					<tr>
					  <td height=65></td>
					</tr>
				</td>
				</table>
			    </tr>
			</table>
		    </DIV>

		    <DIV ID="divPgmMenu" style="VISIBILITY: hidden; POSITION: absolute; zindex: 0">
			<table width="100%" height="100%" border=0 cellpadding="0" cellspacing="0">
				<TR>
					<TD width="195" valign="top">
					  <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
						<tr>
						  <td height="72"><img src="../../CShared/ESSimage/main_img_02.gif"></td>
						</tr>
						<tr>
						  <!-------------------- Left Menu ------------------------->
						  <td align=left valign=top height=*>
							<%		
							If IsArray(LeftMenu) Then
								For i = 0 To Ubound(LeftMenu)
									TempArr = Str_Split(LeftMenu(i),GCOL)
									If IsArray(TempArr) Then
										If Mid(TempArr(5),1,1) = "M" Then
											If i <> 0 Then
												Response.Write "    <tr>" & vbCrLf
												Response.Write "      <td height=*></td>" & vbCrLf
												Response.Write "    </tr>" & vbCrLf
												Response.Write "  </table>" & vbCrLf
												Response.Write "</td>" & vbCrLf
												Response.Write "</TR>" & vbCrLf
												Response.Write "</TABLE>" & vbCrLf
												Response.Write "</DIV>"	 & vbCrLf
											End If
											Response.Write  "<DIV ID='" & TempArr(0) & LEFTNAME & "' CLASS='" & LEFTMENUBAR & "' LEVEL ='" & MENUAREA & "'>" & vbCrLf
											Response.Write " <TABLE width=100% border=0 cellspacing=0 cellpadding=0>"	 & vbCrLf
										End If
										If TempArr(5) = "MM" Then
											Response.Write "<TR>" & vbCrLf
											Response.Write "  <TD height=41 background='../../CShared/ESSimage/left_menu_title.gif' class='ltmntitle'>" & vbCrLf
											Response.Write "    <DIV CLASS='" & LEFTMAINCLASS & "'>"
											Response.Write TempArr(1)
											Response.Write "    </DIV>" & vbCrLf
											Response.Write "  </TD>" & vbCrLf
											Response.Write "</TR>" & vbCrLf
											Response.Write "<TR>" & vbCrLf
											Response.Write "  <td valign=top background=../../CShared/ESSimage/left_menu_bg_02.gif>" & vbCrLf
											Response.Write "    <table width=100% height=340 border=0 cellspacing=0 cellpadding=0>" & vbCrLf
										ElseIf TempArr(5) = "MP" Or TempArr(5) = "PP" Then
											Response.Write "      <TR>" & vbCrLf
											Response.Write "       <td height=31 background='../../CShared/ESSimage/left_menu_bg_01.gif' class='ltmenu02'>" & vbCrLf
											Response.Write "         <A HREF='#' ID='" & TempArr(0) & LEFTID & "' CLASS='" & LEFTSUBCLASS & "' MTitle='" & TempArr(1) & "' LEVEL ='" & TempArr(3) & "'>"
											Response.Write TempArr(1)
											Response.Write "         </A></td>" & vbCrLf
											Response.Write "      </TR>" & vbCrLf
										End If
									End If
								Next
								Response.Write "    <tr>" & vbCrLf
								Response.Write "      <td height=*></td>" & vbCrLf
								Response.Write "    </tr>" & vbCrLf
								Response.Write "  </table>" & vbCrLf
								Response.Write "</td>" & vbCrLf
								Response.Write "</TR>" & vbCrLf
								Response.Write "</TABLE>" & vbCrLf
								Response.Write "</DIV>" & vbCrLf
							End IF
							%>
					      </td>
					    </tr>
					  </table>
                    <p class="submenu"></p>
					</TD>
                    <td width=10></td>
                    <TD valign="top">
                    	<TABLE width="100%" cellSpacing="0" cellPadding="0" BORDER="0">
							<TR>
								<TD height=38>
									<TABLE width=100% height=38 border=0 cellpadding=3 cellspacing=1 bgcolor=DDDDDD>
									<TR>
										<TD bgcolor=F5F5F5>
										<!------------------  Title S ----------------------->
										<table width=100% border=0 cellspacing=0 cellpadding=1>
										 <tr> 
										   <td width=30 height=30 align=center bgcolor=#FFFFFF><img src=../../CShared/ESSimage/title_icon.gif></td>
										   <td bgcolor=#FFFFFF>
										      <INPUT class=contitle NAME="txtTitle" readonly tabindex=-1></td>
										   <td align=right bgcolor=#FFFFFF>
											 <!-------- 사번, 성명 S ------->
											 <DIV id="nextprev" style='VISIBILITY:hidden;'>
											   <table border=0 cellspacing=0 cellpadding=1>
											    <tr> 
													<td><img src=../../CShared/ESSimage/icon_03.gif width=10 height=12></td>
													<td class=ftgray>사번</td>
													<td width=2></td>
													<td><input type=text class=inputbox NAME="txtEmp_no2" tag="1" MAXLENGTH="13" SiZE="13" style=width:100px>&nbsp;<IMG SRC="../ESSimage/button_11.gif" NAME="btnCalType" border="0" TYPE="BUTTON" onMouseOver="javascript: this.style.cursor='hand'" onclick="VBScript:Call OpenEmp(txtemp_no.value)"></td>
													<td width=5></td>
													<td><img src=../../CShared/ESSimage/icon_03.gif width=10 height=12></td>
													<td class=ftgray>성명</td>
													<td width=2></td>
													<td><input type=text class=inputbox NAME="txtName2" MAXLENGTH="15" SiZE="15" style=width:100px></td>
													<td width=55 align=center bgcolor=#FFFFFF class=contitle>
													   <A ONCLICK="VBSCRIPT:CALL FncPrev()" onMouseOver="javascript: this.style.cursor='hand'">
													       <img src=../../CShared/ESSimage/icon_01.gif alt='이전' border=0></a> 
													   <A ONCLICK="VBSCRIPT:CALL FncNext()" onMouseOver="javascript: this.style.cursor='hand'">
														   <img src=../../CShared/ESSimage/icon_02.gif alt='다음' border=0></a> 
													</td>
											    </tr>
											   </table>
											 </DIV>
											 <!--------  사번,성명 E-------->
                                           </td>
                                         </tr>
									    </table>
										<!--------------------- Title E ----------------------->
										</TD>
								    </TR>
								    </TABLE>
								</TD>
							</TR>
							<TR> 
							  <td height=382 valign=top><IFRAME id="formmenu" NAME="formmenu" src="" WIDTH="100%" HEIGHT="100%" FRAMEBORDER="0" framespacing="0" SCROLLING="auto"></IFRAME></td>
							</TR>
						</TABLE>
					</TD>
				</TR>
                <tr> 
                  <td height=5></td>
                </tr>
				<TR height=10 valign="top" >
					<TD width="195" height=10></TD>
					<TD width="15"></TD>
					<TD valign="top" align="right">
						<INPUT type="image" style="display: 'none'" SRC="../ESSimage/button_01.gif" border="0" OnClick="vbscript: FncQuery()" name="SUBMIT" alt='조회' onMouseOver="javascript:this.src='../ESSimage/button_r_01.gif';" onMouseOut="javascript:this.src='../ESSimage/button_01.gif';">
						<INPUT type="image" style="display: 'none'" SRC="../ESSimage/button_05.gif" border="0" OnClick="vbscript: FncAdd()" name="add" alt='추가' onMouseOver="javascript:this.src='../ESSimage/button_r_05.gif';" onMouseOut="javascript:this.src='../ESSimage/button_05.gif';">
						<INPUT type="image" style="display: 'none'" SRC="../ESSimage/button_10.gif" border="0" OnClick="vbscript: FncDel()" name="del" alt='삭제' onMouseOver="javascript:this.src='../ESSimage/button_r_10.gif';" onMouseOut="javascript:this.src='../ESSimage/button_10.gif';">
						<INPUT type="image" style="display: 'none'" SRC="../ESSimage/button_02.gif" border="0" OnClick="vbscript: FncSave()" name="save" alt='저장' onMouseOver="javascript:this.src='../ESSimage/button_r_02.gif';" onMouseOut="javascript:this.src='../ESSimage/button_02.gif';">
						<INPUT type="image" style="display: 'none'" SRC="../ESSimage/button_04.gif" border="0" OnClick="vbscript: FncPrint()" name="prt" alt='출력' onMouseOver="javascript:this.src='../ESSimage/button_r_04.gif';" onMouseOut="javascript:this.src='../ESSimage/button_04.gif';">
					</TD>
				</TR>
			</TABLE>
		    </DIV>   <!-- 4. 해바라기 그림 및 홈 -- end -->
		    </td>
		</TR>        
		
	  </table>
	  </td>
	</tr>
	</table>
		
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

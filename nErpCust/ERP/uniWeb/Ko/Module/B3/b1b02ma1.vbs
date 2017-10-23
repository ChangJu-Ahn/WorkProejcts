	
Const BIZ_PGM_ID  = "b1b02mb1.asp"
Const BIZ_PGM_ID1 = "b1b02mb2.asp"
Const IMG_LOAD_PATH = "../../ComAsp/imgTemp.asp?src="

Const DIR_INIT_FILE = "../../../CShared/image/unierp20logo.gif"

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
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
	On Error Resume Next

	Const CookieSplit = 4877						
	Dim strTemp

	If flgs = 1 Then
		WriteCookie CookieSplit , frm1.txtItemCd.Value
	ElseIf flgs = 0 Then
		frm1.txtItemCd.Value = ReadCookie("txtItemCd")
		frm1.txtItemNm.value = ReadCookie("txtItemNm")
		
		WriteCookie "txtItemCd", ""
		WriteCookie "txtItemNm", ""
		
		If frm1.txtItemCd.value <> "" Then
			Call MainQuery()
		End If
		
	End If
	
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   If pOpt = "Q" Then
      lgKeyStream = Frm1.txtItemCd.Value & parent.gColSep
   Else
      lgKeyStream = Frm1.txtItemCd.Value & parent.gColSep
   End If   
End Sub
	
'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
    Call InitVariables                                                           '⊙: Initializes local global variables

    Call MakeKeyStream("Q")
    
    If GetItemCd = False Then   
		Exit Function           
    End If 
    'Call DbQuery                                                                 '☜: Query db data
       
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
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
    document.all.ImgItemImage.src = DIR_INIT_FILE
	
    Call SetToolbar("11101000000001")
    Call InitVariables                                                        '⊙: Initializes local global variables
    
    frm1.txtItemCd.focus
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
        Call DisplayMsgBox("900002", "X", "X", "X")                                
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")                        '☜: Do you want to delete? 
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    Call MakeKeyStream("D")

	If DbDelete = False Then   
		Exit Function           
    End If 
    
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '☜: Processing is OK

End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD, iStrFileType 
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = False Then 
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    If Not chkField(Document, "1") Then                             '⊙: Check contents area
       Exit Function
    End If
    
    If Not chkField(Document, "2") Then                                          '☜: Check contents area
       Exit Function
    End If

    If Not ggoSaveFile.FileExists(frm1.txtPath.value) = 0 Then
		Call DisplayMsgBox("115191", "X", "X", "X")
		Exit Function
    End If

    iStrFileType = Right(Trim(UCase(frm1.txtPath.value)), 3)

	If Not (iStrFileType = "BMP" Or iStrFileType = "GIF" Or iStrFileType = "JPG") Then
		Call DisplayMsgBox("122904", "X", "X", "X")
		Exit Function
	End If

    Call MakeKeyStream("S")
    
    If DbSave = False Then   
		Exit Function           
    End If 
    
    FncSave = True                                                              '☜: Processing is OK
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
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call MakeKeyStream("P")
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    
    Call InitVariables										                     '⊙: Initializes local global variables

    LayerShowHide(1)
		
    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & "P"	                             '☜: Direction

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
    Set gActiveElement = document.ActiveElement
      
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
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO, "X", "X")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call MakeKeyStream("N")
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    
    Call InitVariables										                     '⊙: Initializes local global variables

    LayerShowHide(1)
		
    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & "N"	                             '☜: Direction

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
    Set gActiveElement = document.ActiveElement
        
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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'⊙: Data is changed.  Do you want to exit? 
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
Function DbQuery(KeyItemVal)
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status
	
	On Error Resume Next
	
    DbQuery = False                                                              '☜: Processing is NG

'	LayerShowHide(1)
	
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	If frm1.txtPrevNext.value = "" Then
		If CommonQueryRs(" ITEM_CD "," b_item_image "," ITEM_CD = " & FilterVar(KeyItemVal, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) = False Then
			Call DisplayMsgBox("122900", "X", "X", "X")
			document.all.ImgItemImage.src= DIR_INIT_FILE
			frm1.txtPath.focus 
			Set gActiveElement = document.ActiveElement 
			Exit Function
		End If
	End If
		
	strVal = "../../ComAsp/CPictRead.asp" & "?txtKeyValue=" & KeyItemVal		  '☜: query key
	strVal = strVal     & "&txtDKeyValue=" & "default"                            '☜: default value
	strVal = strVal     & "&txtTable="     & "b_item_image"                       '☜: Table Name
	strVal = strVal     & "&txtField="     & "item_image"	                      '☜: Field
	strVal = strVal     & "&txtKey="       & "item_cd"	                          '☜: Key
	
	document.all.ImgItemImage.src = ValueEscape(strVal)
	
	lgIntFlgMode = parent.OPMD_UMODE
	
	Call SetToolbar("11111000110001")
    
    DbQuery = True                                                               '☜: Processing is NG

End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
	
	On Error Resume Next
		
	DbSave = False														         '☜: Processing is NG
		
	LayerShowHide(1)
		
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	lgIntFlgMode = parent.OPMD_UMODE
	With Frm1
		.txtMode.value        = parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID1)
		
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
		
	LayerShowHide(1)
		
	strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0003                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
    Set gActiveElement = document.ActiveElement
    
    DbDelete = True                                                              '⊙: Processing is NG
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      = parent.OPMD_UMODE
	
    Frm1.txtItemCd.focus 

	Call SetToolbar("11111000111001")

    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   

End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	
    Call InitVariables
    frm1.txtPath.value = NULL
    
    Call MainQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

	Call InitVariables()
	Call FncNew()	
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtItemCd.value)	' Plant Code
	arrParam(1) = ""							' Item Code
	arrParam(2) = ""							' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field명(0) : "ITEM_CD"
    arrField(1) = 2 							' Field명(1) : "ITEM_NM"
    
	iCalledAspName = AskPRAspName("B1B01PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B01PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItemCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus
	
End Function

'------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemCd(byval arrRet)
	frm1.txtItemCd.Value    = arrRet(0)		
	frm1.txtItemNm.Value    = arrRet(1)
	
	frm1.txtItemCd.focus
	Set gActiveElement = document.activeElement 
	
End Function

'------------------------------------------  GetItemCd()  --------------------------------------------------
'	Name : GetItemCd()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Function GetItemCd()
	
	Dim strVal
    
    On Error Resume Next                                                                    '☜: Clear err status
    
    Err.Clear

    LayerShowHide(1)
		
    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '☜: Direction

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
    Set gActiveElement = document.ActiveElement
    	
End Function

Sub txtPath_OnChange()
	Dim iStrFileType

	lgBlnFlgChgValue = True
	
    If Not ggoSaveFile.FileExists(frm1.txtPath.value) = 0 Then
		Exit Sub
    End If

    iStrFileType = Right(Trim(UCase(frm1.txtPath.value)), 3)

	If Not (iStrFileType = "BMP" Or iStrFileType = "GIF" Or iStrFileType = "JPG") Then
		Call DisplayMsgBox("122904", "X", "X", "X")
		Exit Sub
	End If
	document.all.ImgItemImage.src= ValueEscape(IMG_LOAD_PATH & frm1.txtPath.value)
End sub
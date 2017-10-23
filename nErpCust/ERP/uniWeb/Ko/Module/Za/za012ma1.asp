<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : BA
*  3. Program ID           : za012ma1
*  4. Program Name         : za012ma1
*  5. Program Desc         : za012ma1
*  6. Comproxy List        : 
*  7. Modified date(First) : 2001/05/21
*  8. Modified date(Last)  : 2001/05/21
*  9. Modifier (First)     : Park Sang Hoon
* 10. Modifier (Last)      : Park Sang Hoon
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">        

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incUni2KTV.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                                    

'=========================================================================================================
Dim  C_USER_MENU 
Dim C_ORGCD
Dim C_ORGNM
Dim C_UsrId
Dim C_USERNM
Dim C_CURRDT
Dim C_UseYn
Dim C_CURRDTTime
Dim C_ORGTYPE
Dim C_hOccurDt
Dim  gDragNode , gDropNode, gPrevNode
Dim  lgBlnBizLoadMenu, lgBlnUserLoadMenu, gMenuDat
Dim  lgBlnLoadMenu

Dim  lgBlnFlgConChg                	'☜: Condition 변경 Flag
Dim  lgBlnFlgTopLeftChange      	'☜: Scroll여부를 나타내는 변수 

Dim  lgQueryFlag
Dim  lgRetFlag
Dim  IsOpenPop                    'Popup

Dim  strMode
Public glOrgCode
Public glOrgName
Public glOrgType

Dim  lgSaveModFg
Dim  TempRootNode

Const  C_CMD_TOP_LEVEL = "LISTTOP"
Const  C_CMD_GP_LEVEL = "LISTGP"
Const  C_CMD_ACCT_LEVEL = "LISTACCT"
Const  C_Root  = "Root"
Const  C_USER_MENU_KEY = "$"
Const  C_USER_MENU_STR = "UM_"
Const  C_UNDERBAR = "_"
Const BIZ_Org_Usr_Id = "za012mb1.asp"                                                
Const C_Sep  = "/"
Const C_IMG_Root = "../../../CShared/image/unierp.gif"
Const C_IMG_Org1 = "../../../CShared/image/Orglvl_1.gif"
Const C_IMG_Org1_Open = "../../../CShared/image/Orglvl_1.gif"
Const C_IMG_Org2 = "../../../CShared/image/Orglvl_2.gif"
Const C_IMG_Org2_Open = "../../../CShared/image/Orglvl_2.gif"

C_USER_MENU= CStr(Parent.gCompanyNm)

<!-- #Include file="../../inc/lgvariables.inc" -->    

Sub InitSpreadPosVariables()
    C_ORGCD   = 1
    C_ORGNM   = 2
    C_UsrId  = 3
    C_USERNM  = 4
    C_CURRDT  = 5
    C_UseYn   = 6
    C_CURRDTTime = 7
    C_ORGTYPE = 8
    C_hOccurDt = 9
End Sub

'=========================================================================================================
Sub InitVariables()

    lgBlnBizLoadMenu = False
    lgBlnLoadMenu = False
    lgIntFlgMode = parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgBlnFlgTopLeftChange = False   
    lgIntGrpCount = 0                           
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           
    lgLngCurRows = 0                                
End Sub

'=========================================================================================================
Sub  SetDefaultVal()
    lgBlnFlgChgValue = False
End Sub

'=========================================================================================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "*", "NOCOOKIE","QA") %>
End Sub

'=========================================================================================================
Sub  InitSpreadSheet()
        
    Dim sList

    Call InitSpreadPosVariables()
    
    With frm1.vspdData
    
        sList = "Y" & vbTab  & "N"
   
        ggoSpread.Source = frm1.vspdData    
        Call ggoSpread.Spreadinit("V20021124",,Parent.gAllowDragDropSpread)

        .ReDraw = false    
        .MaxCols = C_hOccurDt + 1                              
        .MaxRows = 0

        Call GetSpreadColumnPos("A")        
  
        ggoSpread.SSSetEdit        C_ORGCD, "조직 코드", 12,,,10
        ggoSpread.SSSetEdit        C_ORGNM, "조직명", 20,,,30
        ggoSpread.SSSetEdit        C_UsrId, "사용자 ID", 13,,,13
        ggoSpread.SSSetEdit        C_USERNM, "사용자명", 18,,,30
        ggoSpread.SSSetDate        C_CURRDT, "소속일자", 15, 2, parent.gDateFormat   
        ggoSpread.SSSetCheck    C_UseYn, "현재조직여부", 20, 2, "", True      
        ggoSpread.SSSetEdit        C_OrgType, "Org Type", 11,2,,10
        ggoSpread.SSSetEdit     C_CURRDTTime, "Time", 15, 2
        ggoSpread.SSSetEdit     C_hOccurDt, "", 15, 2

        .ReDraw = true

        Call SetSpreadLock("I", 0, 1, "")
        
        Call ggoSpread.MakePairsColumn(C_ORGCD,C_ORGNM,"1")

        Call ggoSpread.SSSetColHidden(C_CURRDTTime,C_CURRDTTime,True)
        Call ggoSpread.SSSetColHidden(C_ORGTYPE,C_ORGTYPE,True)
        Call ggoSpread.SSSetColHidden(C_hOccurDt,C_hOccurDt,True)
        Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
        
    End With  
    
End Sub
'=========================================================================================================
Sub  SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2 )
    Dim objSpread
    
    With frm1
    
    Select Case Index
        Case 0
            ggoSpread.Source = .vspdData
            Set objSpread = .vspdData
    End Select
    
    If lRow2 = "" Then lRow2 = objSpread.MaxRows
    
    objSpread.Redraw = False    
    
    Select Case stsFg
        Case "Q"
            Select Case Index
                Case 0
                    ggoSpread.SpreadLock C_ORGCD, -1, C_ORGCD
                    ggoSpread.SpreadLock C_ORGNM, -1, C_ORGNM
                    ggoSpread.SpreadLock C_UsrId, -1, C_UsrId
                    ggoSpread.SpreadLock C_USERNM, -1, C_USERNM
                    ggoSpread.SpreadLock C_CURRDT, -1, C_CURRDT
                       
             End Select
             
        Case "I"
            Select Case Index
                Case 0
                    ggoSpread.SpreadLock C_ORGCD, -1, C_ORGCD
                    ggoSpread.SpreadLock C_ORGNM, -1, C_ORGNM
                    ggoSpread.SpreadLock C_UsrId, -1, C_UsrId
                    ggoSpread.SpreadLock C_USERNM, -1, C_USERNM
                    ggoSpread.SpreadLock C_CURRDT, -1, C_CURRDT
            End Select
    End Select
    
    objSpread.Redraw = True
    Set objSpread = Nothing
    
    End With
    
End Sub

'=========================================================================================================
Sub  SetSpreadColor(ByVal lRow)
    
    With frm1.vspdData 
        .Redraw = False
        ggoSpread.Source = frm1.vspdData
        ggoSpread.SSSetProtected C_ORGCD, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_ORGNM, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_UsrId, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_USERNM, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected C_CURRDT, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired C_UseYn, pvStartRow, pvEndRow

        Call SetActiveCell(frm1.vspdData,1,.ActiveRow,"M","X","X")

        .EditMode = True
        .Redraw = True
    End With
End Sub
'=========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
            C_ORGCD      =  iCurColumnPos(1)
            C_ORGNM      =  iCurColumnPos(2)
            C_UsrId      =  iCurColumnPos(3)
            C_USERNM     =  iCurColumnPos(4)
            C_CURRDT     =  iCurColumnPos(5)            
            C_UseYn      =  iCurColumnPos(6)
            C_CURRDTTime =  iCurColumnPos(7)
            C_ORGTYPE    =  iCurColumnPos(8)
            C_hOccurDt   =  iCurColumnPos(9)
            
    End Select
End Sub

'=========================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> UC_PROTECTED Then
              Call SetActiveCell(Frm1.vspdData,iDx,iRow,"M","X","X")
              Exit For
           End If
       Next
          
    End If   
End Sub

'=========================================================================================================
Function OpenAddUser()

    Dim arrRet, IntRetCD
    Dim Param(2)
    Dim arrField, arrHeader
    Dim iCalledAspName
    
    If glOrgCode = "" Then
          IntRetCD = DisplayMsgBox("211130", "X","X","X")     '조회할 조직을 먼저 선택하십시오.
       'org code level 을 선택하지 않았을 시 Popup 기동하지 않음.
       Exit function
    End if
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("ZA012RA1")

    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZA012RA1", "x")
        IsOpenPop = False
        Exit Function
    End If
        
    Param(0) = CStr(frm1.txtOrgType.value)
    Param(1) = CStr(frm1.txtOrgCd.value)
    'Param =  & parent.gColSep & CStr(frm1.txtOrgCd.value)

    arrRet = window.showModalDialog(iCalledAspName,Array(Window.parent,Param, arrField, arrHeader), _
        "dialogWidth=580px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0, 0) = "" Then
        Exit Function
    Else
        Call SetRefAddUser(arrRet)
    End If

End Function

'=========================================================================================================
Function SetRefAddUser(Byval arrRet)
    
    Dim intRtnCnt, strData
    Dim TempRow, I
    
    With frm1
    
        .vspdData.focus
        lgBlnFlgChgValue = True
        ggoSpread.Source = .vspdData
        .vspdData.ReDraw = False    
    
        TempRow = .vspdData.MaxRows                                                
        .vspdData.MaxRows = .vspdData.MaxRows + (Ubound(arrRet, 1) + 1)            
 
        For I = TempRow to .vspdData.MaxRows - 1
        	.vspdData.SetText 0, I+1, ggoSpread.InsertFlag
        	.vspdData.SetText C_OrgCd, I+1, glOrgCode
        	.vspdData.SetText C_OrgNm, I+1, glOrgName
        	.vspdData.SetText C_UsrId, I+1, arrRet(I - TempRow, 0)
        	.vspdData.SetText C_UserNm, I+1, arrRet(I - TempRow, 1)
        	.vspdData.SetText C_UseYn, I+1, "1"
        	.vspdData.SetText C_OrgType, I+1, glOrgType
        Next    
        
        .vspdData.ReDraw = True
    End With

    Call SetToolbar("1000100100011111")                                        
    
End Function

'=========================================================================================================
'   Function Name : GetNodeLvl
'   Function Desc : 현재 노드의 Level을 찾는다.
'=========================================================================================================
Function  GetNodeLvl(Node)
    
    Dim tempNode
    
    Set tempNode = Node
    
    GetNodeLvl = 0
    
    if tempNode.Key <> "$" Then
        Do        
            GetNodeLvl = GetNodeLvl + 1
            Set tempNode = tempNode.Parent
        Loop Until tempNode.Key = "$"
    End If
    
    Set tempNode = Nothing

End Function


'=========================================================================================================
'   Function Name :GetIndex
'   Function Desc :Node가 부모의 몇번째 위치인가 되돌려준다.
'=========================================================================================================
Function GetIndex(Node)
    Dim i, myIndx,  ChildNode, ParentNode
    
    Set ParentNode = Node.Parent
    
    If ParentNode is Nothing Then    ' Root일 경우 
        GetIndex = 1
        Exit Function
    End If
    
    Set ChildNode = ParentNode.Child
    
    myIndx = 1
    
    For i = 1 to ParentNode.Children
        
        If ChildNode.Key = Node.Key Then
            GetIndex = myIndx
            Exit Function
        End If
        
        If Node.Image = ChildNode.Image Then
            myIndx = myIndx + 1
        End if            
        
        Set ChildNode = ChildNode.Next
    Next
    
End Function

'=========================================================================================================
'   Function Name :GetInsSeq
'   Function Desc : 현재 Insert 되는 Node의 순서를 리턴한다.
'=========================================================================================================
Function GetInsSeq(Node)
    Dim i, myIndx,  ChildNode, ParentNode

    Set ChildNode = Node.Child
    
    myIndx = 1
    
    For i = 1 to Node.Children
        If gDragNode.Image = ChildNode.Image Then
            myIndx = myIndx + 1
        End if            
        Set ChildNode = ChildNode.Next
    Next
    
    GetInsSeq = myIndx
    
End Function

'=========================================================================================================
'   Function Name :GetTotalCnt
'   Function Desc :Add에 관련되 자식수를 되돌려준다.
'=========================================================================================================
Function GetTotalCnt(Node)
    
    If Node.children = 0 Then    ' Root일 경우 
        GetTotalCnt = 1
    Else
        GetTotalCnt = Node.children + 1
    End If
    
End Function

Sub DispDivConf(pVal) 
    if pVal = 2 then
        divconf.style.display = "none"
        tdConf.height = 1
    else
        divconf.style.display = ""
        tdConf.height = 22
    end if
End Sub

Sub MenuRefresh()
    
    if lgBlnBizLoadMenu = False Then
        Call DisplayAcct()
    End If
    
End Sub

'=========================================================================================================
'    메뉴를 읽어 TreeView에 넣음 
'=========================================================================================================
Sub  DisplayAcct()

    Dim NodX

  
    frm1.uniTree1.Nodes.Clear 
    Set NodX = frm1.uniTree1.Nodes.Add(, tvwChild, C_USER_MENU_KEY, C_USER_MENU, C_Root, C_Root)
    
    Call SetDefaultVal()
        
    frm1.uniTree1.MousePointer = 11
    
    Call InitNodes()
        
End Sub

'=========================================================================================================
Function DisplayAcctOK()
    Dim NodX

    Set NodX = frm1.uniTree1.Nodes(C_USER_MENU_KEY)
        
    If Not (nodX.child Is Nothing) Then
        Call uniTree1_NodeClick(nodX.child)        
    End If
    
End Function

'=========================================================================================================
' Function Name : GetImage
' Function Desc : 이미지 정보 
'=========================================================================================================
Function GetImage(Byval arrLine)
    Dim strImg
    Select Case arrLine(C_MNU_AUTH)
        Case "A"
            If arrLine(C_MNU_TYPE) = "M" Then
                strImg = C_Folder
            Else
                strImg = C_URL
            End If
        Case "I"
            strImg = C_Const
        Case "N"
            strImg = C_None
    End Select
    GetImage = strImg
End Function

'=========================================================================================================
Sub  Form_Load()

    Call LoadInfTB19029                                                         
    Call ggoOper.LockField(Document, "N")                                   
    '----------  Coding part  -------------------------------------------------------------
    Call InitSpreadSheet                                                    
    Call InitVariables                                                      
    Call SetToolbar("1000000000001111")                                                         

    With frm1
        .unitree1.HideSelection = false    
        .uniTree1.SetAddImageCount = 5 
        .uniTree1.Indentation = "200"    ' 줄 간격 
                        ' 파일위치,    키명, 위치 
        .uniTree1.AddImage C_IMG_Root, C_Root, 0                    '⊙: TreeView에 보일 이미지 지정 
        .uniTree1.AddImage C_IMG_Org1, C_Folder, 0                    '⊙: TreeView에 보일 이미지 지정 
        .uniTree1.AddImage C_IMG_Org1_Open, C_Open, 0
        .uniTree1.AddImage C_IMG_Org2, C_URL, 0
        .uniTree1.AddImage C_IMG_Org2_Open, C_None, 0
    
        .uniTree1.OLEDragMode = 0                                    '⊙: Drag & Drop 을 가능하게 할 것인가 정의 
        .uniTree1.OLEDropMode = 0                                    ' Drag & Drop 불가 

    End With    
    Set gDragNOde = Nothing

End Sub

'=========================================================================================================
Sub  Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
    Call SetPopupMenuItemInf("0001111111")    
    
    gMouseClickStatus = "SPC"   
    
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows <= 0 Then                                                    
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

    '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'=========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)                
    If Row <= 0 Then
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    '------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
    
End Sub

'=========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)        
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'=========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 
    '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub
'=========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'=========================================================================================================
Sub  vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
End Sub
    
'=========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If CheckRunningBizProcess = True Then
       Exit Sub
    End If

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    lgBlnFlgTopLeftChange = True
        
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)  And lgStrPrevKey <> "" Then        
        Call uniTree1_NodeClick(frm1.uniTree1.selecteditem)
    End if
    
    lgBlnFlgTopLeftChange = False
    
End Sub

'=========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'=========================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
    Call InitData()
End Sub


'=========================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'=========================================================================================================
'   Event Name : uniTree1_NodeClick
'   Event Desc : Node를 클릭하면 발생 이벤트 
'=========================================================================================================
Sub uniTree1_NodeClick(Node)

    DIm IntRetCD

    
    If ggoSpread.SSCheckChange = True then
       IntRetCD = DisplayMsgBox("800442", parent.VB_YES_NO,"X","X")

       If IntRetCD = vbYes Then          
		  Call FncSave()
		  Exit Sub
       End If

    End If

    If lgBlnFlgTopLeftChange = False Then
        Call ggoOper.ClearField(Document, "2")                                        
        Call ggoSpread.ClearSpreadData()        
    End If
    
    ggoSpread.Source = frm1.vspdData

    Dim strVal
    Dim arg

    Call LayerShowHide(1)

    If Node.Image = C_URL Then                
'
        arg = Split(Node.key,"::", -1, 1)

        glOrgType = arg(0) ' org type
        glOrgCode = arg(1) ' org code
        glOrgName = arg(2) ' org name
        
        frm1.txtOrgType.value = arg(0)        'For 현재조직에 지금 속해있는 사용자를 제외한 사용자쿼리 
        frm1.txtOrgCd.value = arg(1)        'For 현재조직에 지금 속해있는 사용자를 제외한 사용자쿼리 
                
        strVal = BIZ_Org_Usr_Id & "?txtMode=" & parent.UID_M0001                                
        strVal = strVal & "&strCmd=SHOW" 
        strVal = strVal & "&strType=" & arg(0) 
        strVal = strVal & "&strCd=" & arg(1)
        strVal = strVal & "&txtMaxRows=" & frm1.vspdData.MaxRows
        
        If lgBlnFlgTopLeftChange = True Then
            strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey    
        End If
            
        Call SetToolbar("1000100100011111")                                                             
        Call RunMyBizASP(MyBizASP, strVal)
    Else
		glOrgCode = ""
    End If

    lgBlnFlgChgValue = False
        
    Call LayerShowHide(0)

End Sub

'=========================================================================================================
'   Event Name : InitNodes
'   Event Desc : TreeView 초기화 
'=========================================================================================================
Sub InitNodes()
    
    Dim strVal

    strVal = BIZ_Org_Usr_Id & "?txtMode=" & parent.UID_M0001    
    strVal = strVal & "&strCmd=INIT"
    
    Call RunMyBizASP(MyBizASP, strVal)

End sub

'=========================================================================================================
'    Event  Name : uniTree1_onAddImgReady()
'    Description : SetAddImageCount수의 Image가 다운로드 완료되고 TreeView의 ImageList에 
'                 추가되면 발생하는 이벤트 
'=========================================================================================================
Sub uniTree1_onAddImgReady()
    if lgBlnBizLoadMenu = False Then
        Call DisplayAcct()        
    End If
End Sub

'=========================================================================================================
Function  FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               


    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")                '데이타가 변경되었습니다. 조회하시겠습니까?
        If IntRetCD = vbNo Then
              Exit Function
        End If
    End If    
    

    Call InitVariables
    Call InitSpreadSheet                                                
    

    If DbQuery = False Then
       Exit Function
    End If
       
    FncQuery = True                                                                
    
End Function

'=========================================================================================================
Function  FncSave() 
    Dim IntRetCD    
    
    FncSave = False                                                         
    
    Err.Clear                                                               
    On Error Resume Next                                                    
    

    ggoSpread.Source = frm1.vspdData

    If lgBlnFlgChgValue = False and ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X","X","X")                          
        Exit Function
    End If 


    If DbSave = False Then
       Exit Function
    End If
    
    FncSave = True                                                          
    
End Function

'=========================================================================================================
Function  FncCopy() 
    If frm1.vspdData.Maxrows < 1 Then Exit Function
    frm1.vspdData.ReDraw = False
    
    ggoSpread.Source = frm1.vspdData    
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    
    frm1.vspdData.ReDraw = True
End Function

'=========================================================================================================
Function  FncCancel() 

    if frm1.vspdData.Maxrows < 1 Then Exit Function
    
    ggoSpread.Source = frm1.vspdData    
    ggoSpread.EditUndo                                                  

End Function

'=========================================================================================================
Function  FncInsertRow() 

End Function

'=========================================================================================================
Function  FncDeleteRow()
End Function
'=========================================================================================================
Function  FncPrint() 
    On Error Resume next                                                    
    parent.FncPrint()
End Function

'=========================================================================================================
Function  FncPrev() 
    On Error Resume next                                                    
End Function

'=========================================================================================================
Function  FncNext() 
    On Error Resume next                                                    
End Function

'=========================================================================================================
Function  FncExcel() 
    Call parent.FncExport(parent.C_MULTI)
End Function

'=========================================================================================================
Function  FncFind() 
    Call parent.FncFind(parent.C_MULTI , True)                          
End Function

'=========================================================================================================
Function  FncExit()
    Dim IntRetCD
    FncExit = False
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")                '데이타가 변경되었습니다. 조회하시겠습니까?
            If IntRetCD = vbNo Then
                  Exit Function
            End If
    End If

    FncExit = True
End Function

'=========================================================================================================
Function  DbQuery() 

    Err.Clear                                                               

    DbQuery = False

    Call LayerShowHide(1)
    Call DisplayAcct()        
    
    DbQuery = True    

End Function
'=========================================================================================================
Function DbQueryOk()                                                        
    
    lgIntFlgMode = parent.OPMD_UMODE                                                        
    
    SetSpreadLock "Q", 0, 1, ""
    
    Call ggoOper.LockField(Document, "Q")                                    
    
    Call LayerShowHide(0)
End Function

'=========================================================================================================
Function DbSave()     
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
    Dim strVal
    Dim iColSep, iRowSep
    iColSep = parent.gColSep
    iRowSep = parent.gRowSep
    
    On Error Resume next                                                   

    Call LayerShowHide(1)
    
    DbSave = False                                                          
    
    
    lgRetFlag = False

    With frm1
        .txtMode.value = parent.UID_M0002
    

        lGrpCnt = 1
        
        strVal = ""
        

        For lRow = 1 To .vspdData.MaxRows
            Select Case GetSpreadText(.vspdData, 0, lRow, "X", "X")
                Case ggoSpread.InsertFlag                                            
					strVal = strVal & "C"                      & iColSep        '
					strVal = strVal & lRow                     & iColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData, C_OrgCd, lRow, "X", "X"))      & iColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData, C_OrgNm, lRow, "X", "X"))      & iColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData, C_UsrId, lRow, "X", "X"))      & iColSep   
					If CInt(Trim(GetSpreadText(.vspdData, C_UseYn, lRow, "X", "X"))) = 1 Then
					strVal = strVal & "Y"      & iColSep
					Else
					strVal = strVal & "N"      & iColSep
					End If
					strVal = strVal & Trim(GetSpreadText(.vspdData, C_OrgType, lRow, "X", "X"))      & iRowSep

                    lGrpCnt = lGrpCnt + 1
                Case ggoSpread.UpdateFlag                                            
					strVal = strVal & "U"                      & iColSep        '
					strVal = strVal & lRow                     & iColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData, C_OrgCd, lRow, "X", "X"))      & iColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData, C_OrgNm, lRow, "X", "X"))      & iColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData, C_UsrId, lRow, "X", "X"))      & iColSep   
					If CInt(Trim(GetSpreadText(.vspdData, C_UseYn, lRow, "X", "X"))) = 1 Then
					strVal = strVal & "Y"      & iColSep
					Else
					strVal = strVal & "N"      & iColSep
					End If
					strVal = strVal & Trim(GetSpreadText(.vspdData, C_OrgType, lRow, "X", "X"))      & iColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData, C_hOccurDt, lRow, "X", "X"))      & iRowSep                    

                    lGrpCnt = lGrpCnt + 1
            End Select
                    
        Next
        
        .txtMaxRows.value = lGrpCnt-1
        .txtSpread.value = strVal

        Call ExecMyBizASP(frm1, BIZ_Org_Usr_Id)                                        
    
    End With
    
    DbSave = True                                                           
    lgRetFlag = True

End Function
'=========================================================================================================
Function DbSaveOk()                                                    

    lgBlnFlgChgValue = False
     ggoSpread.ssdeleteflag 1

    frm1.uniTree1.Setfocus
    
    lgStrPrevKey = ""            
    call uniTree1_NodeClick(frm1.uniTree1.selecteditem)

    Call LayerShowHide(0)
  
    frm1.uniTree1.MousePointer = 0
    
    'Call ggoOper.ClearField(Document, "2")                                            

End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->    
</HEAD>
<BODY TABINDEX="-1" SCROLL="no" oncontextmenu="javascript:return false">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">

<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
    <TR>
        <TD <%=HEIGHT_TYPE_00%>></TD>
    </TR>
    <TR HEIGHT=23>
        <TD WIDTH="100%" HEIGHT="100%">
            <TABLE <%=LR_SPACE_TYPE_10%>>
                <TR>
                <!-- TreeView AREA -->
                    <TD HEIGHT=* WIDTH=20%>                     <%' <TD CLASS="CLSMTABP">로 하지 말 것!! %>
                        <script language =javascript src='./js/za012ma1_uniTree1_N976313573.js'></script>
                    </TD>
                    <!-- DATA AREA -->
                    <TD WIDTH="70%" HEIGHT="100%">
                        <TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
                            <TR HEIGHT=23>
                                <TD WIDTH="100%">
                                    <TABLE <%=LR_SPACE_TYPE_10%>>
                                        <TR>
                                            <TD WIDTH=10>&nbsp;</TD>
                                            <TD CLASS="CLSMTABP">
                                                    <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 >    
                                                    <TR>
                                                        <td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG height=23 src="../../../CShared/image/table/seltab_up_left.gif" width=9></td>
                                                        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>조직별 사용자 관리</font></td>
                                                        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG height=23 src="../../../CShared/image/table/seltab_up_right.gif" width=10></td>
                                                    </TR>
                                                </TABLE>                                                
                                            </TD>        
                                            <TD WIDTH=* align=right><A href="VBScript:OpenAddUser()">사용자 추가</A></TD>
                                            <TD WIDTH=10>&nbsp;</TD>                                                                                                                                                                                                                            
                                        </TR>
                                    </TABLE>
                                </TD>
                            </TR>

                            <TR>
                                <TD WIDTH="100%" CLASS="Tab11">
									<script language =javascript src='./js/za012ma1_vaSpread1_vspdData.js'></script>
                                </TD>
                            </TR>
                        </TABLE>
                    </TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR>
        <TD WIDTH="100%" HEIGHT=<%=BizSize%>>
            <IFRAME NAME="MyBizASP" SRC="" WIDTH="100%" HEIGHT=20 FRAMEBORDER=1 SCROLLING=no noresize framespacing=0>assd</IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="24">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24">
<INPUT TYPE=hidden NAME="lgstrCmd" tag="24">
<INPUT TYPE=hidden NAME="txtlgMode" tag="24">
<INPUT TYPE=hidden NAME="txtOrgType" tag="24">
<INPUT TYPE=hidden NAME="txtOrgCd" tag="24">
<INPUT TYPE=hidden NAME="txtParentGP_CD" tag="21">
<INPUT TYPE=hidden NAME="txtParentGP_LVL" tag="21">
<INPUT TYPE=hidden NAME="txtParentGP_SEQ" tag="21">
<INPUT TYPE=hidden NAME="txtToParentGP_CD" tag="21">
<INPUT TYPE=hidden NAME="txtToParentGP_LVL" tag="21">
<INPUT TYPE=hidden NAME="txtToParentGP_SEQ" tag="21">
<INPUT TYPE=hidden NAME="txtToGP_CD" tag="21">
<INPUT TYPE=hidden NAME="txtToGP_LVL" tag="21">
<INPUT TYPE=hidden NAME="txtToGP_SEQ" tag="21">

<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</FORM>
</BODY>
</HTML>


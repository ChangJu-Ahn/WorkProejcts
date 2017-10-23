
<!--
======================================================================================================
*  1. Module Name          : BA
*  2. Function Name        : 
*  3. Program ID           : 
*  4. Program Name         : 
*  5. Program Desc         :
*  6. Comproxy List        : 
*  7. Modified date(First) : 1999/09/10
*  8. Modified date(Last)  : 1999/09/10
*  9. Modifier (First)     : Lee JaeHoo
* 10. Modifier (Last)      : Lee JaeHoo
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit     


Const BIZ_PGM_ID  = "zc055mb1.asp"
Const JUMP_PGM_ID = "zc050ma1"

Dim C_Mnu_ID              
Dim C_Mnu_Nm              
Dim C_USER_ID             
Dim C_USER_NM             
Dim C_BIZ_AREA_CD_ALL     
Dim C_BIZ_AREA_CD         
Dim C_BIZ_AREA_POPUP      
Dim C_BIZ_AREA_NM         
Dim C_INTERNAL_CD_ALL     
Dim C_INTERNAL_CD         
Dim C_INTERNAL_POPUP      
Dim C_INTERNAL_NM         
Dim C_SUB_INTERNAL_CD_ALL 
Dim C_SUB_INTERNAL_CD     
Dim C_SUB_INTERNAL_POPUP  
Dim C_SUB_INTERNAL_NM     
Dim C_PERSONAL_ID_ALL     
Dim C_PERSONAL_ID         
Dim C_PERSONAL_POPUP      
Dim C_PERSONAL_NM         
Dim C_PLANT_CD_ALL        
Dim C_PLANT_CD            
Dim C_PLANT_POPUP         
Dim C_PLANT_NM            
Dim C_PUR_ORG_CD_ALL      
Dim C_PUR_ORG_CD          
Dim C_PUR_ORG_POPUP       
Dim C_PUR_ORG_NM          
Dim C_PUR_GRP_CD_ALL      
Dim C_PUR_GRP_CD          
Dim C_PUR_GRP_POPUP       
Dim C_PUR_GRP_NM          
Dim C_SALES_ORG_CD_ALL    
Dim C_SALES_ORG_CD        
Dim C_SALES_ORG_POPUP     
Dim C_SALES_ORG_NM        
Dim C_SALES_GRP_CD_ALL    
Dim C_SALES_GRP_CD        
Dim C_SALES_GRP_POPUP     
Dim C_SALES_GRP_NM        
Dim C_SL_CD_ALL           
Dim C_SL_CD               
Dim C_SL_POPUP            
Dim C_SL_NM               
Dim C_WC_CD_ALL           
Dim C_WC_CD               
Dim C_WC_POPUP            
Dim C_WC_NM               
    
Dim C_ALLOW_YN            
Dim C_BIZ_AREA_YN         
Dim C_INTERNAL_YN         
Dim C_SUB_INTERNAL_YN     
Dim C_PERSONAL_YN         
Dim C_PLANT_YN            
Dim C_PUR_ORG_YN          
Dim C_PUR_GRP_YN          
Dim C_SALES_ORG_YN        
Dim C_SALES_GRP_YN        
Dim C_SL_YN               
Dim C_WC_YN   
            
Dim C_DUMMY               

<!-- #Include file="../../inc/lgvariables.inc" -->    

Dim IsOpenPop
'=========================================================================================================
Sub InitSpreadPosVariables()

    C_Mnu_ID              = 1 
    C_Mnu_Nm              = 2 
    C_USER_ID             = 3 
    C_USER_NM             = 4 
    C_BIZ_AREA_CD_ALL     = 5 
    C_BIZ_AREA_CD         = 6 
    C_BIZ_AREA_POPUP      = 7 
    C_BIZ_AREA_NM         = 8 
    C_INTERNAL_CD_ALL     = 9 
    C_INTERNAL_CD         = 10
    C_INTERNAL_POPUP      = 11
    C_INTERNAL_NM         = 12
    C_SUB_INTERNAL_CD_ALL = 13
    C_SUB_INTERNAL_CD     = 14
    C_SUB_INTERNAL_POPUP  = 15
    C_SUB_INTERNAL_NM     = 16
    C_PERSONAL_ID_ALL     = 17
    C_PERSONAL_ID         = 18
    C_PERSONAL_POPUP      = 19
    C_PERSONAL_NM         = 20
    C_PLANT_CD_ALL        = 21
    C_PLANT_CD            = 22
    C_PLANT_POPUP         = 23
    C_PLANT_NM            = 24
    C_PUR_ORG_CD_ALL      = 25
    C_PUR_ORG_CD          = 26
    C_PUR_ORG_POPUP       = 27
    C_PUR_ORG_NM          = 28
    C_PUR_GRP_CD_ALL      = 29
    C_PUR_GRP_CD          = 30
    C_PUR_GRP_POPUP       = 31
    C_PUR_GRP_NM          = 32
    C_SALES_ORG_CD_ALL    = 33
    C_SALES_ORG_CD        = 34
    C_SALES_ORG_POPUP     = 35
    C_SALES_ORG_NM        = 36
    C_SALES_GRP_CD_ALL    = 37
    C_SALES_GRP_CD        = 38
    C_SALES_GRP_POPUP     = 39
    C_SALES_GRP_NM        = 40
    C_SL_CD_ALL           = 41
    C_SL_CD               = 42
    C_SL_POPUP            = 43
    C_SL_NM               = 44
    C_WC_CD_ALL           = 45
    C_WC_CD               = 46
    C_WC_POPUP            = 47
    C_WC_NM               = 48
    
    C_ALLOW_YN            = 49
    C_BIZ_AREA_YN         = 50
    C_INTERNAL_YN         = 51
    C_SUB_INTERNAL_YN     = 52
    C_PERSONAL_YN         = 53
    C_PLANT_YN            = 54
    C_PUR_ORG_YN          = 55
    C_PUR_GRP_YN          = 56
    C_SALES_ORG_YN        = 57
    C_SALES_GRP_YN        = 58
    C_SL_YN               = 59
    C_WC_YN               = 60
    C_DUMMY               = 61
    
End Sub

'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE

    lgIntGrpCount = 0
    
    lgStrPrevKey = ""
        
    lgLngCurRows = 0
    lgSortKey = 1    
    
End Sub
'=========================================================================================================
Sub SetDefaultVal()
End Sub

'=========================================================================================================
Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "*", "NOCOOKIE","QA") %>
End Sub

'=========================================================================================================
Sub InitSpreadSheet()

    On Error Resume Next

    Call InitSpreadPosVariables()
    
    With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
        Call ggoSpread.Spreadinit("V20021124",,Parent.gAllowDragDropSpread)

        .ReDraw = false                   
        .MaxCols = C_DUMMY
        .MaxRows = 0

        Call GetSpreadColumnPos("A")
        frm1.vspdData.RowHeight(0) = 22
       
        ggoSpread.SSSetEdit   C_Mnu_ID              , "메뉴ID"     , 12
        ggoSpread.SSSetEdit   C_Mnu_Nm              , "메뉴명"     , 18
        ggoSpread.SSSetEdit   C_USER_ID             , "사용자"     ,  8
        ggoSpread.SSSetEdit   C_USER_NM             , "사용자명"   , 12

        ggoSpread.SSSetCheck  C_BIZ_AREA_CD_ALL     , "전체"       ,  6,,,True
        ggoSpread.SSSetEdit	  C_BIZ_AREA_CD         , "사업장"     , 12
        ggoSpread.SSSetButton C_BIZ_AREA_POPUP
        ggoSpread.SSSetEdit   C_BIZ_AREA_NM         , "사업장명"   , 16
        
        ggoSpread.SSSetCheck  C_INTERNAL_CD_ALL     , "전체"       ,  6,,,True
        ggoSpread.SSSetEdit	  C_INTERNAL_CD         , "내부부서"   , 10
        ggoSpread.SSSetButton C_INTERNAL_POPUP
        ggoSpread.SSSetEdit   C_INTERNAL_NM         , "내부부서명" , 15
        
        ggoSpread.SSSetCheck  C_SUB_INTERNAL_CD_ALL , "전체"       ,  6,,,True
        ggoSpread.SSSetEdit	  C_SUB_INTERNAL_CD     , "내부부서"   & vbcrlf  & "(하위포함)", 11
        ggoSpread.SSSetButton C_SUB_INTERNAL_POPUP
        ggoSpread.SSSetEdit   C_SUB_INTERNAL_NM     , "내부부서명" & vbcrlf  & "(하위포함)", 15
        
        ggoSpread.SSSetCheck  C_PERSONAL_ID_ALL     , "전체"      ,  6,,,True
        ggoSpread.SSSetEdit   C_PERSONAL_ID         , "개인"      ,  8
        ggoSpread.SSSetButton C_PERSONAL_POPUP
        ggoSpread.SSSetEdit   C_PERSONAL_NM         , "개인명"    , 12
        
        ggoSpread.SSSetCheck  C_PLANT_CD_ALL        , "전체"      ,  6,,,True
        ggoSpread.SSSetEdit   C_PLANT_CD            , "공장"      ,  7
        ggoSpread.SSSetButton C_PLANT_POPUP
        ggoSpread.SSSetEdit   C_PLANT_NM            , "공장명"    , 10
        
        ggoSpread.SSSetCheck  C_PUR_ORG_CD_ALL      , "전체"      ,  6,,,True
        ggoSpread.SSSetEdit   C_PUR_ORG_CD          , "구매조직"  , 10
        ggoSpread.SSSetButton C_PUR_ORG_POPUP
        ggoSpread.SSSetEdit   C_PUR_ORG_NM          , "구매조직명", 15
        
        ggoSpread.SSSetCheck  C_PUR_GRP_CD_ALL      , "전체"       , 6,,,True
        ggoSpread.SSSetEdit   C_PUR_GRP_CD          , "구매그룹"   ,10
        ggoSpread.SSSetButton C_PUR_GRP_POPUP
        ggoSpread.SSSetEdit   C_PUR_GRP_NM          , "구매그룹명" ,14
        
        ggoSpread.SSSetCheck  C_SALES_ORG_CD_ALL    , "전체"       , 6,,,True
        ggoSpread.SSSetEdit   C_SALES_ORG_CD        , "영업조직"   , 8
        ggoSpread.SSSetButton C_SALES_ORG_POPUP
        ggoSpread.SSSetEdit   C_SALES_ORG_NM        , "영업조직명" ,10
        
        ggoSpread.SSSetCheck  C_SALES_GRP_CD_ALL    , "전체"       ,  6,,,True
        ggoSpread.SSSetEdit   C_SALES_GRP_CD        , "영업그룹"   ,  8
        ggoSpread.SSSetButton C_SALES_GRP_POPUP
        ggoSpread.SSSetEdit   C_SALES_GRP_NM        , "영업그룹명" , 10
        
        ggoSpread.SSSetCheck  C_SL_CD_ALL           , "전체"      ,  6,,,True
        ggoSpread.SSSetEdit   C_SL_CD               , "창고"      ,  8
        ggoSpread.SSSetButton C_SL_POPUP
        ggoSpread.SSSetEdit   C_SL_NM               , "창고명"    , 12
        
        ggoSpread.SSSetCheck  C_WC_CD_ALL           , "전체"      ,  6,,,True
        ggoSpread.SSSetEdit   C_WC_CD               , "작업장"    ,  7
        ggoSpread.SSSetButton C_WC_POPUP
        ggoSpread.SSSetEdit   C_WC_NM               , "작업장명"  , 15
         
        ggoSpread.SSSetEdit   C_ALLOW_YN          , "allow"       ,10
        ggoSpread.SSSetEdit   C_BIZ_AREA_YN       , "biz_area"    ,10
        ggoSpread.SSSetEdit   C_INTERNAL_YN       , "internal"    ,10
        ggoSpread.SSSetEdit   C_SUB_INTERNAL_YN   , "sub_internal",10
        ggoSpread.SSSetEdit   C_PERSONAL_YN         , "usr_id"      ,10
        ggoSpread.SSSetEdit   C_PLANT_YN          , "plant"       ,10
        ggoSpread.SSSetEdit   C_PUR_ORG_YN        , "pur_org"     ,10
        ggoSpread.SSSetEdit   C_PUR_GRP_YN        , "pur_grp"     ,10
        ggoSpread.SSSetEdit   C_SALES_ORG_YN      , "sales_org"   ,10
        ggoSpread.SSSetEdit   C_SALES_GRP_YN      , "sales_grp"   ,10
        ggoSpread.SSSetEdit   C_SL_YN             , "sl"          ,10
        ggoSpread.SSSetEdit   C_WC_YN             , "wc"          ,10
        
		Call ggoSpread.MakePairsColumn(C_BIZ_AREA_CD    , C_BIZ_AREA_POPUP     )
		Call ggoSpread.MakePairsColumn(C_INTERNAL_CD    , C_INTERNAL_POPUP     )
		Call ggoSpread.MakePairsColumn(C_SUB_INTERNAL_CD, C_SUB_INTERNAL_POPUP )
		Call ggoSpread.MakePairsColumn(C_PERSONAL_ID    , C_PERSONAL_POPUP     )
		Call ggoSpread.MakePairsColumn(C_PLANT_CD       , C_PLANT_POPUP        )
		Call ggoSpread.MakePairsColumn(C_PUR_ORG_CD     , C_PUR_ORG_POPUP      )
		Call ggoSpread.MakePairsColumn(C_PUR_GRP_CD     , C_PUR_GRP_POPUP      )
		Call ggoSpread.MakePairsColumn(C_SALES_ORG_CD   , C_SALES_ORG_POPUP    )
		Call ggoSpread.MakePairsColumn(C_SALES_GRP_CD   , C_SALES_GRP_POPUP    )
		Call ggoSpread.MakePairsColumn(C_SL_CD          , C_SL_POPUP           )
		Call ggoSpread.MakePairsColumn(C_WC_CD          , C_WC_POPUP           )

        .ReDraw = true

        Call SetSpreadLock

        Call ggoSpread.SSSetColHidden(C_ALLOW_YN          ,C_ALLOW_YN        , True)
        Call ggoSpread.SSSetColHidden(C_BIZ_AREA_YN       ,C_BIZ_AREA_YN     , True)
        Call ggoSpread.SSSetColHidden(C_INTERNAL_YN       ,C_INTERNAL_YN     , True)
        Call ggoSpread.SSSetColHidden(C_SUB_INTERNAL_YN   ,C_SUB_INTERNAL_YN , True)
        Call ggoSpread.SSSetColHidden(C_PERSONAL_YN       ,C_PERSONAL_YN     , True)
        Call ggoSpread.SSSetColHidden(C_PLANT_YN          ,C_PLANT_YN        , True)
        Call ggoSpread.SSSetColHidden(C_PUR_ORG_YN        ,C_PUR_ORG_YN      , True)
        Call ggoSpread.SSSetColHidden(C_PUR_GRP_YN        ,C_PUR_GRP_YN      , True)
        Call ggoSpread.SSSetColHidden(C_SALES_ORG_YN      ,C_SALES_ORG_YN    , True)
        Call ggoSpread.SSSetColHidden(C_SALES_GRP_YN      ,C_SALES_GRP_YN    , True)
        Call ggoSpread.SSSetColHidden(C_SL_YN             ,C_SL_YN           , True)
        Call ggoSpread.SSSetColHidden(C_WC_YN             ,C_WC_YN           , True)

        Call ggoSpread.SSSetColHidden(C_DUMMY, C_DUMMY, True)

    End With
    
    
End Sub
'=========================================================================================================
Sub SetSpreadLock()
    Dim iLoop
    
    With frm1
    
        .vspdData.ReDraw = False
        
        For iLoop = 1 To C_DUMMY
            ggoSpread.SpreadLock iLoop            , -1, iLoop
        Next
        
'        ggoSpread.SpreadLock C_Mnu_ID            , -1, C_Mnu_ID
'        ggoSpread.SpreadLock C_Mnu_Nm            , -1, C_Mnu_Nm
'        ggoSpread.SpreadLock C_BIZ_AREA_NM      , -1, C_BIZ_AREA_NM
'        ggoSpread.SpreadLock C_INTERNAL_NM      , -1, C_INTERNAL_NM
'        ggoSpread.SpreadLock C_SUB_INTERNAL_NM  , -1, C_SUB_INTERNAL_NM
'        ggoSpread.SpreadLock PERSONAL_NM          , -1, PERSONAL_NM
'        ggoSpread.SpreadLock C_PLANT_NM         , -1, C_PLANT_NM
'        ggoSpread.SpreadLock C_PUR_ORG_NM       , -1, C_PUR_ORG_NM
'        ggoSpread.SpreadLock C_PUR_GRP_NM       , -1, C_PUR_GRP_NM
'        ggoSpread.SpreadLock C_SALES_ORG_NM     , -1, C_SALES_ORG_NM
'        ggoSpread.SpreadLock C_SALES_GRP_NM     , -1, C_SALES_GRP_NM
'        ggoSpread.SpreadLock C_SL_NM            , -1, C_SL_NM
'        ggoSpread.SpreadLock C_WC_NM            , -1, C_WC_NM

        .vspdData.ReDraw = True    

    End With
End Sub
'=========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SSSetRequired        C_Mnu_ID , pvStartRow, pvEndRow        
        ggoSpread.SSSetRequired        C_USER_ID, pvStartRow, pvEndRow
        ggoSpread.SpreadLock    C_Mnu_Nm           , -1, C_Mnu_Nm
        ggoSpread.SpreadLock    C_USER_NM          , -1, C_USER_NM
        ggoSpread.SpreadLock    C_BIZ_AREA_NM      , -1, C_USER_NM        
        ggoSpread.SpreadLock    C_INTERNAL_NM      , -1, C_BIZ_AREA_NM    
        ggoSpread.SpreadLock    C_SUB_INTERNAL_NM  , -1, C_INTERNAL_NM    
        ggoSpread.SpreadLock    C_PERSONAL_NM      , -1, C_SUB_INTERNAL_NM
        ggoSpread.SpreadLock    C_PLANT_NM         , -1, C_PERSONAL_NM    
        ggoSpread.SpreadLock    C_PUR_ORG_NM       , -1, C_PLANT_NM       
        ggoSpread.SpreadLock    C_PUR_GRP_NM       , -1, C_PUR_ORG_NM     
        ggoSpread.SpreadLock    C_SALES_ORG_NM     , -1, C_PUR_GRP_NM     
        ggoSpread.SpreadLock    C_SALES_GRP_NM     , -1, C_SALES_ORG_NM   
        ggoSpread.SpreadLock    C_SL_NM            , -1, C_SALES_GRP_NM   
        ggoSpread.SpreadLock    C_WC_NM            , -1, C_SL_NM          
        .vspdData.ReDraw = True
    End With
End Sub

'=========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
            C_Mnu_ID                      = iCurColumnPos(1 )
            C_Mnu_Nm                      = iCurColumnPos(2 )
            C_USER_ID                     = iCurColumnPos(3 )
            C_USER_NM                     = iCurColumnPos(4 )
            C_BIZ_AREA_CD_ALL             = iCurColumnPos(5 )
            C_BIZ_AREA_CD                 = iCurColumnPos(6 )
            C_BIZ_AREA_POPUP              = iCurColumnPos(7 )
            C_BIZ_AREA_NM                 = iCurColumnPos(8 )
            C_INTERNAL_CD_ALL             = iCurColumnPos(9 )
            C_INTERNAL_CD                 = iCurColumnPos(10)
            C_INTERNAL_POPUP              = iCurColumnPos(11)
            C_INTERNAL_NM                 = iCurColumnPos(12)
            C_SUB_INTERNAL_CD_ALL         = iCurColumnPos(13)
            C_SUB_INTERNAL_CD             = iCurColumnPos(14)
            C_SUB_INTERNAL_POPUP          = iCurColumnPos(15)
            C_SUB_INTERNAL_NM             = iCurColumnPos(16)
            C_PERSONAL_ID_ALL             = iCurColumnPos(17)
            C_PERSONAL_ID                 = iCurColumnPos(18)
            C_PERSONAL_POPUP              = iCurColumnPos(19)
            C_PERSONAL_NM                 = iCurColumnPos(20)
            C_PLANT_CD_ALL                = iCurColumnPos(21)
            C_PLANT_CD                    = iCurColumnPos(22) 
            C_PLANT_POPUP                 = iCurColumnPos(23) 
            C_PLANT_NM                    = iCurColumnPos(24) 
            C_PUR_ORG_CD_ALL              = iCurColumnPos(25) 
            C_PUR_ORG_CD                  = iCurColumnPos(26) 
            C_PUR_ORG_POPUP               = iCurColumnPos(27) 
            C_PUR_ORG_NM                  = iCurColumnPos(28) 
            C_PUR_GRP_CD_ALL              = iCurColumnPos(29) 
            C_PUR_GRP_CD                  = iCurColumnPos(30) 
            C_PUR_GRP_POPUP               = iCurColumnPos(31) 
            C_PUR_GRP_NM                  = iCurColumnPos(32) 
            C_SALES_ORG_CD_ALL            = iCurColumnPos(33) 
            C_SALES_ORG_CD                = iCurColumnPos(34) 
            C_SALES_ORG_POPUP             = iCurColumnPos(35) 
            C_SALES_ORG_NM                = iCurColumnPos(36) 
            C_SALES_GRP_CD_ALL            = iCurColumnPos(37) 
            C_SALES_GRP_CD                = iCurColumnPos(38)                   
            C_SALES_GRP_POPUP             = iCurColumnPos(39) 
            C_SALES_GRP_NM                = iCurColumnPos(40) 
            C_SL_CD_ALL                   = iCurColumnPos(41) 
            C_SL_CD                       = iCurColumnPos(42) 
            C_SL_POPUP                    = iCurColumnPos(43)
            C_SL_NM                       = iCurColumnPos(44)
            C_WC_CD_ALL                   = iCurColumnPos(45)
            C_WC_CD                       = iCurColumnPos(46)
            C_WC_POPUP                    = iCurColumnPos(47)
            C_WC_NM                       = iCurColumnPos(48)
            C_ALLOW_YN                    = iCurColumnPos(49)
            C_BIZ_AREA_YN                 = iCurColumnPos(50)                   
            C_INTERNAL_YN                 = iCurColumnPos(51)
            C_SUB_INTERNAL_YN             = iCurColumnPos(52)
            C_PERSONAL_YN                 = iCurColumnPos(53)
            C_PLANT_YN                    = iCurColumnPos(54)
            C_PUR_ORG_YN                  = iCurColumnPos(55)
            C_PUR_GRP_YN                  = iCurColumnPos(56)
            C_SALES_ORG_YN                = iCurColumnPos(57)
            C_SALES_GRP_YN                = iCurColumnPos(58)
            C_SL_YN                       = iCurColumnPos(59)
            C_WC_YN                       = iCurColumnPos(60)
            C_DUMMY                       = iCurColumnPos(61)

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
              Frm1.vspdData.Action = 0 
              Exit For
           End If
           
       Next
          
    End If   
End Sub

'=========================================================================================================
Sub InitComboBox()

End Sub

'=========================================================================================================
Sub InitSpreadComboBox()

    Dim strCboData
    Dim IntRetCD

    ggoSpread.Source = frm1.vspdData

End Sub
'=========================================================================================================
Sub Form_Load()

    Dim IntRetCD
    
    On Error Resume Next
    
    Call ggoOper.LockField(Document, "N")

    Call InitSpreadSheet
    
    Call InitVariables

    Call InitComboBox
    Call InitSpreadComboBox
    Call SetDefaultVal
    Call SetToolbar("11001000001111")
    
    frm1.txtMnuID.focus
        
	Set gActiveElement = document.activeElement
    
    IntRetCD = CommonQueryRs("Lang_Nm","B_LANGUAGE","Lang_Cd =  " & FilterVar(Parent.gLang , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)        
    lgF0 = Replace(lgF0, Chr(11), "")

    frm1.txtLangNm.value = Trim(lgF0)
    
End Sub
'=========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
End Sub

'=========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
    Call SetPopupMenuItemInf("1101111111")    
    
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


Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1
	
	If Row <= 0 Then
	   Exit Sub
	End If
   
	With frm1.vspdData 
 		ggoSpread.Source = frm1.vspdData
 		
 		Select Case Col
              Case C_BIZ_AREA_POPUP     :  Call OpenBizAreaCd   (    GetSpreadText(frm1.vspdData, C_BIZ_AREA_CD     , Row, "X", "X"))
              Case C_INTERNAL_POPUP     :  Call OpenDeptOrgPopup("A",GetSpreadText(frm1.vspdData, C_INTERNAL_CD     , Row, "X", "X"))
              Case C_SUB_INTERNAL_POPUP :  Call OpenDeptOrgPopup("B",GetSpreadText(frm1.vspdData, C_SUB_INTERNAL_CD , Row, "X", "X"))
              Case C_PERSONAL_POPUP     :  Call OpenUSER        (    GetSpreadText(frm1.vspdData, C_PERSONAL_ID     , Row, "X", "X"))
              Case C_PLANT_POPUP        :  Call OpenPlant       (    GetSpreadText(frm1.vspdData, C_PLANT_CD        , Row, "X", "X"))
              Case C_PUR_ORG_POPUP      :  Call OpenPurOrgGrp   ("O",GetSpreadText(frm1.vspdData, C_PUR_ORG_CD      , Row, "X", "X"))
              Case C_PUR_GRP_POPUP      :  Call OpenPurOrgGrp   ("G",GetSpreadText(frm1.vspdData, C_PUR_GRP_CD      , Row, "X", "X"))
              Case C_SALES_ORG_POPUP    :  Call OpenSaleOrgGrp  ("O",GetSpreadText(frm1.vspdData, C_SALES_ORG_CD    , Row, "X", "X"))
              Case C_SALES_GRP_POPUP    :  Call OpenSaleOrgGrp  ("G",GetSpreadText(frm1.vspdData, C_SALES_GRP_CD    , Row, "X", "X"))
              Case C_SL_POPUP           :  Call OpenSL          (    GetSpreadText(frm1.vspdData, C_SL_CD           , Row, "X", "X")) 
              Case C_WC_POPUP           :  Call OpenWorkCenter  (    GetSpreadText(frm1.vspdData, C_WC_CD           , Row, "X", "X")) 

		End Select 
		     	
	End With
End Sub




'=========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    
    If CheckRunningBizProcess = True Then
       Exit Sub
    End If
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) _
    And Not(lgStrPrevKey = "") Then    
        Call DisableToolBar(Parent.TBC_QUERY)
        If DBQuery = False Then
            Call RestoreToolBar()
            Exit Sub
        End If 
    End if
    
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
    Call InitSpreadComboBox
    Call ggoSpread.ReOrderingSpreadData()
    Call InitData()
End Sub
'=========================================================================================================
Function FncQuery()

    Dim IntRetCD 
    
    FncQuery = False
    
    Err.Clear

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
          Exit Function
        End If
    End If
    
    Call ggoOper.ClearField(Document, "2")
    Call ggoSpread.ClearSpreadData()            
    Call InitVariables
                                                                
    If Not chkField(Document, "1") Then
       Exit Function
    End If

    If DbQuery = False Then
       Exit Function
    End If
       
    FncQuery = True
    
End Function
'=========================================================================================================
Function FncNew() 
End Function
'=========================================================================================================
Function FncDelete() 
End Function
'=========================================================================================================
Function FncSave() 
    Dim IntRetCD 
   
    FncSave = False
    
    Err.Clear
    On Error Resume Next
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001", "x", "x", "x")
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then
       Exit Function
    End If
    
    If DbSave = False Then
       Exit Function
    End If
    
    FncSave = True
    
End Function
'=========================================================================================================
Function FncCopy() 
	Dim nActiveRow
    With frm1.vspdData
        If .ActiveRow > 0 Then
            .focus
            .ReDraw = False
            
            ggoSpread.Source = frm1.vspdData 
            ggoSpread.CopyRow
            nActiveRow = frm1.vspdData.ActiveRow
            SetSpreadColor nActiveRow, nActiveRow
    
    		frm1.vspdData.SetText C_Mnu_ID, nActiveRow, ""
            .ReDraw = True
        End If
    End With
End Function
'=========================================================================================================
Function FncCancel() 
    ggoSpread.EditUndo
End Function
'=========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    
    On Error Resume Next                                                              
    Err.Clear                                                                     
    
    FncInsertRow = False                                                             

    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else
        imRow = AskSpdSheetAddRowCount()
        If imRow = "" Then
            Exit Function
        End If
    End If
    
    With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
    
    '------ Developer Coding part (Start ) -------------------------------------------------------------- 

    '------ Developer Coding part (End )   -------------------------------------------------------------- 

    If Err.number = 0 Then
       FncInsertRow = True                                                              
    End If   
    
    Set gActiveElement = document.ActiveElement   
End Function
'=========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows

    With frm1.vspdData 
        .focus
        ggoSpread.Source = frm1.vspdData 
        lDelRows = ggoSpread.DeleteRow
    End With
End Function
'=========================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function
'=========================================================================================================
Function FncPrev() 
End Function
'=========================================================================================================
Function FncNext() 
End Function
'=========================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function
'=========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)
End Function


'=========================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'=========================================================================================================
Function FncExit()
Dim IntRetCD
    FncExit = False
    
    ggoSpread.Source = frm1.vspdData    
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
            Exit Function
        End If
    End If
    FncExit = True
End Function
'=========================================================================================================
Function DbQuery() 

    Dim IntRetCD
    
    DbQuery = False
    
    Call LayerShowHide(1)    
    
    Err.Clear

    Dim strVal    
    
    With frm1
    
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then    
        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
        
        strVal = strVal & "&txtMnuID=" & Trim(.hMnuID.value)
        strVal = strVal & "&txtUsrID=" & Trim(.hUsrID.value)
        strVal = strVal & "&cboMnuType=P"
        strVal = strVal & "&cboUseYN=1"   ' 1 : 사용  , 0 : 미사용 
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    Else    
        strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
        
        strVal = strVal & "&txtMnuID=" & Trim(.txtMnuID.value)
        strVal = strVal & "&txtUsrID=" & Trim(.txtUsrID.value)
        strVal = strVal & "&cboMnuType=P"
        strVal = strVal & "&cboUseYN=1"
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
    
        strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    End If   
    
    
    Call RunMyBizASP(MyBizASP, strVal)
    
    End With
       
    
    DbQuery = True
    
End Function
'=========================================================================================================
Function DbQueryOk()
    
    lgIntFlgMode = Parent.OPMD_UMODE
    
    Call ggoOper.LockField(Document, "Q")
    Call SetToolbar("11001000001111")
    
    Call AutoHWidth(frm1.vspdData)

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
    Dim strVal, strDel
    Dim iColSep, iRowSep
    Dim IntRetCD
    
    iColSep = parent.gColSep
    iRowSep = parent.gRowSep
        
    DbSave = False
    
    Call LayerShowHide(1)    
    
    On Error Resume Next

    With frm1
        .txtMode.value        = Parent.UID_M0002
        .txtUpdtUserId.value  = Parent.gUsrID
        .txtInsrtUserId.value = Parent.gUsrID
        
        lGrpCnt = 1
    
        strVal = ""
        strDel = ""
        
        For lRow = 1 To .vspdData.MaxRows
            Select Case GetSpreadText(.vspdData, 0, lRow, "X", "X")

                Case ggoSpread.UpdateFlag

                    strVal = strVal & "U"                                                                    & iColSep   '0
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_Mnu_ID              , lRow, "X", "X")) & iColSep   '1
                    strVal = strVal & "P"                                                                    & iColSep   '2
                    strVal = strVal & Trim(GetSpreadText(.vspdData, C_USER_ID             , lRow, "X", "X")) & iColSep   '1

                    If  Trim(GetSpreadText(.vspdData, C_BIZ_AREA_YN     , lRow, "X", "X")) = "1" Then
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_BIZ_AREA_CD_ALL     , lRow, "X", "X")) & iColSep   '3
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_BIZ_AREA_CD         , lRow, "X", "X")) & iColSep   '3
                    Else
                        strVal = strVal & ""                                                                     & iColSep   '10
                        strVal = strVal & ""                                                                     & iColSep   '10
                    End If    

                    If  Trim(GetSpreadText(.vspdData, C_INTERNAL_YN     , lRow, "X", "X")) = "1" Then
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_INTERNAL_CD_ALL     , lRow, "X", "X")) & iColSep   '4
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_INTERNAL_CD         , lRow, "X", "X")) & iColSep   '4
                    Else
                        strVal = strVal & ""                                                                     & iColSep   '10
                        strVal = strVal & ""                                                                     & iColSep   '10
                    End If    

                    If  Trim(GetSpreadText(.vspdData, C_SUB_INTERNAL_YN     , lRow, "X", "X")) = "1" Then
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_SUB_INTERNAL_CD_ALL , lRow, "X", "X")) & iColSep   '5
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_SUB_INTERNAL_CD     , lRow, "X", "X")) & iColSep   '5
                    Else
                        strVal = strVal & ""                                                                     & iColSep   '10
                        strVal = strVal & ""                                                                     & iColSep   '10
                    End If    

                    If  Trim(GetSpreadText(.vspdData, C_PERSONAL_YN     , lRow, "X", "X")) = "1" Then
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_PERSONAL_ID_ALL     , lRow, "X", "X")) & iColSep   '6
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_PERSONAL_ID         , lRow, "X", "X")) & iColSep   '6
                    Else
                        strVal = strVal & ""                                                                     & iColSep   '10
                        strVal = strVal & ""                                                                     & iColSep   '10
                    End If    

                    If  Trim(GetSpreadText(.vspdData, C_PLANT_YN     , lRow, "X", "X")) = "1" Then
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_PLANT_CD_ALL        , lRow, "X", "X")) & iColSep   '7
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_PLANT_CD            , lRow, "X", "X")) & iColSep   '7
                    Else
                        strVal = strVal & ""                                                                     & iColSep   '10
                        strVal = strVal & ""                                                                     & iColSep   '10
                    End If    

                    If  Trim(GetSpreadText(.vspdData, C_PUR_ORG_YN     , lRow, "X", "X")) = "1" Then
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_PUR_ORG_CD_ALL      , lRow, "X", "X")) & iColSep   '8
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_PUR_ORG_CD          , lRow, "X", "X")) & iColSep   '8
                    Else
                        strVal = strVal & ""                                                                     & iColSep   '10
                        strVal = strVal & ""                                                                     & iColSep   '10
                    End If    

                    If  Trim(GetSpreadText(.vspdData, C_PUR_GRP_YN     , lRow, "X", "X")) = "1" Then
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_PUR_GRP_CD_ALL      , lRow, "X", "X")) & iColSep   '9
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_PUR_GRP_CD          , lRow, "X", "X")) & iColSep   '9
                    Else
                        strVal = strVal & ""                                                                     & iColSep   '10
                        strVal = strVal & ""                                                                     & iColSep   '10
                    End If    
                    
                    If  Trim(GetSpreadText(.vspdData, C_SALES_ORG_YN     , lRow, "X", "X")) = "1" Then
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_SALES_ORG_CD_ALL    , lRow, "X", "X")) & iColSep   '10
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_SALES_ORG_CD        , lRow, "X", "X")) & iColSep   '10
                    Else
                        strVal = strVal & ""                                                                     & iColSep   '10
                        strVal = strVal & ""                                                                     & iColSep   '10
                    End If    
                    
                    If  Trim(GetSpreadText(.vspdData, C_SALES_GRP_YN     , lRow, "X", "X")) = "1" Then
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_SALES_GRP_CD_ALL    , lRow, "X", "X")) & iColSep   '11
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_SALES_GRP_CD        , lRow, "X", "X")) & iColSep   '11
                    Else
                        strVal = strVal & ""                                                                     & iColSep   '10
                        strVal = strVal & ""                                                                     & iColSep   '10
                    End If    

                    If  Trim(GetSpreadText(.vspdData, C_SL_YN     , lRow, "X", "X")) = "1" Then
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_SL_CD_ALL           , lRow, "X", "X")) & iColSep   '12
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_SL_CD               , lRow, "X", "X")) & iColSep   '12
                    Else
                        strVal = strVal & ""                                                                     & iColSep   '10
                        strVal = strVal & ""                                                                     & iColSep   '10
                    End If    

                    If  Trim(GetSpreadText(.vspdData, C_WC_YN     , lRow, "X", "X")) = "1" Then
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_WC_CD_ALL           , lRow, "X", "X")) & iColSep   '13
                        strVal = strVal & Trim(GetSpreadText(.vspdData, C_WC_CD               , lRow, "X", "X")) & iColSep   '13
                    Else
                        strVal = strVal & ""                                                                     & iColSep   '10
                        strVal = strVal & ""                                                                     & iColSep   '10
                    End If    

                    strVal = strVal & lRow                                                                   & iRowSep   '14

                    '---------------------------------------------------------------------------------------
                    
                    lGrpCnt = lGrpCnt + 1

            End Select
                    
        Next
        
        .txtMaxRows.value = lGrpCnt-1
        .txtSpread.value = strDel & strVal
    
        Call ExecMyBizASP(frm1, BIZ_PGM_ID)
    
    End With
    
    DbSave = True
    
End Function
'=========================================================================================================
Function DbSaveOk()
   
    Call InitVariables
    frm1.vspdData.MaxRows = 0    
    
    Call MainQuery()

End Function

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    
    If IsNumeric(iPosArr) Then
       iRow = CInt(iPosArr)
       
       If iRow <=0 Then
          Exit Sub
       End if
       
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If           
       Next          
    End If   
End Sub



'=========================================================================================================
Function DbDelete() 
End Function

'=========================================================================================================

Function CheckNumeric(ByVal strNum) 
  Dim Ret
  Dim intlen, intCnt, intAsc

  intlen = len(strNum)

  for intCnt = 1 to intlen

      intAsc = asc(mid(strNum, intCnt, 1))

      if intAsc < 48 or intAsc > 57  then
         CheckNumeric = 1
         Exit function
      end if
  next

End Function
'=========================================================================================================
Function OpenRefMenu()

    Dim arrRet
    Dim arrParam(2)
    Dim arrField, arrHeader
    Dim iCalledAspName, IntRetCD
    
    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    iCalledAspName = AskPRAspName("ZC004RA1")
    
    If Trim(iCalledAspName) = "" Then
        IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "ZC004RA1", "x")
        IsOpenPop = False
        Exit Function
    End If
    
    With frm1
        If .vspdData.ActiveRow > 0 then 
            arrParam(0) = GetSpreadText(.vspdData, 3, .vspdData.ActiveRow, "X", "X")
            arrParam(1) = GetSpreadText(.vspdData, 4, .vspdData.ActiveRow, "X", "X")
        End If            
    End With    

    arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam, arrField, arrHeader), _
             "dialogWidth=550px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
End Function

'=========================================================================================================
'    Name : OpenUsrId()
'    Description : User PopUp
'=========================================================================================================
Function OpenUsrId()
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "사용자정보 팝업"                                     ' 팝업 명칭 
    arrParam(1) = "z_usr_mast_rec"                                          ' TABLE 명칭 
    arrParam(2) = frm1.txtUsrId.value                                       ' Code Condition
    arrParam(3) = ""                                                        ' Name Cindition
    arrParam(4) = ""                                                        ' Where Condition
    arrParam(5) = "사용자 ID"            
    
    arrField(0) = "Usr_id"                                                  ' Field명(0)
    arrField(1) = "Usr_nm"                                                  ' Field명(1)
    
    arrHeader(0) = "사용자"                                                ' Header명(0)
    arrHeader(1) = "사용자명"                                           ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp",  Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
        frm1.txtUsrId.value = arrRet(0)
        frm1.txtUsrNm.value = arrRet(1)
    End If    
	frm1.txtUsrId.focus
	Set gActiveElement = document.activeElement

End Function



'=========================================================================================================
Function OpenLangInfo(Byval strCode)'khy200307

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "언어코드 팝업"
    arrParam(1) = "B_LANGUAGE"
    arrParam(2) = Trim(strCode)
    arrParam(3) = ""
    arrParam(4) = ""
    arrParam(5) = "언어 코드"
    
    arrField(0) = "LANG_CD"
    arrField(1) = "LANG_NM"
    
    arrHeader(0) = "언어코드"
    arrHeader(1) = "언어명"
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    
    
    If arrRet(0) = "" Then    
        Exit Function
    Else
        Call SetLangInfo(arrRet)
    End If    

End Function 

Function SetLangInfo(Byval arrRet)
	Dim nActiveRow

    With frm1.vspdData
    	nActiveRow = .ActiveRow
    	.SetText C_LangCD, nActiveRow, arrRet(0)
        Call vspdData_Change(C_LangCD, nActiveRow)
    End With

End Function
'==============================================================================================================
Function OpenMnuInfo(Byval strCode, Byval iWhere)

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    Dim IntRetCD    
    
    IntRetCD = CommonQueryRs("Lang_Nm","B_LANGUAGE","Lang_Cd =  " & FilterVar(frm1.txtLangCd.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
    lgF0 = Replace(lgF0, Chr(11), "") 'unusual case
    'lgF0 = Replace(lgF0, " ","")    

    If lgF0 = "" then 
        Call DisplayMsgBox("211432", "x", "x", "x")
        frm1.txtLangNm.value = ""        
        frm1.txtLangCd.select
        Exit Function
    End if     

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "메뉴 팝업"
    arrParam(1) = "Z_LANG_CO_MAST_MNU"    
    arrParam(2) = strCode
    arrParam(3) = ""    
    
    Select Case iWhere
            Case  1
                arrParam(4) = "LANG_CD = " & FilterVar(Parent.gLang, "''", "S") & ""
            Case  2
                arrParam(4) = "LANG_CD = " & FilterVar(Parent.gLang, "''", "S") & ""
            Case  3
                arrParam(4) = "LANG_CD = " & FilterVar(Parent.gLang, "''", "S") & " AND MNU_TYPE = " & FilterVar("M", "''", "S") & " "
    End Select
                
    arrParam(5) = "메뉴ID"
    
    arrField(0) = "MNU_ID"
    arrField(1) = "MNU_NM"
    
    arrHeader(0) = "메뉴ID"
    arrHeader(1) = "메뉴명"
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    	If iWhere = 1 Then
		    frm1.txtMnuID.focus
	        Set gActiveElement = document.activeElement
	    End If
        Exit Function
    Else
        Call SetMnuInfo(arrRet, iWhere)
    End If    

    
End Function

'=========================================================================================================
Function SetLangCD(Byval arrRet)
    frm1.txtLangCD.Value    = Trim(arrRet(0))
    frm1.txtLangNm.value    = Trim(arrRet(1))
End Function
'=========================================================================================================
Function SetMnuInfo(Byval arrRet, Byval iWhere)
	Dim nActiveRow
    Select Case iWhere
        Case  1
            frm1.txtMnuID.Value    = arrRet(0)
            frm1.txtMnuNm.Value    = arrRet(1)
            frm1.txtMnuID.focus
            Set gActiveElement = document.activeElement
        Case  2
            With frm1.vspdData
            	nActiveRow = .ActiveRow
            	.SetText C_Mnu_ID, nActiveRow, arrRet(0)
            	.SetText C_Mnu_Nm, nActiveRow, arrRet(1)
                Call vspdData_Change(C_Mnu_Nm, nActiveRow)
            End With
        Case  3
            With frm1.vspdData
            	nActiveRow = .ActiveRow
            	.SetText C_UpperMnuID, nActiveRow, arrRet(0)
                Call vspdData_Change(C_UpperMnuID, nActiveRow)
            End With
    End Select

End Function
'=========================================================================================================
Function ProgramJump
    Call PgmJump(JUMP_PGM_ID)
End Function



'=========================================================================================================
Function OpenPlant(byval strCon)
	If IsOpenPop = True Then Exit Function

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	IsOpenPop = True

	arrParam(0) = "공장팝업" 
	arrParam(1) = "B_Plant"    
	arrParam(2) = Trim(strCon)
	arrParam(4) = ""   
	arrParam(5) = "공장"   
	 
	arrField(0) = "Plant_Cd" 
	arrField(1) = "Plant_NM" 
	    
	arrHeader(0) = "공장"  
	arrHeader(1) = "공장명"  
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False

	If arrRet(0) = "" Then 
		Exit Function
	Else
		With frm1.vspdData 
			.Row = .ActiveRow 
			.Col = C_PLANT_CD
			.text = arrRet(0) 
			.Col = C_PLANT_NM
			.text = arrRet(1)
			Call SetFocusToDocument("M") 
            Call vspdData_Change(C_PLANT_CD, .ActiveRow)
			.focus
		End With 
	End If 
End Function

'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사업장 팝업"				' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "사업장 코드"			

    arrField(0) = "BIZ_AREA_CD"					' Field명(0)
    arrField(1) = "BIZ_AREA_NM"					' Field명(1)

    arrHeader(0) = "사업장코드"				' Header명(0)
	arrHeader(1) = "사업장명"				' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1.vspdData 
			.Row = .ActiveRow 
			.Col = C_BIZ_AREA_CD
			.text = arrRet(0) 
			.Col = C_BIZ_AREA_NM
			.text = arrRet(1)
			Call SetFocusToDocument("M") 
            Call vspdData_Change(C_BIZ_AREA_CD, .ActiveRow)
			.focus
		End With 
	End If
End Function

'------------------------------------------  OpenPurGRP()  -------------------------------------------------
'	Name : OpenPurGRP()	구매조직 
'	Description : PurGRP PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPurOrgGrp(ByVal pOP, Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim i_CD
	Dim i_NM

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case pOP
    	Case "O"
                  arrParam(0) = "구매조직팝업"	
                  arrParam(1) = "B_PUR_ORG"				
                  arrParam(2) = Trim(strCode)
                  arrParam(3) = ""
                  arrParam(4) = ""			
                  arrParam(5) = "구매조직"
	
                  arrField(0) = "PUR_ORG"	
                  arrField(1) = "PUR_ORG_NM"	
    
                  arrHeader(0) = "구매조직"		
                  arrHeader(1) = "구매조직명"
                  
        		  i_CD = C_PUR_ORG_CD
                  i_NM = C_PUR_ORG_NM

    	Case "G"
                  arrParam(0) = "구매그룹팝업"	
                  arrParam(1) = "B_PUR_GRP"				
                  arrParam(2) = Trim(strCode)
                  arrParam(3) = ""
                  arrParam(4) = ""			
                  arrParam(5) = "구매조직"
	
                  arrField(0) = "PUR_GRP"	
                  arrField(1) = "PUR_GRP_NM"	
    
                  arrHeader(0) = "구매그룹"		
                  arrHeader(1) = "구매그룹명"

        		  i_CD = C_PUR_GRP_CD
                  i_NM = C_PUR_GRP_NM

    End Select             
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		With frm1.vspdData 
			.Row = .ActiveRow 
			.Col = i_CD
			.text = arrRet(0) 
			.Col = i_NM
			.text = arrRet(1)
			Call SetFocusToDocument("M") 
            Call vspdData_Change(i_CD, .ActiveRow)
			.focus
		End With 
	End If	
	
End Function


'------------------------------------------  OpenPurGRP()  -------------------------------------------------
'	Name : OpenSaleOrgGrp()	영업조직,그룹 
'--------------------------------------------------------------------------------------------------------- 
Function OpenSaleOrgGrp(ByVal pOP, Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim i_CD
	Dim i_NM

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case pOP
    	Case "O"
			arrParam(0) = "영업조직팝업"	
			arrParam(1) = "B_SALES_ORG"
			arrParam(2) = Trim(strCode)
			arrParam(3) = ""											
			arrParam(4) = ""											
			arrParam(5) = "영업조직"							
	
			arrField(0) = "sales_org"						
			arrField(1) = "sales_org_nm"						
    
			arrHeader(0) = "영업조직"				
    		arrHeader(1) = "영업조직명"
    		
    		i_CD = C_SALES_ORG_CD
    		i_NM = C_SALES_ORG_NM
    		
    	Case "G"
			arrParam(0) = "영업그룹팝업"	
			arrParam(1) = "B_SALES_GRP"
			arrParam(2) = Trim(strCode)
			arrParam(3) = ""
			arrParam(4) = ""
			arrParam(5) = "영업그룹"							
	
			arrField(0) = "sales_grp"						
			arrField(1) = "sales_grp_nm"						
    
			arrHeader(0) = "영업그룹"				
    		arrHeader(1) = "영업그룹명"
    		
    		i_CD = C_SALES_GRP_CD
    		i_NM = C_SALES_GRP_NM

    End Select  		
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

	
	If arrRet(0) <> "" Then
		With frm1.vspdData 
			.Row = .ActiveRow 
			.Col = i_CD
			.text = arrRet(0) 
			.Col = i_NM
			.text = arrRet(1)
			Call SetFocusToDocument("M") 
            Call vspdData_Change(i_CD, .ActiveRow)
			.focus
		End With 
	End If	
	
End Function


'------------------------------------------  OpenSL()  -------------------------------------------------
Function OpenSL(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "창고팝업"	
	arrParam(1) = "B_STORAGE_LOCATION"				
	arrParam(2) = Trim(strCode)
	arrParam(3) = ""	
	arrParam(4) = ""
	arrParam(5) = "창고"			
	
	arrField(0) = "SL_CD"	
	arrField(1) = "SL_NM"	
	
	arrHeader(0) = "창고"		
	arrHeader(1) = "창고명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1.vspdData 
			.Row = .ActiveRow 
			.Col = C_SL_CD
			.text = arrRet(0) 
			.Col = C_SL_NM
			.text = arrRet(1)
			Call SetFocusToDocument("M") 
            Call vspdData_Change(C_SL_CD, .ActiveRow)
			.focus
		End With 
	End If	

End Function


'------------------------------------------  OpenWorkCenter()  -------------------------------------------------
'	Name : OpenWorkCenter()	작업장 
'	Description : Work Center Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenWorkCenter(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "작업장팝업"												' 팝업 명칭 
	arrParam(1) = "P_WORK_CENTER"												' TABLE 명칭 
	arrParam(2) = Trim(strCode)								' Code Condition
	arrParam(3) = ""															' Name Cindition
	arrParam(4) = ""										' Where Condition
	arrParam(5) = "작업장"													' TextBox 명칭 
	
    arrField(0) = "WC_CD"														' Field명(0)
    arrField(1) = "WC_NM"														' Field명(1)
    arrField(2) = "INSIDE_FLG"													' Field명(0)
    arrField(3) = "WC_MGR"														' Field명(1)
    arrField(4) = "CAL_TYPE"													' Field명(0)
    
    arrHeader(0) = "작업장"													' Header명(0)
    arrHeader(1) = "작업장명"												' Header명(1)
    arrHeader(2) = "사내외구분"												' Header명(0)
    arrHeader(3) = "작업장담당자"											' Header명(1)
    arrHeader(4) = "칼렌다타입"												' Header명(0)
    
    
	arrRet = window.showModalDialog("../../ComAsp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		With frm1.vspdData 
			.Row = .ActiveRow 
			.Col = C_WC_CD
			.text = arrRet(0) 
			.Col = C_WC_NM
			.text = arrRet(1)
			Call SetFocusToDocument("M") 
            Call vspdData_Change(C_WC_CD, .ActiveRow)
			.focus
		End With 
	End If	
	
End Function


'=========================================================================================================
'    Name : OpenUsrId()
'    Description : User PopUp
'=========================================================================================================
Function OpenUSER(Byval strCode)
    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "사용자정보 팝업"                                     ' 팝업 명칭 
    arrParam(1) = "z_usr_mast_rec"                                          ' TABLE 명칭 
    arrParam(2) = strCode                                                   ' Code Condition
    arrParam(3) = ""                                                        ' Name Cindition
    arrParam(4) = ""                                                        ' Where Condition
    arrParam(5) = "사용자 ID"            
    
    arrField(0) = "Usr_id"                                                  ' Field명(0)
    arrField(1) = "Usr_nm"                                                  ' Field명(1)
    
    arrHeader(0) = "사용자"                                                ' Header명(0)
    arrHeader(1) = "사용자명"                                           ' Header명(1)
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp",  Array(arrParam, arrField, arrHeader), _
        "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
    
    IsOpenPop = False
    
    If arrRet(0) = "" Then
    Else
		With frm1.vspdData 
			.Row = .ActiveRow 
			.Col = C_PERSONAL_ID
			.text = arrRet(0) 
			.Col = C_PERSONAL_NM
			.text = arrRet(1)
			Call SetFocusToDocument("M") 
            Call vspdData_Change(C_PERSONAL_ID, .ActiveRow)
			.focus
		End With 
    End If    
	Set gActiveElement = document.activeElement

End Function

'------------------------------------------  OpenDeptOrgPopup()  ---------------------------------------
'	Name : OpenDeptOrgPopup()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDeptOrgPopup(ByVal pOP,Byval strCode)
	Dim arrRet
	Dim arrParam(4)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = UNIConvDateAToB("<%=GetsvrDate%>",parent.gServerDateFormat,parent.gDateFormat) 
   	arrParam(1) = UNIConvDateAToB("<%=GetsvrDate%>",parent.gServerDateFormat,parent.gDateFormat) 
	arrParam(2) = "" 'lgUsrIntCd               '자료권한 Condition  
	arrParam(3) = ""            
	arrParam(4) = "F"				           '결의일자 상태 Condition  
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
	
       With frm1.vspdData 
        Select Case pOP
	   
	       Case "A"
                  .Row = .ActiveRow 
                  .Col = C_INTERNAL_CD
                  .text = arrRet(3)
                  .Col = C_INTERNAL_NM
                  .text = arrRet(1)
                  Call SetFocusToDocument("M") 
                  Call vspdData_Change(C_INTERNAL_CD, .ActiveRow)
                  .focus
           Case "B"
                  .Row = .ActiveRow 
                  .Col = C_SUB_INTERNAL_CD
                  .text = arrRet(3)
                  .Col = C_SUB_INTERNAL_NM
                  .text = arrRet(1)
                  Call SetFocusToDocument("M") 
                  Call vspdData_Change(C_SUB_INTERNAL_CD, .ActiveRow)
                  .focus
           
        End Select     
       End With
	

	End If	
End Function

Function ProgramJump

    Dim IntRetCD
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "x", "x")
        If IntRetCD = vbNo Then
          Exit Function
        End If
    End If
    Call PgmJump(JUMP_PGM_ID)
    
End Function

Sub AutoHWidth(pSpread)
    Dim iLoop
    
    For iLoop = 1 To pSpread.MaxCols
        if C_BIZ_AREA_POPUP  = iLoop Or  C_INTERNAL_POPUP  = iLoop Or  C_SUB_INTERNAL_POPUP  = iLoop Or  C_PERSONAL_POPUP  = iLoop Or  C_PLANT_POPUP = iLoop Or  C_PUR_ORG_POPUP = iLoop Or  C_PUR_GRP_POPUP = iLoop Or  C_SALES_ORG_POPUP = iLoop Or  C_SALES_GRP_POPUP = iLoop Or  C_SL_POPUP  = iLoop Or  C_WC_POPUP  = iLoop Then
        else
           pSpread.ColWidth(iLoop) = pSpread.MaxTextColWidth(iLoop) + 1
        end if   
    Next

End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->    

</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BasicTB" CELLSPACING=0>
    <TR>
        <TD HEIGHT=5>&nbsp;</TD>
    </TR>
    <TR HEIGHT=23>
        <TD WIDTH=100%>
            <TABLE CLASS="BasicTB" CELLSPACING=0>
                <TR>
                    <TD WIDTH=10>&nbsp;</TD>
                    <TD CLASS="CLSMTABP">
                        <TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
                            <TR>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
                                <td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
                            </TR>
                        </TABLE>
                    </TD>
                    <TD WIDTH=* align=right>&nbsp;</TD>
                    <TD WIDTH=10>&nbsp;</TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR HEIGHT=*>
        <TD WIDTH=100% CLASS="Tab11">
            <TABLE CLASS="BasicTB" CELLSPACING=0>
                <TR>
                    <TD HEIGHT=5 WIDTH=100%></TD>
                </TR>
                <TR>
                    <TD HEIGHT=20 WIDTH=100%>
                    <FIELDSET CLASS="CLSFLD"><TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
                    <TR>
                        <TD CLASS="TD5">메뉴 ID</TD>
                        <TD CLASS="TD6"><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtMnuID" SIZE=15 MAXLENGTH=15 tag="11XXXU"  ALT="메뉴 ID"><IMG SRC="../../../CShared/image/btnPopup.gif"   NAME="btnMnuID" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMnuInfo frm1.txtMnuID.value,1 ">&nbsp;<INPUT TYPE=TEXT NAME="txtMnuNm" SIZE=40 tag="14"></TD>                        
                        <TD CLASS="TD5" NOWRAP>사 용 자</TD>
                        <TD CLASS="TD656"><INPUT TYPE=TEXT NAME="txtUsrId" SIZE=13 MAXLENGTH=13 tag="11XXXU" ALT="사용자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUsrId" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenUsrId()">
                        <INPUT TYPE=TEXT ID="txtUsrNm" NAME="txtUsrNm" size=30 tag="14"></TD>
                    </TR>
                </TABLE></FIELDSET></TD>
            </TR>
            <TR>
                <TD WIDTH=100% HEIGHT=* valign=top><TABLE WIDTH="100%" HEIGHT="100%">
                    <TR>
                        <TD HEIGHT="100%">
                        <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT></TD>
                    </TR></TABLE>
                </TD>
            </TR>
        </TABLE></TD>
    </TR>
    <TR>
        <TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>  
    <TR HEIGHT=20>   
        <TD WIDTH=100%>   
            <TABLE <%=LR_SPACE_TYPE_30%>>   
                <TR>   
                <TD WIDTH=50%>&nbsp;</TD>   
                <TD WIDTH=50%>   
                    <TABLE WIDTH=100%>                           
                        <TD WIDTH=* Align=right><A href="Vbscript:ProgramJump()">프로그램별자료권한(속성)</A>&nbsp;</TD>                                                                                     
                        <TD WIDTH=10>&nbsp;</TD>                           
                    </TABLE>   
                </TD>   
                </TR>   
            </TABLE>   
        </TD>   
    </TR>           
    <TR>
        <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
        </TD>
    </TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hUsrID" tag="24">
<INPUT TYPE=HIDDEN NAME="hMnuID" tag="24">
<INPUT TYPE=HIDDEN NAME="hMnuType" tag="24">
<INPUT TYPE=HIDDEN NAME="hUseYN" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

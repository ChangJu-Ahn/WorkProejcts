<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3112ra5_ko441.asp														*
'*  4. Program Name         : 발주내역참조(입고등록ADO)													*
'*  5. Program Desc         : 구매입고에서 발주참조 
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2007/12/11																*
'*  9. Modifier (First)     : Shin Jin-hyun																*
'* 10. Modifier (Last)      : HAN cheol  																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<Script Language="VBScript">

Option Explicit		
Const BIZ_PGM_ID 		= "m3112rb6_ko441.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 32                                           '☆: key count of SpreadSheet

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim gblnWinEvent
Dim arrReturn					<% '--- Return Parameter Group %>
Dim arrParent
Dim arrParam
Dim EndDate, StartDate
Dim IsOpenPop  

'20071224::hanc
'Dim    C_MVMT_RCPT_NO
'Dim    C_IO_TYPE_CD
'Dim    C_IO_TYPE_NM
'Dim    C_MVMT_RCPT_DT
'Dim    C_BP_CD
'Dim    C_BP_NM
'Dim    C_PUR_GRP
'Dim    C_PUR_GRP_NM
'Dim    C_PLANT_CD

Dim    C_BP_CD
Dim    C_IO_TYPE_NM
Dim    C_IO_TYPE_CD
Dim    C_BP_NM

Dim    C_PO_NO
Dim    C_PO_SEQ_NO
Dim    C_PLANT_CD
Dim    C_SL_CD
Dim    C_ITEM_CD
Dim    C_ITEM_NM
Dim    C_SPEC
Dim    C_TRACKING_NO
Dim    C_PO_QTY

Dim    C_PO_UNIT
Dim    C_PO_PRC
Dim    C_PO_DOC_AMT
Dim    C_PO_CUR
Dim    C_DLVY_DT

Dim    C_RCPT_QTY
Dim    C_LC_QTY
Dim    C_PRE_IV_QTY
Dim    C_INSPECT_QTY
Dim    C_IV_QTY
Dim    C_RECV_INSPEC_FLG
Dim    C_MINOR_NM
Dim    C_INSPECT_METHOD
Dim    C_PLANT_NM
Dim    C_SL_NM
Dim    C_PUR_GRP
Dim    C_LC_RCPT_QTY
Dim    C_LOT_FLG
Dim    C_LOT_GEN_MTHD
Dim    C_FLAG
Dim    C_MVMT_NO

'2008-03-29 7:57오후 :: hanc
Dim    C_TRANS_TIME
Dim    C_MAIN_LOT
Dim    C_IMPORT_TIME
Dim    C_CREATE_TYPE




'================================================================================================================================    
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName
arrParam= arrParent(1)

EndDate = UNIConvDateAtoB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)
'================================================================================================================================
Function InitVariables()
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                        'Indicates that current mode is Create mode
    lgSortKey        = 1
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
						
	frm1.vspdData.MaxRows = 0	
	
	gblnWinEvent = False
	ReDim arrReturn(0,0)
	Self.Returnvalue = arrReturn
End Function
'================================================================================================================================
Sub SetDefaultVal()
	
	Dim iCodeArr
		
	Err.Clear
	
	With frm1
		.txtFrPoDt.text = StartDate
		.txtToPoDt.text = EndDate
	
		.hdnSupplierCd.value 	= arrParam(0)
		.hdnGroupCd.value 		= arrParam(2)
		.txtGroupCd.value 		= arrParam(2)
		.hdnGroupNm.value 		= arrParam(3)
		.txtGroupNm.value 		= arrParam(3)
		.hdnRefType.value 		= arrParam(8)
		.hdnRcptType.value 		= arrParam(9)
		
		.txtPlantCd.value		=  PopupParent.gPlant
		.txtPlantNm.value		=  PopupParent.gPlantNm
	End With
	
	Call CommonQueryRs(" RCPT_FLG", " M_MVMT_TYPE", " IO_TYPE_CD =  " & FilterVar(frm1.hdnRcptType.value, "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    IF Len(lgF0) Then
		iCodeArr = Split(lgF0, Chr(11))
		    
		If Err.number <> 0 Then
			MsgBox Err.description,vbInformation,PopupParent.gLogoName 
			Err.Clear 
			Exit Sub
		End If
		frm1.hdnRcptFlg.value 	= iCodeArr(0)
	End if	

	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtGroupCd, "Q") 
		frm1.txtGroupCd.Tag = left(frm1.txtGroupCd.Tag,1) & "4" & mid(frm1.txtGroupCd.Tag,3,len(frm1.txtGroupCd.Tag))
        frm1.txtGroupCd.value = lgPGCd
	End If
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        frm1.txtPlantCd.value = lgPLCd
	End If
	
End Sub
'================================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>
End Sub
'20080211::hanc =========================================================================================
' Name : InitSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub InitSpreadPosVariables()

'    C_MVMT_RCPT_NO      =   1  
'    C_IO_TYPE_CD        =   2  
'    C_IO_TYPE_NM        =   3  
'    C_MVMT_RCPT_DT      =   4  
'    C_BP_CD             =   5  
'    C_BP_NM             =   6
'    C_PUR_GRP           =   7
'    C_PUR_GRP_NM        =   8
'    C_PLANT_CD          =   9

    C_BP_NM                         =   1  
    C_IO_TYPE_NM                    =   2  
    C_IO_TYPE_CD                    =   3  
    C_BP_CD                         =   4  
    
    C_PO_NO                         =   5  
    C_PO_SEQ_NO                     =   6  
    C_PLANT_CD                      =   7  
    C_SL_CD                         =   8  
    C_ITEM_CD                       =   9  
    C_ITEM_NM                       =   10  
    C_SPEC                          =   11  
    C_TRACKING_NO                   =   12  
    C_PO_QTY                        =   13  
    
    C_PO_UNIT             =   14  
    C_PO_PRC                        =   15  
    C_PO_DOC_AMT                    =   16  
    C_PO_CUR                        =   17  
    C_DLVY_DT                       =   18  

    C_RCPT_QTY                      =   19  
    C_LC_QTY                        =   20 
    C_PRE_IV_QTY                    =   21  
    C_INSPECT_QTY                   =   22  
    C_IV_QTY                        =   23  
    C_RECV_INSPEC_FLG               =   24  
    C_MINOR_NM                      =   25 
    C_INSPECT_METHOD                =   26  
    C_PLANT_NM                      =   27 
    C_SL_NM                         =   28 
    C_PUR_GRP                       =   29 
    C_LC_RCPT_QTY                   =   30 
    C_LOT_FLG                       =   31  
    C_LOT_GEN_MTHD                  =   32 
    C_FLAG                          =   33
    C_MVMT_NO                       =   34
    '2008-03-29 7:58오후 :: hanc
    C_TRANS_TIME                    =   35
    C_MAIN_LOT                      =   36
    C_IMPORT_TIME                   =   37
    C_CREATE_TYPE                   =   38
    
End Sub
'================================================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()   '20080211::hanc

'20080211::hanc	Call SetZAdoSpreadSheet("m3112ra5_ko441","S","A","V20030528",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
'20080211::hanc    Call SetSpreadLock("A")
    frm1.vspdData.OperationMode = 5

	With frm1.vspdData
       ggoSpread.Source = frm1.vspdData
       ggoSpread.Spreadinit "V20021105",, PopupParent.gAllowDragDropSpread
       .ReDraw = false
       .MaxCols   = C_CREATE_TYPE + 1                                                  ' ☜:☜: Add 1 to Maxcols
       Call ggoSpread.ClearSpreadData()
       Call AppendNumberPlace("6","4","2")
       Call GetSpreadColumnPos("A")
 
        ggoSpread.SSSetEdit    C_BP_NM                ,"공급처명"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_IO_TYPE_NM           ,"입고형태"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_IO_TYPE_CD           ,"입고형태"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_BP_CD                ,"공급처"  ,10     ,0     ,     ,100     ,2

       ggoSpread.SSSetEdit    C_PO_NO                ,"발주번호"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_PO_SEQ_NO            ,"발주순번"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_PLANT_CD             ,"공장"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_SL_CD                ,"창고"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_ITEM_CD              ,"품목"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_ITEM_NM              ,"품목명"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_SPEC                 ,"규격"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_TRACKING_NO          ,"Tranking no"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetFloat   C_PO_QTY     , "발주수량" , 10, "C" ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec

       ggoSpread.SSSetEdit    C_PO_UNIT    ,"발주단위"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetFloat   C_PO_PRC     , "단가" , 10, "C" ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
       ggoSpread.SSSetFloat   C_PO_DOC_AMT     , "발주금액" , 10, "C" ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
       ggoSpread.SSSetEdit    C_PO_CUR               ,"화폐"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_DLVY_DT              ,"납기일"  ,10     ,0     ,     ,100     ,2

       ggoSpread.SSSetFloat   C_RCPT_QTY     , "입고량" , 10, "C" ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
       ggoSpread.SSSetFloat   C_LC_QTY     , "LC_QTY" , 10, "C" ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
       ggoSpread.SSSetFloat   C_PRE_IV_QTY     , "PRE_IV_QTY" , 10, "C" ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
       ggoSpread.SSSetFloat   C_INSPECT_QTY     , "검사중수량" , 10, "C" ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
       ggoSpread.SSSetFloat   C_IV_QTY     , "IV_QTY" , 10, "C" ,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
       ggoSpread.SSSetEdit    C_RECV_INSPEC_FLG      ,"RECV_INSPEC_FLG"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_MINOR_NM             ,"MINOR_NM"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_INSPECT_METHOD       ,"INSPECT_METHOD"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_PLANT_NM             ,"PLANT_NM"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_SL_NM                ,"SL_NM"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_PUR_GRP              ,"PUR_GRP"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_LC_RCPT_QTY          ,"LC_RCPT_QTY"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_LOT_FLG              ,"LOT_FLG"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_LOT_GEN_MTHD         ,"LOT_GEN_MTHD"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_FLAG         ,"유무상"  ,10     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_MVMT_NO             ,"MES입고NO"    ,10     ,0     ,     ,100     ,2

       ggoSpread.SSSetEdit    C_TRANS_TIME          ,"TRANS_TIME "  ,25     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_MAIN_LOT            ,"MAIN_LOT   "  ,25     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_IMPORT_TIME         ,"IMPORT_TIME"  ,25     ,0     ,     ,100     ,2
       ggoSpread.SSSetEdit    C_CREATE_TYPE         ,"CREATE_TYPE"  ,25     ,0     ,     ,100     ,2

'       ggoSpread.SSSetEdit    C_MVMT_RCPT_NO    ,"입고번호"  ,10     ,0     ,     ,100     ,2
'       ggoSpread.SSSetEdit    C_IO_TYPE_CD      ,"입고형태"    ,10     ,0     ,     ,100     ,2
'       ggoSpread.SSSetEdit    C_IO_TYPE_NM      ,"입고형태명"    ,10     ,0     ,     ,100     ,2
'       ggoSpread.SSSetEdit    C_MVMT_RCPT_DT    ,"입고일자"  ,10     ,0     ,     ,100     ,2
'       ggoSpread.SSSetEdit    C_BP_CD           ,"공급처"         ,10     ,0     ,     ,100     ,2
'       ggoSpread.SSSetEdit    C_BP_NM           ,"공급처명"         ,10     ,0     ,     ,100     ,2
'       ggoSpread.SSSetEdit    C_PUR_GRP         ,"구매그룹"       ,10     ,0     ,     ,100     ,2
'       ggoSpread.SSSetEdit    C_PUR_GRP_NM      ,"구매그룹명"    ,10     ,0     ,     ,100     ,2
'       ggoSpread.SSSetEdit    C_PLANT_CD        ,"공장"    ,10     ,0     ,     ,100     ,2


       Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)


	   .ReDraw = true
       Call SetSpreadLock 
    End With    

End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_BP_NM                 =   iCurColumnPos(1)
            C_IO_TYPE_NM            =   iCurColumnPos(2)
            C_IO_TYPE_CD            =   iCurColumnPos(3)
            C_BP_CD                 =   iCurColumnPos(4)

            C_PO_NO                 =   iCurColumnPos(5)
            C_PO_SEQ_NO             =   iCurColumnPos(6)
            C_PLANT_CD              =   iCurColumnPos(7)
            C_SL_CD                 =   iCurColumnPos(8)
            C_ITEM_CD               =   iCurColumnPos(9)
            C_ITEM_NM               =   iCurColumnPos(10)
            C_SPEC                  =   iCurColumnPos(11)
            C_TRACKING_NO           =   iCurColumnPos(12)
            C_PO_QTY                =   iCurColumnPos(13)

            C_PO_UNIT               =   iCurColumnPos(14)
            C_PO_PRC                =   iCurColumnPos(15)
            C_PO_DOC_AMT            =   iCurColumnPos(16)
            C_PO_CUR                =   iCurColumnPos(17)
            C_DLVY_DT               =   iCurColumnPos(18)

            C_RCPT_QTY              =   iCurColumnPos(19)
            C_LC_QTY                =   iCurColumnPos(20)
            C_PRE_IV_QTY            =   iCurColumnPos(21)
            C_INSPECT_QTY           =   iCurColumnPos(22)
            C_IV_QTY                =   iCurColumnPos(23)
            C_RECV_INSPEC_FLG       =   iCurColumnPos(24)
            C_MINOR_NM              =   iCurColumnPos(25)
            C_INSPECT_METHOD        =   iCurColumnPos(26)
            C_PLANT_NM              =   iCurColumnPos(27)
            C_SL_NM                 =   iCurColumnPos(28)
            C_PUR_GRP               =   iCurColumnPos(29)
            C_LC_RCPT_QTY           =   iCurColumnPos(30)
            C_LOT_FLG               =   iCurColumnPos(31)
            C_LOT_GEN_MTHD          =   iCurColumnPos(32)  
            C_LOT_GEN_MTHD          =   iCurColumnPos(33)  

'            C_MVMT_RCPT_NO      =   iCurColumnPos(1)
'            C_IO_TYPE_CD        =   iCurColumnPos(2)
'            C_IO_TYPE_NM        =   iCurColumnPos(3)
'            C_MVMT_RCPT_DT      =   iCurColumnPos(4)
'            C_BP_CD             =   iCurColumnPos(5)
'            C_BP_NM             =   iCurColumnPos(6)
'            C_PUR_GRP           =   iCurColumnPos(7)
'            C_PUR_GRP_NM        =   iCurColumnPos(8)
'            C_PLANT_CD          =   iCurColumnPos(9)
            
    End Select    
End Sub

'================================================================================================================================
Sub SetSpreadLock()
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'================================================================================================================================
Function OKClick()

	Dim intColCnt, intRowCnt, intInsRow, i_RowCnt
	Dim before_supplier, curr_supplier, before_MvmtType, curr_MvmtType, curr_FLAG, before_FLAG


    '2008-04-07 11:36오전 :: hanc    
    '참조는 유상일 경우만 가능 (소부장님.. 현업과 협의 사항)
	If frm1.rdoClsFlg(0).checked Then
	ElseIf frm1.rdoClsFlg(1).checked Then
		Msgbox "참조는 유상일 경우만 가능합니다.",vbInformation, parent.gLogoName
		Exit Function
	Else
		Msgbox "참조는 유상일 경우만 가능합니다.",vbInformation, parent.gLogoName
		Exit Function
	End If

		If frm1.vspdData.SelModeSelCount > 0 Then 

			intInsRow = 0
			i_RowCnt        =   0
			before_supplier =   "" 
			curr_supplier   =   "" 
			before_MvmtType =   "" 
			curr_MvmtType   =   ""
			before_FLAG =   "" 
			curr_FLAG   =   "" 

			Redim arrReturn(frm1.vspdData.SelModeSelCount-1, frm1.vspdData.MaxCols - 2)

			For intRowCnt = 1 To frm1.vspdData.MaxRows
				frm1.vspdData.Row = intRowCnt

				If frm1.vspdData.SelModeSelected Then
				i_RowCnt    =   i_RowCnt  + 1
					For intColCnt = 0 To frm1.vspdData.MaxCols - 2
                        
                    	Select Case intColCnt
                    		Case 0
                                frm1.vspdData.Col = C_BP_NM               
                        	Case 1                                        
                                frm1.vspdData.Col = C_IO_TYPE_NM          
                    		Case 2                                        
                                frm1.vspdData.Col = C_IO_TYPE_CD          
                    		Case 3                                        
                                frm1.vspdData.Col = C_BP_CD               
                    		Case 4                                        
                                frm1.vspdData.Col = C_PO_NO               
                    		Case 5                                        
                                frm1.vspdData.Col = C_PO_SEQ_NO           
                    		Case 6                                        
                                frm1.vspdData.Col = C_PLANT_CD            
                    		Case 7                                        
                                frm1.vspdData.Col = C_SL_CD               
                    		Case 8                                        
                                frm1.vspdData.Col = C_ITEM_CD             
                    		Case 9                                        
                                frm1.vspdData.Col = C_ITEM_NM             
                    		Case 10                                       
                                frm1.vspdData.Col = C_SPEC                
                    		Case 11                                       
                                frm1.vspdData.Col = C_TRACKING_NO         
                    		Case 12                                       
                                frm1.vspdData.Col = C_PO_QTY              
                    		Case 13                                       
                                frm1.vspdData.Col = C_PO_UNIT             
                    		Case 14                                       
                                frm1.vspdData.Col = C_PO_PRC              
                    		Case 15                                       
                                frm1.vspdData.Col = C_PO_DOC_AMT          
                    		Case 16                                       
                                frm1.vspdData.Col = C_PO_CUR              
                    		Case 17                                       
                                frm1.vspdData.Col = C_DLVY_DT             
                    		Case 18                                       
                                frm1.vspdData.Col = C_RCPT_QTY            
                    		Case 19                                       
                                frm1.vspdData.Col = C_LC_QTY              
                    		Case 20                                       
                                frm1.vspdData.Col = C_PRE_IV_QTY          
                    		Case 21                                       
                                frm1.vspdData.Col = C_INSPECT_QTY         
                    		Case 22                                       
                                frm1.vspdData.Col = C_IV_QTY              
                    		Case 23                                       
                                frm1.vspdData.Col = C_RECV_INSPEC_FLG     
                    		Case 24                                       
                                frm1.vspdData.Col = C_MINOR_NM            
                    		Case 25                                       
                                frm1.vspdData.Col = C_INSPECT_METHOD      
                    		Case 26                                       
                                frm1.vspdData.Col = C_PLANT_NM            
                    		Case 27                                       
                                frm1.vspdData.Col = C_SL_NM               
                    		Case 28                                       
                                frm1.vspdData.Col = C_PUR_GRP             
                    		Case 29                                       
                                frm1.vspdData.Col = C_LC_RCPT_QTY         
                    		Case 30                                       
                                frm1.vspdData.Col = C_LOT_FLG             
                    		Case 31                                       
                                frm1.vspdData.Col = C_LOT_GEN_MTHD        
                    		Case 32
                                frm1.vspdData.Col = C_MVMT_NO
                    		Case 33
                                frm1.vspdData.Col = C_TRANS_TIME
                    		Case 34
                                frm1.vspdData.Col = C_MAIN_LOT
                    		Case 35
                                frm1.vspdData.Col = C_IMPORT_TIME
                    		Case 36
                                frm1.vspdData.Col = C_CREATE_TYPE
                    	End Select

                        if intColCnt = 32 then  '유무상구분 
                            curr_FLAG = frm1.vspdData.Text
                        end if
                        if intColCnt = 3 then  '공급처
                            curr_supplier = frm1.vspdData.Text
                        end if
                        if intColCnt = 2 then  '입고형태
                            curr_MvmtType = frm1.vspdData.Text
                        end if
                        
                        
                        if i_RowCnt <> 1 then
                            if curr_supplier <> before_supplier then
                                call DisplayMsgBox("ZZ0001", PopupParent.VB_INFORMATION, "X", "X")
                                Exit Function
                            end if
                           
                            if curr_MvmtType <> before_MvmtType then
                                call DisplayMsgBox("ZZ0002", PopupParent.VB_INFORMATION, "X", "X")
                               Exit Function
                            end if

                            if curr_FLAG <> before_FLAG then
                                call DisplayMsgBox("ZZ0005", PopupParent.VB_INFORMATION, "X", "X")
                                Exit Function
                            end if
                           
                        end if

                        
						arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
					Next
					intInsRow = intInsRow + 1
				End IF								
                before_supplier     =   curr_supplier
                before_MvmtType     =   curr_MvmtType
                before_FLAG         =   curr_FLAG
                
			Next
			
		End if			
		Self.Returnvalue = arrReturn
		Self.Close()
End Function	
'================================================================================================================================
Function CancelClick()
	Redim arrReturn(1,1)
	arrReturn(0,0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function
'================================================================================================================================
Function OpenPoNo()
	
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If gblnWinEvent = True Or UCase(frm1.txtPoNo.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function
		
	gblnWinEvent = True
	
	iCalledAspName = AskPRAspName("M3111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "M3111PA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If strRet(0) = "" Then
		frm1.txtPoNo.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
		frm1.txtPoNo.focus	
		Set gActiveElement = document.activeElement
	End If

End Function
'================================================================================================================================
Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Or UCase(frm1.txtGroupCd.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
'	arrParam(3) = Trim(frm1.txtGroupNm.Value)	
	
	arrParam(4) = ""			
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	gblnWinEvent = False
	
	If arrRet(0) = "" Then
		frm1.txtGroupCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtGroupCd.Value= arrRet(0)		
		frm1.txtGroupNm.Value= arrRet(1)	
		frm1.txtGroupCd.focus	
		Set gActiveElement = document.activeElement
	End If	

End Function 
'===============================  OpenTrackingNo()  ============================
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD
	
	If gblnWinEvent = True Then Exit Function
	
	gblnWinEvent = True

	arrParam(0) = ""	'주문처 
	arrParam(1) = ""	'영업그룹 
    arrParam(2) = ""	'공장 
    arrParam(3) = ""	'모품목 
    arrParam(4) = ""	'수주번호 
    arrParam(5) = ""	'추가 Where절 
    
	iCalledAspName = AskPRAspName("S3135PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "S3135PA1", "X")
		gblnWinEvent = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(PopupParent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
    
	gblnWinEvent = False

	If arrRet = "" Then
		frm1.txtTrackingNo.focus
		Exit Function
	Else
		frm1.txtTrackingNo.Value = Trim(arrRet)
		frm1.txtTrackingNo.focus
		lgBlnFlgChgValue = True
		Set gActiveElement = document.activeElement
	End If	

End Function
'================================================================================================================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next
	
	'If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'================================================================================================================================
'20071211::hanc
Function OpenMvmtType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Or UCase(frm1.txtMvmtType.className) = UCase(PopupParent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = "입고형태"	
	arrParam(1) = "( select distinct  IO_Type_Cd, io_type_nm from  M_CONFIG_PROCESS a,  m_mvmt_type b where a.rcpt_type = b.io_type_cd    and a.sto_flg = " & FilterVar("N", "''", "S") & "  AND a.USAGE_FLG=" & FilterVar("Y", "''", "S") & "  and ((b.RCPT_FLG=" & FilterVar("Y", "''", "S") & "  AND b.RET_FLG=" & FilterVar("N", "''", "S") & " ) or (b.RET_FLG=" & FilterVar("N", "''", "S") & "  And b.SUBCONTRA_FLG=" & FilterVar("N", "''", "S") & " )) ) c"
	arrParam(2) = Trim(frm1.txtMvmtType.Value)
	'arrParam(4) = "((RCPT_FLG='Y' AND RET_FLG='N') or (RET_FLG='N' And SUBCONTRA_FLG='N')) AND USAGE_FLG='Y' "
	arrParam(5) = "입고형태"			
	
    arrField(0) = "IO_Type_Cd"
    arrField(1) = "IO_Type_NM"
    
    arrHeader(0) = "입고형태"		
    arrHeader(1) = "입고형태명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
				
	IsOpenPop = False
	
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		frm1.txtMvmtType.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtMvmtType.Value	= arrRet(0)		
		frm1.txtMvmtTypeNm.Value= arrRet(1)
		Call changeMvmtType()
		lgBlnFlgChgValue = True
		frm1.txtMvmtType.focus	
		Set gActiveElement = document.activeElement
	End If	
End Function
'==============================================================================================================================
'20071211::hanc
Function changeMvmtType()
    changeMvmtType = False                 

	With frm1
		If 	CommonQueryRs(" A.IO_TYPE_NM, A.RCPT_FLG, A.IMPORT_FLG, A.RET_FLG, B.SUBCONTRA_FLG ", _
					" M_MVMT_TYPE A, M_CONFIG_PROCESS B ", _
					" A.IO_TYPE_CD = B.RCPT_TYPE AND B.STO_FLG = " & FilterVar("N", "''", "S") & "  AND B.USAGE_FLG= " & FilterVar("Y", "''", "S") & "  AND (A.RET_FLG = " & FilterVar("N", "''", "S") & "   AND (A.RCPT_FLG = " & FilterVar("Y", "''", "S") & "  OR A.SUBCONTRA_FLG = " & FilterVar("N", "''", "S") & " )) AND A.USAGE_FLG = " & FilterVar("Y", "''", "S") & "  AND A.IO_TYPE_CD = " & FilterVar(.txtMvmtType.Value, "''", "S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
			Call DisplayMsgBox("171900","X","X","X")
			.txtMvmtTypeNm.Value = ""
			Call SetFocusToDocument("M") 
			.txtMvmtType.focus
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		lgF1 = Split(lgF1, Chr(11))
		lgF2 = Split(lgF2, Chr(11))
		lgF3 = Split(lgF3, Chr(11))
		lgF4 = Split(lgF4, Chr(11))
		
		.txtMvmtTypeNm.Value	= lgF0(0)
		.hdnRcptflg.Value 		= lgF1(0)
		.hdnImportflg.Value		= lgF2(0)
		.hdnRetflg.Value 		= lgF3(0)
		.hdnSubcontraflg.Value  = lgF4(0)
		


	End With

	lgBlnFlgChgValue = true
    
    changeMvmtType = True                  

End Function
'==============================================================================================================================
'20071211::hanc
Function OpenSppl()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Or UCase(frm1.txtSupplierCd.className)=UCase(PopupParent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"				
	arrParam(1) = "B_Biz_Partner"
	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	arrParam(3) = ""							
	arrParam(4) = "Bp_Type in (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") AND usage_flag=" & FilterVar("Y", "''", "S") & "  AND  in_out_flag = " & FilterVar("O", "''", "S") & " "	
	arrParam(5) = "공급처"				
	
    arrField(0) = "BP_CD"					
    arrField(1) = "BP_NM"					

	arrHeader(0) = "공급처"				
	arrHeader(1) = "공급처명"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	
	IsOpenPop = False
	
	If isEmpty(arrRet) Then Exit Function				'페이지를 찾을 수 없는 에러발생시.
	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtSupplierCd.Value = arrRet(0)
		frm1.txtSupplierNm.Value = arrRet(1)
		frm1.txtSupplierCd.focus	
		Set gActiveElement = document.activeElement
	End If	
	
End Function
'==============================================================================================================================
'20071211::hanc
Function changeSpplCd()
	With frm1
		If 	CommonQueryRs(" BP_NM, BP_TYPE, usage_flag, in_out_flag "," B_Biz_Partner ", " BP_CD = " & FilterVar(.txtSuppliercd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
			Call DisplayMsgBox("229927","X","X","X")
			.txtSupplierNm.Value = ""
			Call SetFocusToDocument("M") 
			.txtSuppliercd.focus
			Exit function
		End If
		lgF0 = Split(lgF0, Chr(11))
		lgF1 = Split(lgF1, Chr(11))
		lgF2 = Split(lgF2, Chr(11))
		lgF3 = Split(lgF3, Chr(11))
		.txtSupplierNm.Value = lgF0(0)

		If Trim(lgF2(0)) <> "Y" Then
			Call DisplayMsgBox("179021","X","X","X")
			Call SetFocusToDocument("M") 
			.txtSuppliercd.focus
			Exit function
		End If
		If Trim(lgF1(0)) <> "S" and Trim(lgF1(0)) <> "CS" Then
			Call DisplayMsgBox("179020","X","X","X")
			Call SetFocusToDocument("M") 
			.txtSuppliercd.focus
			Exit function
		End If
		If Trim(lgF3(0)) <> "O" Then
			Call DisplayMsgBox("17C003","X","X","X")
			Call SetFocusToDocument("M") 
			.txtSuppliercd.focus
			Exit function
		End If
	End With        

End Function
'================================================================================================================================
Function OpenPlant()
	Dim arrRet,lgIsOpenPop
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
    If frm1.txtPlantCd.className = "protected" Then Exit Function
    
	lgIsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_PLANT"
	arrParam(2) = Trim(frm1.txtPlantCd.value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else	
		frm1.txtPlantCd.value = arrRet(0)
		frm1.txtPlantNm.value = arrRet(1)
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
	End If	
		
End Function	
'================================================================================================================================
Sub Form_Load()
	
	Call LoadInfTB19029															'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)	                                           
	Call ggoOper.LockField(Document, "N")										'⊙: Lock  Suitable  Field 
	Call InitVariables														    '⊙: Initializes local global variables
    Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")

	Call FncQuery()
End Sub
'================================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    
End Sub
'================================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
	     Exit Sub
	End If

	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub
'================================================================================================================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function
'================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub
'================================================================================================================================
Sub txtFrPoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'================================================================================================================================
Sub txtToPoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'================================================================================================================================
Sub txtFrPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrPoDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtFrPoDt.Focus
	End if
End Sub
'================================================================================================================================
Sub txtToPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToPoDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtToPoDt.Focus
	End if
End Sub
'================================================================================================================================
Function FncQuery() 
    
    FncQuery = False                                                 
    
    Err.Clear                                                        
	
	With frm1
		if (UniConvDateToYYYYMMDD(.txtFrPoDt.text,PopupParent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtToPoDt.text,PopupParent.gDateFormat,"")) And Trim(.txtFrPoDt.text) <> "" And Trim(.txtToPoDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","발주일", "X")	
			.txtToPoDt.Focus()
			Exit Function
		End if   
	End with
	
	ggoSpread.Source = frm1.vspdData	
    ggoSpread.ClearSpreadData
    
	Call InitVariables												
	
	If CheckRunningBizProcess = True Then Exit Function
    If DbQuery = False Then Exit Function

    FncQuery = True									
        
End Function
'================================================================================================================================
Function DbQuery()
	
	Dim strVal
	Dim strClsFlg
		
	Err.Clear															<%'☜: Protect system from crashing%>

	DbQuery = False														<%'⊙: Processing is NG%>

    If LayerShowHide(1) = False Then Exit Function
    
	If frm1.rdoClsFlg(0).checked Then
		strClsFlg = "Y"
	ElseIf frm1.rdoClsFlg(1).checked Then
		strClsFlg = "N"
	Else
		strClsFlg = ""
	End If

	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&txtPoNo=" &	Trim(frm1.hdnPoNo.Value)
		strVal = strVal & "&txtFrPoDt=" & Trim(frm1.hdnFrPoDt.value)
		strVal = strVal & "&txtToPoDt=" & Trim(frm1.hdnToPoDt.value)
		strVal = strVal & "&txtGroup=" & Trim(frm1.hdnGroupCd.value)
		strVal = strVal & "&txtPlant=" & Trim(frm1.hdnPlantCd.value)
		strVal = strVal & "&txtSupplier=" & Trim(frm1.hdnSupplierCd.value)
		strVal = strVal	& "&rdoClsFlg="	& Trim(strClsFlg)
        strVal = strVal     & "&lgStrPrevKey="  & lgStrPrevKey        '☜: Next key tag
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001
		strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.value)
		strVal = strVal & "&txtFrPoDt=" & Trim(frm1.txtFrPoDt.text)
		strVal = strVal & "&txtToPoDt=" & Trim(frm1.txtToPoDt.text)
		strVal = strVal & "&txtGroup=" & Trim(frm1.txtGroupCd.value)
		strVal = strVal & "&txtPlant=" & Trim(frm1.txtPlantCd.value)
		strVal = strVal & "&txtSupplier=" & Trim(frm1.txtSupplierCd.value)
		strVal = strVal	& "&rdoClsFlg="	& Trim(strClsFlg)
        strVal = strVal     & "&lgStrPrevKey="  & lgStrPrevKey        '☜: Next key tag
	End if 

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  

	Call RunMyBizASP(MyBizASP, strVal)								<%'☜: 비지니스 ASP 를 가동 %>

	DbQuery = True														<%'⊙: Processing is NG%>
End Function
'================================================================================================================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>

	lgIntFlgMode = PopupParent.OPMD_UMODE

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True
	Else
		frm1.txtPoNo.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS="TD5" NOWRAP>발주번호</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo" SIZE=32 MAXLENGTH=18 ALT="발주번호" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"><div style="Display:none"><input type="text" name=none></div></TD>
						<TD CLASS="TD5" NOWRAP>입고기간</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td NOWRAP>
										<script language =javascript src='./js/m3112ra5_fpDateTime1_txtFrPoDt.js'></script>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<script language =javascript src='./js/m3112ra5_fpDateTime1_txtToPoDt.js'></script>
									</td>
								<tr>
							</table>
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>유무상구분</TD> 
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio AlT="유상"   NAME="rdoClsFlg" ID="rdoClsFlg0" CLASS="RADIO" value = "Y" tag="11" checked><label for="rdoClsFlg0">&nbsp;유상&nbsp;&nbsp;</label>
											   <INPUT TYPE=radio AlT="무상"   NAME="rdoClsFlg" ID="rdoClsFlg1" CLASS="RADIO" value = "N" tag="11"><label for="rdoClsFlg1">&nbsp;무상&nbsp;</label>
											   <INPUT TYPE=radio AlT="전체"   NAME="rdoClsFlg" ID="rdoClsFlg2" CLASS="RADIO" value = "A" tag="11"><label for="rdoClsFlg2">&nbsp;전체&nbsp;&nbsp;</label></TD>
						<TD CLASS="TD5" NOWRAP>공장</TD>
						<TD CLASS="TD6" NOWRAP>
						<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="공장" tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">
						<INPUT TYPE=TEXT AlT="공장" ID="txtPlantNm" tag="14X">
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP></TD>
						<TD CLASS="TD6" NOWRAP></TD>
						<TD CLASS="TD5" NOWRAP>공급처</TD>
						<TD CLASS="TD6" NOWRAP>
						<INPUT STYLE = "text-transform:uppercase" TYPE=TEXT ALT="공급처" NAME="txtSupplierCd" MAXLENGTH=10 SIZE=10 ALT="공급처" tag="11XXXU" OnChange="VBScript:changeSpplCd()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
						<INPUT TYPE=TEXT ALT="공급처" NAME="txtSupplierNm" SIZE=20 tag="24X">
						</TD>
					</TR>
					<TR>
						<TD style="display:none" CLASS=TD5 NOWRAP>구매그룹</TD> 
						<TD style="display:none" CLASS=TD6 NOWRAP>
						<INPUT TYPE=TEXT AlT="구매그룹" NAME="txtGroupCd" SIZE=10 MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()">
						<INPUT TYPE=TEXT AlT="구매그룹" ID="txtGroupNm" NAME="arrCond" tag="14X"></TD>
						</TD>
						<TD style="display:none" CLASS="TD5" NOWRAP>입고형태</TD>
						<TD style="display:none" CLASS="TD6" NOWRAP>
						<INPUT STYLE = "text-transform:uppercase" TYPE=TEXT Alt="입고형태" NAME="txtMvmtType" SIZE=10 MAXLENGTH=5 tag="1XNXXU" OnChange="VBScript:changeMvmtType()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMoveType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenMvmtType()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
						<INPUT TYPE=TEXT Alt="입고형태" NAME="txtMvmtTypeNm" SIZE=20 tag="24X">
						</TD>
					</TR>
					<TR style="display:none">
						<TD style="display:none" CLASS="TD5" NOWRAP>Tracking No.</TD>
						<TD style="display:none" CLASS="TD6" NOWRAP><INPUT NAME="txtTrackingNo" ALT="Tracking번호" TYPE="Text" MAXLENGTH=25 SiZE=26  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo">
						</TD>
						<TD CLASS="TD5" NOWRAP></TD>
						<TD CLASS="TD6" NOWRAP>
						</TD>
					</TR>				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
						<script language =javascript src='./js/m3112ra5_vspdData_vspdData.js'></script>
					</TD>
				</TR>			
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
						                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
                                         <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
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



<INPUT TYPE=HIDDEN NAME="hdnPoNo" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnFrPoDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToPoDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSupplierCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnGroupNm" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnClsflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnReleaseflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnImportflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSubcontraflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRefType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnRcptType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnIvFlg" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPoType" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="14">

</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                

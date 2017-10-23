'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************

'Grid 1 - Order Header
Const BIZ_PGM_QRY1_ID	= "p4120mb1_ko119.asp"						'☆: Head Query 비지니스 로직 ASP명 
'Post Production Results
Const BIZ_PGM_SAVE_ID	= "p4120mb2_ko119.asp"						'☆: 비지니스 로직 ASP명 							
'Jump (E)Production Order 
Const BIZ_PGM_JUMPREWORKRUN_ID = "p4111ma1"
'Jump (E)Resource Consumption (By Order)
Const BIZ_PGM_JUMPORDRSCCOMPT_ID = "p4712ma1"

Dim C_GRIDCOUNT
Dim MaxCount
Dim lgShiftCnt		

'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================

'' Grid 1(vspdData1) - Order Header
'Dim C_ProdtOrderNo			
'Dim C_ItemCd				
'Dim C_ItemNm				
'Dim C_Spec
'Dim C_SecItemCd                 '삼성코드				
'Dim C_RoutNo				
'Dim C_ProdtOrderQty					
'Dim C_ProdtOrderSumQty				
'Dim C_BaseUnit		
'Dim C_TrackingNo							 
'Dim C_LampMaker				    'Lamp Maker
'Dim C_InvertMaker				'인버터 Maker


' Grid 1(vspdData1) - Results
' Hidden
Const C_JobLineCd		=	1
Const C_JobLine			=	2
Const C_ItemCd			=	3
Const C_ItemNm          =   4
Const C_Spec            =   5
Const C_ProdtOrderNo	=   6
Const C_ProdtOrderPopup =   7
Const C_ProdtOrderQty   =   8		
Const C_ProdtOrderLineSumQty = 9
Const C_JobPlanTime1	=   10
Const C_JobOrderNo1		=   11     
Const C_JobQty1			=	12
Const C_JobSeq1			=	13
Const C_CheckBox1       =   14
Const C_JobPlanTime2	=   15
Const C_JobOrderNo2     =   16
Const C_JobQty2         =	17 
Const C_JobSeq2			=	18
Const C_CheckBox2       =   19
Const C_JobPlanTime3	=   20
Const C_JobOrderNo3     =   21
Const C_JobQty3         =	22
Const C_JobSeq3			=	23
Const C_CheckBox3       =   24
Const C_JobPlanTime4	=   25
Const C_JobOrderNo4     =   26 
Const C_JobQty4         =	27 
Const C_JobSeq4			=	28
Const C_CheckBox4       =   29
Const C_JobPlanTime5	=   30
Const C_JobOrderNo5		=	31 
Const C_JobQty5			=	32
Const C_JobSeq5			=	33
Const C_CheckBox5       =   34
Const C_JobPlanTime6	=   35
Const C_JobOrderNo6		=	36
Const C_JobQty6			=	37
Const C_JobSeq6			=	38
Const C_CheckBox6       =   39
Const C_JobPlanTime7	=   40
Const C_JobOrderNo7		=	41
Const C_JobQty7			=	42
Const C_JobSeq7			=	43
Const C_CheckBox7       =   44
Const C_JobPlanTime8	=   45
Const C_JobOrderNo8		=	46
Const C_JobQty8			=	47
Const C_JobSeq8			=	48
Const C_CheckBox8       =   49
Const C_JobPlanTime9	=   50
Const C_JobOrderNo9		=	51
Const C_JobQty9			=	52
Const C_JobSeq9			=	53
Const C_CheckBox9       =   54
Const C_JobPlanTime10	=   55
Const C_JobOrderNo10	=	56	
Const C_JobQty10		=	57
Const C_JobSeq10		=	58
Const C_CheckBox10      =   59
Const C_JobPlanTime11	=   60
Const C_JobOrderNo11	=	61
Const C_JobQty11		=	62
Const C_JobSeq11		=	63
Const C_CheckBox11      =   64
Const C_JobPlanTime12	=   65
Const C_JobOrderNo12	=	66
Const C_JobQty12		=	67
Const C_JobSeq12		=	68
Const C_CheckBox12      =   69
Const C_JobPlanTime13	=   70
Const C_JobOrderNo13	=	71
Const C_JobQty13		=	72
Const C_JobSeq13		=	73
Const C_CheckBox13      =   74
Const C_JobPlanTime14	=   75
Const C_JobOrderNo14	=	76
Const C_JobQty14		=	77
Const C_JobSeq14		=	78
Const C_CheckBox14      =   79
Const C_JobPlanTime15	=   80
Const C_JobOrderNo15	=	81
Const C_JobQty15		=	82
Const C_JobSeq15		=	83
Const C_CheckBox15      =   84
Const C_JobPlanTime16	=   85
Const C_JobOrderNo16	=	86
Const C_JobQty16		=	87
Const C_JobSeq16		=	88
Const C_CheckBox16      =   89
Const C_JobPlanTime17	=   90
Const C_JobOrderNo17	=	91
Const C_JobQty17		=	92
Const C_JobSeq17		=	93
Const C_CheckBox17      =   94
Const C_JobPlanTime18	=   95
Const C_JobOrderNo18	=	96
Const C_JobQty18		=	97
Const C_JobSeq18		=	98
Const C_CheckBox18      =   99
Const C_JobPlanTime19	=   100
Const C_JobOrderNo19	=	101
Const C_JobQty19		=	102
Const C_JobSeq19		=	103
Const C_CheckBox19      =   104
Const C_JobPlanTime20	=   105
Const C_JobOrderNo20	=	106
Const C_JobQty20		=	107
Const C_JobSeq20		=	108
Const C_CheckBox20      =   109
Const C_JobPlanTime21	=   110
Const C_JobOrderNo21	=	111
Const C_JobQty21		=	112
Const C_JobSeq21		=	113
Const C_CheckBox21      =   114
Const C_JobPlanTime22	=   115
Const C_JobOrderNo22	=	116
Const C_JobQty22		=	117
Const C_JobSeq22		=	118
Const C_CheckBox22      =   119
Const C_JobPlanTime23	=   120
Const C_JobOrderNo23	=	121
Const C_JobQty23		=	122
Const C_JobSeq23		=	123
Const C_CheckBox23      =   124
Const C_JobPlanTime24	=   125
Const C_JobOrderNo24	=	126
Const C_JobQty24		=	127
Const C_JobSeq24		=	128
Const C_CheckBox24      =   129

Dim C_JobQtyNm1
Dim C_JobQtyNm2
Dim C_JobQtyNm3
Dim C_JobQtyNm4
Dim C_JobQtyNm5
Dim C_JobQtyNm6
Dim C_JobQtyNm7
Dim C_JobQtyNm8
Dim C_JobQtyNm9
Dim C_JobQtyNm10
Dim C_JobQtyNm11
Dim C_JobQtyNm12
Dim C_JobQtyNm13
Dim C_JobQtyNm14
Dim C_JobQtyNm15
Dim C_JobQtyNm16
Dim C_JobQtyNm17
Dim C_JobQtyNm18
Dim C_JobQtyNm19
Dim C_JobQtyNm20
Dim C_JobQtyNm21
Dim C_JobQtyNm22
Dim C_JobQtyNm23
Dim C_JobQtyNm24

		

' Grid 3(vspdData3) - Hidden
Const C_HJobPlanTime1	=   1
Const C_HJobPlanTime2	=   2
Const C_HJobPlanTime3	=   3
Const C_HJobPlanTime4	=   4
Const C_HJobPlanTime5	=   5
Const C_HJobPlanTime6	=   6
Const C_HJobPlanTime7	=   7
Const C_HJobPlanTime8	=   8
Const C_HJobPlanTime9	=   9
Const C_HJobPlanTime10	=   10
Const C_HJobPlanTime11	=   11
Const C_HJobPlanTime12	=   12
Const C_HJobPlanTime13	=   13
Const C_HJobPlanTime14	=   14
Const C_HJobPlanTime15	=   15
Const C_HJobPlanTime16	=   16
Const C_HJobPlanTime17	=   17
Const C_HJobPlanTime18	=   18
Const C_HJobPlanTime19	=   19
Const C_HJobPlanTime20	=   20
Const C_HJobPlanTime21	=   21
Const C_HJobPlanTime22	=   22
Const C_HJobPlanTime23	=   23
Const C_HJobPlanTime24	=   24

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================
Dim lgIntFlgMode								'Variable is for Operation Status
Dim lgIntPrevKey
Dim lgStrPrevKey
Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim lgLngCurRows
Dim lgCurrRow
Dim lgShift
'==========================================  1.2.3 Global Variable값 정의  ==================================
'============================================================================================================
'----------------------  공통 Global 변수값 정의  -----------------------------------------------------------
Dim IsOpenPop 
Dim lgOldRow    
Dim lgSortKey1   
Dim lgSortKey2 
Dim GridColCount
Dim lgKeyStream2
'++++++++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++

'#########################################################################################################
'												2. Function부 
'
'	내용 : 개발자가 정의한 함수, 즉 Event관련 함수를 제외한 모든 사용자 정의 함수 기슬 
'	공통으로 적용 사항 : 1. Sub 또는 Function을 호출할 때 반드시 Call을 쓴다.
'		     	     	 2. Sub, Function 이름에 _를 쓰지 않도록 한다. (Event와 구별하기 위함) 
'#########################################################################################################
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""							'initializes Previous Key
    lgStrPrevKey1 = ""							'initializes Previous Key
    lgStrPrevKey2 = ""							'initializes Previous Key
    lgIntPrevKey = 0
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgOldRow = 0
    lgSortKey1 = 1
    lgSortKey2 = 1
    lgKeyStream2 = ""
    lgShiftCnt = 0    
End Sub

'******************************************  2.2 화면 초기화 함수  ***************************************
'	기능: 화면초기화 
'	설명: 화면초기화, Combo Display, 화면 Clear 등 화면 초기화 작업을 한다. 
'*********************************************************************************************************
'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim strWhere
	
	if frm1.txtPlantCd.value = "" then
	strWhere    = ""
	strWhere	= strWhere &  " plant_cd >= '" & frm1.txtPlantCd.value & "'"
	else
	strWhere    = ""
	strWhere	= strWhere &  " plant_cd = '" & frm1.txtPlantCd.value & "'"
	end if
	strWhere	= strWhere & " order by line_group, work_line "

	Call CommonQueryRs(" WORK_LINE, WORK_LINE_DESC ", " p_work_line_ko119 ", strWhere , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboLine, lgF0, lgF1, Chr(11)) 

'    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_JOBLINECD			
'    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_JOBLINE
End Sub

Sub InitComboBox2()
	Dim i
	Dim strWhere
	Dim strSelect
	Dim strFrom
	Dim arrVal1, arrVal2
	Dim ii, jj
	Dim strPlantCd
	
	strSelect	=			 " plant_cd "
	strFrom		=			 " b_plant (NOLOCK) "
	strWhere	=			 " plant_cd > '' "

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 

		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1) 
       For ii = 0 To jj - 1 
			arrVal2			= Split(arrVal1(ii), chr(11))
			strPlantCd		= Ucase(Trim(arrVal2(1)))
'			strMinorNm		= Trim(arrVal2(2))

	
		strWhere    = ""
		strWhere	=	strWhere &  " plant_cd = '" & strPlantCd & "'"
   	   	If  CommonQueryRs(" count(WORK_LINE) "," p_work_line_ko119 (NOLOCK) ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) then
	      	If UNICdbl(Replace(lgF0,Chr(11),"")) = 0 Then
	      	   if lgShiftCnt > 0 then
	      	   else
					lgShiftCnt = 0
			   end if	
			Else
			   if lgShiftCnt = 0 then
			     if int(Replace(lgF0,Chr(11),"")) > int(lgShiftCnt) then
			     	lgShiftCnt = Replace(lgF0,Chr(11),"")
			     end if	
			   end if  
			End if	
		End if
	  Next    
   end if	       

'	If Trim(frm1.txtPlantCd.value) = "" Then
'		frm1.txtPlantNm.value = ""
'	Else
		For i = lgShiftCnt + 1 To 1 Step -1
			frm1.cboLine.remove(i) 
		Next
'	End If
End Sub

Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
Dim strSelect, strFrom, strWhere
Dim arrVal1, arrVal2
Dim strMinorCd, strMinorNm
Dim HJobPlanTime1, HJobPlanTime2, HJobPlanTime3, HJobPlanTime4, HJobPlanTime5, HJobPlanTime6, HJobPlanTime7
Dim HJobPlanTime8, HJobPlanTime9, HJobPlanTime10, HJobPlanTime11, HJobPlanTime12, HJobPlanTime13, HJobPlanTime14
Dim HJobPlanTime15, HJobPlanTime16, HJobPlanTime17, HJobPlanTime18, HJobPlanTime19, HJobPlanTime20, HJobPlanTime21
Dim HJobPlanTime22, HJobPlanTime23, HJobPlanTime24

   with frm1
       .vspdData3.Col = C_HJobPlanTime1
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime2
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime3
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime4
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime5
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime6
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime7
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime8
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime9
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime10
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime11
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime12
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime13
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime14
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime15
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime16
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime17
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime18
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime19
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime20
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime21
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime22
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime23
       .vspdData3.value = ""
       .vspdData3.Col = C_HJobPlanTime24
       .vspdData3.value = ""
    end with
 
 HJobPlanTime1  = ""  : HJobPlanTime2  = ""  : HJobPlanTime3  = ""  : HJobPlanTime4  = ""  : HJobPlanTime5  = ""
 HJobPlanTime6  = ""  : HJobPlanTime7  = ""  : HJobPlanTime8  = ""  : HJobPlanTime9  = ""  : HJobPlanTime10 = "" 
 HJobPlanTime11 = ""  : HJobPlanTime12 = ""  : HJobPlanTime13 = ""  : HJobPlanTime14 = ""  : HJobPlanTime15 = "" 
 HJobPlanTime16 = ""  : HJobPlanTime17 = ""  : HJobPlanTime18 = ""  : HJobPlanTime19 = ""  : HJobPlanTime20 = "" 
 HJobPlanTime21 = ""  : HJobPlanTime22 = ""  : HJobPlanTime23 = ""  : HJobPlanTime24 = ""  

 With frm1	
   If C_GRIDCOUNT > 0 then

	strSelect	=			 " a.minor_cd, a.minor_nm "
	strFrom		=			 " b_minor a (NOLOCK), b_configuration b (nolock) "
	strWhere	=			 " a.major_cd = b.major_cd and a.major_cd = 'M2110' and b.seq_no = 99 and a.minor_cd = b.minor_cd "
	strWhere	= strWhere & " order by b.reference "

	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 

		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1) 

		For ii = 0 To jj - 1 
			i = ii + 1
			arrVal2			= Split(arrVal1(ii), chr(11))
			strMinorCd		= Ucase(Trim(arrVal2(1)))
			strMinorNm		= Trim(arrVal2(2))
			
			select case i 
				case "1"
					.vspdData3.Col = C_HJobPlanTime1
					.vspdData3.value = strMinorCd
					 HJobPlanTime1 = Trim(.vspdData3.value) 
				case "2"
					.vspdData3.Col = C_HJobPlanTime2
					.vspdData3.value = strMinorCd
					HJobPlanTime2 = Trim(.vspdData3.value) 
				case "3"
					.vspdData3.Col = C_HJobPlanTime3
					.vspdData3.value = strMinorCd
					HJobPlanTime3 = Trim(.vspdData3.value) 
				case "4"
					.vspdData3.Col = C_HJobPlanTim4
					.vspdData3.value = strMinorCd
					HJobPlanTime4 = Trim(.vspdData3.value) 
				case "5"
					.vspdData3.Col = C_HJobPlanTime5
					.vspdData3.value = strMinorCd
					HJobPlanTime5 = Trim(.vspdData3.value) 
				case "6"
					.vspdData3.Col = C_HJobPlanTime6
					.vspdData3.value = strMinorCd
					HJobPlanTime6 = Trim(.vspdData3.value) 
				case "7"
					.vspdData3.Col = C_HJobPlanTime7
					.vspdData3.value = strMinorCd
					HJobPlanTime7 = Trim(.vspdData3.value) 
				case "8"
					.vspdData3.Col = C_HJobPlanTime8
					.vspdData3.value = strMinorCd
					HJobPlanTime8 = Trim(.vspdData3.value) 
				case "9"
					.vspdData3.Col = C_HJobPlanTime9
					.vspdData3.value = strMinorCd
					HJobPlanTime9 = Trim(.vspdData3.value) 
				case "10"
					.vspdData3.Col = C_HJobPlanTime10
					.vspdData3.value = strMinorCd
					HJobPlanTime10 = Trim(.vspdData3.value) 
				case "11"
					.vspdData3.Col = C_HJobPlanTime11
					.vspdData3.value = strMinorCd
					HJobPlanTime11 = Trim(.vspdData3.value) 
				case "12"
					.vspdData3.Col = C_HJobPlanTime12
					.vspdData3.value = strMinorCd
					HJobPlanTime12 = Trim(.vspdData3.value) 
				case "13"
					.vspdData3.Col = C_HJobPlanTime13
					.vspdData3.value = strMinorCd
					HJobPlanTime13 = Trim(.vspdData3.value) 
				case "14"
					.vspdData3.Col = C_HJobPlanTime14
					.vspdData3.value = strMinorCd
					HJobPlanTime14 = Trim(.vspdData3.value) 
				case "15"
					.vspdData3.Col = C_HJobPlanTime15
					.vspdData3.value = strMinorCd
					HJobPlanTime15 = Trim(.vspdData3.value) 
				case "16"
					.vspdData3.Col = C_HJobPlanTime16
					.vspdData3.value = strMinorCd
					HJobPlanTime16 = Trim(.vspdData3.value) 
				case "17"
					.vspdData3.Col = C_HJobPlanTime17
					.vspdData3.value = strMinorCd
					HJobPlanTime17 = Trim(.vspdData3.value) 
				case "18"
					.vspdData3.Col = C_HJobPlanTime18
					.vspdData3.value = strMinorCd
					HJobPlanTime18 = Trim(.vspdData3.value) 
				case "19"
					.vspdData3.Col = C_HJobPlanTime19
					.vspdData3.value = strMinorCd
					HJobPlanTime19 = Trim(.vspdData3.value) 
				case "20"
					.vspdData3.Col = C_HJobPlanTime20
					.vspdData3.value = strMinorCd
					HJobPlanTime20 = Trim(.vspdData3.value) 
				case "21"
					.vspdData3.Col = C_HJobPlanTime21
					.vspdData3.value = strMinorCd
					HJobPlanTime21 = Trim(.vspdData3.value) 
				case "22"
					.vspdData3.Col = C_HJobPlanTime22
					.vspdData3.value = strMinorCd
					HJobPlanTime22 = Trim(.vspdData3.value) 
				case "23"
					.vspdData3.Col = C_HJobPlanTime23
					.vspdData3.value = strMinorCd
					HJobPlanTime23 = Trim(.vspdData3.value) 
				case "24"
					.vspdData3.Col = C_HJobPlanTime24
					.vspdData3.value = strMinorCd
					HJobPlanTime24 = Trim(.vspdData3.value) 
			End Select		
			Next
		End if
	End if	 
End With 
    

	lgKeyStream2 = HJobPlanTime1  & Parent.gColSep       'You Must append one character(gColSep)
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime2 & Parent.gColSep       'You Must append one character(gColSep)
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime3 & Parent.gColSep       'You Must append one character(gColSep)
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime4 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime5 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime6 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime7 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime8 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime9 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime10 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime11 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime12 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime13 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime14 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime15 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime16 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime17 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime18 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime19 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime20 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime21 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime22 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime23 & Parent.gColSep
	lgKeyStream2 = lgKeyStream2 & HJobPlanTime24 & Parent.gColSep

   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        

'==========================================  2.2.1 SetDefaultVal()  ======================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
    frm1.txtProdFromDt.text = StartDate
'    frm1.txtProdToDt.text   = EndDate
End Sub

Sub InitSpreadSheet2()
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim IntRetCD
	
   	strWhere = " MAJOR_CD = 'M2110' "
	
   	If  CommonQueryRs(" count(minor_cd) "," B_MINOR (NOLOCK) ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) then
      	If UNICdbl(Replace(lgF0,Chr(11),"")) = 0 Then
			C_GRIDCOUNT = 0
		Else
			C_GRIDCOUNT = Replace(lgF0,Chr(11),"")
		End if	
	else
	End if	

End Sub

Sub InitSpreadSheet3(plantcd)
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim IntRetCD
	Dim arrVal1, arrVal2
	Dim ii, jj
	Dim	StartDate2
	
	startDate2 = frm1.txtProdFromDt.text
	
	strSelect = " distinct job_line, item_cd, prodt_order_no "
	strFrom  = " p_prod_time_order_ko119 (NOLOCK) "
   	strWhere = " plant_cd = " & FilterVar(plantcd,"''","S")
   	strWhere = strWhere & " and job_plan_dt = '" & UNIConvDate(Trim(StartDate2)) & "'"   	   	
   	
   	if frm1.cboLine.value <> "" then
   		strWhere = strWhere & " and job_line = " & FilterVar(frm1.cboLine.value,"''","S")
   	end if 
   	
   	if frm1.txtProdOrderNo.value <> "" then
   		strWhere = strWhere & " and prodt_order_no = " & FilterVar(frm1.txtProdOrderNo.value,"''","S")
   	end if
   	
   	if frm1.txtItemCd.value <> "" then
   		strWhere = strWhere & " and item_cd = " & FilterVar(frm1.txtItemCd.value,"''","S")
   	end if 
   	
   	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 

		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
		jj = Ubound(arrVal1,1) 
		MaxCount = jj 
	else
	    MaxCount = 0
   	End if
	
'   	If  CommonQueryRs(" job_line "," p_prod_time_order_ko119 (NOLOCK) ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) then
'      	If UNICdbl(Replace(lgF0,Chr(11),"")) = 0 Then
'			MaxCount = 0
'		Else
'			MaxCount = Replace(lgF0,Chr(11),"")
'		End if	
'	else
'	End if	

End Sub

'============================= 2.2.3 InitSpreadSheet() ==================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim ii, jj, i
	Dim arrVal1, arrVal2
	Dim strMinorCd, strMinorNm
	Dim iDx, iDx2
	Dim strJobCaption	

'	Call InitSpreadPosVariables(pvSpdNo)

	Call AppendNumberPlace("6", "18", "0")
	Call AppendNumberPlace("7", "5", "0")
	
	If pvSpdNo = "A" Or pvSpdNo = "*" Then

		'------------------------------------------
		' Grid 2 - Component Spread Setting
		'------------------------------------------
	    Call InitSpreadSheet2()

      If C_GRIDCOUNT > 0 then

		  	strSelect	=			 " a.minor_cd, a.minor_nm "
			strFrom		=			 " b_minor a (NOLOCK), b_configuration b (nolock) "
			strWhere	=			 " a.major_cd = b.major_cd and a.major_cd = 'M2110' and b.seq_no = 99 and a.minor_cd = b.minor_cd "
			strWhere	= strWhere & " order by b.reference "

			If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 

				arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
				jj = Ubound(arrVal1,1) 

				For ii = 0 To jj - 1 
					i = ii + 1
					arrVal2			= Split(arrVal1(ii), chr(11))
					strMinorCd		= Ucase(Trim(arrVal2(1)))
					strMinorNm		= Trim(arrVal2(2))

				select case i 
					case "1"
						C_JobQtyNm1 = strMinorNm
					case "2"
						C_JobQtyNm2 = strMinorNm
					case "3"
						C_JobQtyNm3 = strMinorNm
					case "4"
						C_JobQtyNm4 = strMinorNm
					case "5"
						C_JobQtyNm5 = strMinorNm
					case "6"
						C_JobQtyNm6 = strMinorNm
					case "7"
						C_JobQtyNm7 = strMinorNm
					case "8"
						C_JobQtyNm8 = strMinorNm
					case "9"
						C_JobQtyNm9 = strMinorNm
					case "10"
						C_JobQtyNm10 = strMinorNm
					case "11"
						C_JobQtyNm11 = strMinorNm
					case "12"
						C_JobQtyNm12 = strMinorNm
					case "13"
						C_JobQtyNm13 = strMinorNm
					case "14"
						C_JobQtyNm14 = strMinorNm
					case "15"
						C_JobQtyNm15 = strMinorNm
					case "16"
						C_JobQtyNm16 = strMinorNm
					case "17"
						C_JobQtyNm17 = strMinorNm
					case "18"
						C_JobQtyNm18 = strMinorNm
					case "19"
						C_JobQtyNm19 = strMinorNm
					case "20"
						C_JobQtyNm20 = strMinorNm
					case "21"
						C_JobQtyNm21 = strMinorNm
					case "22"
						C_JobQtyNm22 = strMinorNm
					case "23"
						C_JobQtyNm23 = strMinorNm
					case "24"
						C_JobQtyNm24 = strMinorNm
				End Select		
				Next
			End If 
      End if
		
		With frm1.vspdData1
			
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20030808", ,Parent.gAllowDragDropSpread
    
			.ReDraw = false
    
          
          Select Case C_GRIDCOUNT
			Case "1"
				.MaxCols = C_CheckBox1 +1										'☜: 최대 Columns의 항상 1개 증가시킴 
			Case "2"
				.MaxCols = C_CheckBox2 +1
			Case "3"
				.MaxCols = C_CheckBox3 +1
			Case "4"
				.MaxCols = C_CheckBox4 +1	
			Case "5"
				.MaxCols = C_CheckBox5 +1	
			Case "6"
				.MaxCols = C_CheckBox6 +1
			Case "7"
				.MaxCols = C_CheckBox7 +1		
			Case "8"
				.MaxCols = C_CheckBox8 +1		
			Case "9"
				.MaxCols = C_CheckBox9 +1		
			Case "10"
				.MaxCols = C_CheckBox10 +1		
			Case "11"
				.MaxCols = C_CheckBox11 +1		
			Case "12"
				.MaxCols = C_CheckBox12 +1		
			Case "13"
				.MaxCols = C_CheckBox13 +1		
			Case "14"
				.MaxCols = C_CheckBox14 +1		
			Case "15"
				.MaxCols = C_CheckBox15 +1		
			Case "16"
				.MaxCols = C_CheckBox16 +1		
			Case "17"
				.MaxCols = C_CheckBox17 +1		
			Case "18"
				.MaxCols = C_CheckBox18 +1		
			Case "19"
				.MaxCols = C_CheckBox19 +1		
			Case "20"
				.MaxCols = C_CheckBox20 +1		
			Case "21"
				.MaxCols = C_CheckBox21 +1		
			Case "22"
				.MaxCols = C_CheckBox22 +1
			Case "23"
				.MaxCols = C_CheckBox23 +1		
			Case "24"
				.MaxCols = C_CheckBox24 +1						
		  End Select				
				
			.MaxRows = 0
			
'			Call GetSpreadColumnPos("B") 

			ggoSpread.SSSetCombo	C_JOBLINECD, "LineCd", 8
			ggoSpread.SSSetCombo	C_JOBLINE, "Line", 8
			ggoSpread.SSSetEdit		C_ItemCd, "품목", 12,,,18
			ggoSpread.SSSetEdit		C_ItemNm, "품목명", 25
			ggoSpread.SSSetEdit		C_Spec, "규격", 25,,,25				
			ggoSpread.SSSetEdit		C_ProdtOrderNo, "제조오더번호", 12,,,18
			ggoSpread.SSSetButton	C_ProdtOrderPopup
			ggoSpread.SSSetFloat	C_ProdtOrderQty,"오더수량",10, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_ProdtOrderLineSumQty,"Line합",10, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"

			For ii = 1  to C_GRIDCOUNT * 5
				Select Case ii 
				  Case "1" 
   					iDx = "10"                       '작업계획시간
				  Case "2"	                        
   					iDx = "11"                       '작업지시번호
				  Case "3"
   					iDx = "12"                       '작업지시수량  
   					strJobCaption = C_JobQtyNm1   					
   				  Case "4"	
   					iDx = "13"						 '작업지시순서 
				  Case "5"
					iDx = "14"						 '작업지시확정
					
				  Case "6"
					iDx = "15"
				  Case "7"
				    iDx = "16"
				  Case "8"
					iDx = "17"
					strJobCaption = C_JobQtyNm2
				  Case "9"
					iDx = "18"
				  Case "10"
					iDx = "19"
					
				  Case "11"
					iDx = "20"					
				  Case "12"
					iDx = "21"
				  Case "13"
					iDx = "22"
					strJobCaption = C_JobQtyNm3				    
				  Case "14"
					iDx = "23"
				  Case "15"
					iDx = "24"
										
				  Case "16"
					iDx = "25"
				  Case "17"
					iDx = "26"
				  Case "18"
					iDx = "27"
					strJobCaption = C_JobQtyNm4
				  Case "19"
					iDx = "28"					
				  Case "20"
					iDx = "29"
										
				  Case "21"
					iDx = "30"
				  Case "22"
					iDx = "31"
				  Case "23"
					iDx = "32"
					strJobCaption = C_JobQtyNm5										
				  Case "24"
					iDx = "33"
				  Case "25"
					iDx = "34"
										
  				  Case "26"
					iDx = "35"
   				  Case "27"
					iDx = "36"					
				  Case "28"
					iDx = "37"
					strJobCaption = C_JobQtyNm6
				  Case "29"
					iDx = "38"
				  Case "30"
					iDx = "39"
					
				  Case "31"
					iDx = "40"					
				  Case "32"
					iDx = "41"
				  Case "33"
					iDx = "42"
					strJobCaption = C_JobQtyNm7
				  Case "34"
					iDx = "43"
				  Case "35"
					iDx = "44"
															
				  Case "36"
					iDx = "45"
				  Case "37"
					iDx = "46"
				  Case "38"
					iDx = "47"
					strJobCaption = C_JobQtyNm8
				  Case "39"
					iDx = "48"					
				  Case "40"
					iDx = "49"
					
				  Case "41"
					iDx = "50"
				  Case "42"
					iDx = "51"
				  Case "43"
					iDx = "52"
					strJobCaption = C_JobQtyNm9					
				  Case "44"
					iDx = "53"	
				  Case "45"
					iDx = "54"
							
				  Case "46"
					iDx = "55"	
				  Case "47"
					iDx = "56"					 	
				  Case "48"
					iDx = "57"	
					strJobCaption = C_JobQtyNm10
				  Case "49"
					iDx = "58"	
				  Case "50"
					iDx = "59"
					
				  Case "51"
					iDx = "60"							
				  Case "52"
					iDx = "61"	
				  Case "53"
					iDx = "62"
					strJobCaption = C_JobQtyNm11
				  Case "54"
					iDx = "63"		
				  Case "55"
					iDx = "64"
										
				  Case "56"
					iDx = "65"
				  Case "57"
					iDx = "66"		
				  Case "58"
					iDx = "67"	
					strJobCaption = C_JobQtyNm12
				  Case "59"
					iDx = "68"					
				  Case "60"
					iDx = "69"
							
				  Case "61"
					iDx = "70"	
				  Case "62"
					iDx = "71"
				  Case "63"
					iDx = "72"
					strJobCaption = C_JobQtyNm13													
				  Case "64"
					iDx = "73"	
				  Case "65"
					iDx = "74"
					
				  Case "66"
					iDx = "75"		
				  Case "67"
					iDx = "76"					
				  Case "68"
					iDx = "77"
					strJobCaption = C_JobQtyNm14
				  Case "69"
					iDx = "78"							
				  Case "70"
					iDx = "79"	
					
				  Case "71"
					iDx = "80"					
				  Case "72"
					iDx = "81"		
				  Case "73"
					iDx = "82"
					strJobCaption = C_JobQtyNm15
				  Case "74"
					iDx = "83"
				  Case "75"
					iDx = "84"
										
				  Case "76"
					iDx = "85"
				  Case "77"
					iDx = "86"
				  Case "78"
					iDx = "87"
					strJobCaption = C_JobQtyNm16
				  Case "79"
					iDx = "88"										
				  Case "80"
					iDx = "89"
					
				  Case "81"
					iDx = "90"
				  Case "82"
					iDx = "91"
				  Case "83"
					iDx = "92"
					strJobCaption = C_JobQtyNm17
				  Case "84"
					iDx = "93"
				  Case "85"
					iDx = "94"
					
				  Case "86"
					iDx = "95"
				  Case "87"
					iDx = "96"					
				  Case "88"
					iDx = "97"
					strJobCaption = C_JobQtyNm18
				  Case "89"
					iDx = "98"
				  Case "90"
					iDx = "99"
					
				  Case "91"
					iDx = "100"					
				  Case "92"
					iDx = "101"
				  Case "93"
					iDx = "102"
					strJobCaption = C_JobQtyNm19
				  Case "94"
					iDx = "103"
				  Case "95"
					iDx = "104"
										
				  Case "96"
					iDx = "105"
				  Case "97"
					iDx = "106"
				  Case "98"
					iDx = "107"
					strJobCaption = C_JobQtyNm20
				  Case "99"
					iDx = "108"			
				  Case "100"
					iDx = "109"
					
				  Case "101"
					iDx = "110"
				  Case "102"
					iDx = "111"
				  Case "103"
					iDx = "112"
					strJobCaption = C_JobQtyNm21
				  Case "104"
					iDx = "113"
				  Case "105"
					iDx = "114"
										
				  Case "106"
					iDx = "115"
				  Case "107"
					iDx = "116"
				  Case "108"
					iDx = "117"
					strJobCaption = C_JobQtyNm22
				  Case "109"
					iDx = "118"
				  Case "110"
					iDx = "119"
					
				  Case "111"
					iDx = "120"
				  Case "112"
					iDx = "121"
				  Case "113"
					iDx = "122"
					strJobCaption = C_JobQtyNm23
				  Case "114"
					iDx = "123"
				  Case "115"
					iDx = "124"	
					
				  Case "116"
					iDx = "125"
				  Case "117"
					iDx = "126"
				  Case "118"
					iDx = "127"
					strJobCaption = C_JobQtyNm24
				  Case "119"
					iDx = "128"
				  Case "120"
					iDx = "129"					 						 															  					 		
				End Select	

			if (iDx Mod 5 = 0) then 
			    ggoSpread.SSSetEdit		iDx, "작업계획시간" , 4
				Call ggoSpread.SSSetColHidden(iDx, iDx, True)
			elseif (iDx = "11" or iDx = "16" or iDx = "21" or iDx = "26" or iDx = "31" or iDx = "36" or iDx = "41" or iDx = "46" or iDx = "51" or iDx = "56" or iDx = "61" or iDx = "66" or iDx = "71" or iDx = "76" or iDx = "81" or iDx = "86" or iDx = "91" or iDx = "96" or iDx = "101" or iDx = "106" or iDx = "111" or iDx = "116" or iDx = "121" or iDx = "126") then  
				ggoSpread.SSSetEdit		iDx, "작업지시번호" , 13,,,18
				Call ggoSpread.SSSetColHidden(iDx, iDx, True)
			elseif (iDx = "12" or iDx = "17" or iDx = "22" or iDx = "27" or iDx = "32" or iDx = "37" or iDx = "42" or iDx = "47" or iDx = "52" or iDx = "57" or iDx = "62" or iDx = "67" or iDx = "72" or iDx = "77" or iDx = "82" or iDx = "87" or iDx = "92" or iDx = "97" or iDx = "102" or iDx = "107" or iDx = "112" or iDx = "117" or iDx = "122" or iDx = "127") then
				ggoSpread.SSSetFloat	iDx, strJobCaption ,8, "6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			elseif (iDx = "13" or iDx = "18" or iDx = "23" or iDx = "28" or iDx = "33" or iDx = "38" or iDx = "43" or iDx = "48" or iDx = "53" or iDx = "58" or iDx = "63" or iDx = "68" or iDx = "73" or iDx = "78" or iDx = "83" or iDx = "88" or iDx = "93" or iDx = "98" or iDx = "103" or iDx = "108" or iDx = "113" or iDx = "118" or iDx = "123" or iDx = "128") then
				ggoSpread.SSSetEdit	    iDx, "순서" ,4, 2, -1, 5
			else
				ggoSpread.SSSetCheck    iDx ,"확정",4 , ,"",true 							 								
			end if	

		Next

			Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
			Call ggoSpread.SSSetColHidden(C_JobLineCd, C_JobLineCd, True)
'			Call ggoSpread.SSSetColHidden(C_ProdtOrderNo, C_ProdtOrderNo, True)

			ggoSpread.SSSetSplit2(4)
			Call SetSpreadLock("A")
			Call InitData

			.ReDraw = true

		End With
    End If

	If pvSpdNo = "C" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 3 - Hidden Setting
		'------------------------------------------
		
		
		With frm1.vspdData3
			
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.Spreadinit
			
			.ReDraw = false
					
			.MaxCols = C_HJobPlanTime24 +1										'☜: 최대 Columns의 항상 1개 증가시킴 
			.MaxRows = 0
			
			ggoSpread.SSSetEdit		C_HJobPlanTime1,	"작업계획시간1", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime2,	"작업계획시간2", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime3,	"작업계획시간3", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime4,	"작업계획시간4", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime5,	"작업계획시간5", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime6,	"작업계획시간6", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime7,	"작업계획시간7", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime8,	"작업계획시간8", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime9,	"작업계획시간9", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime10,	"작업계획시간10", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime11,	"작업계획시간11", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime12,	"작업계획시간12", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime13,	"작업계획시간13", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime14,	"작업계획시간14", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime15,	"작업계획시간15", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime16,	"작업계획시간16", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime17,	"작업계획시간17", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime18,	"작업계획시간18", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime19,	"작업계획시간19", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime20,	"작업계획시간20", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime21,	"작업계획시간21", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime22,	"작업계획시간22", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime23,	"작업계획시간23", 18
			ggoSpread.SSSetEdit		C_HJobPlanTime24,	"작업계획시간24", 18
			
			.ReDraw = true
				
			Call SetSpreadLock("C")
    
		End With
    End If

End Sub

'============================ 2.2.4 SetSpreadLock() =====================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
	Dim ii, iDx
	Dim value1
	
	iDx = 0
	With frm1
'		If pvSpdNo = "A" Then
			'--------------------------------
			'Grid 1
			'--------------------------------
'			ggoSpread.Source = frm1.vspdData1
'			ggoSpread.SpreadLockWithOddEvenRowColor()
'		End If
			
		If pvSpdNo = "A" Then
			'--------------------------------
			'Grid 1
			'--------------------------------
			ggoSpread.Source = .vspdData1
			.vspdData1.ReDraw = False

    
				.vspdData1.ReDraw = False
				
				ggoSpread.SSSetRequired		C_JOBLINECD, -1, C_JOBLINECD
				ggoSpread.SSSetRequired		C_JobLine, -1, C_JobLine
				ggoSpread.SSSetProtected	C_ItemCd, -1, C_ItemCd 
				ggoSpread.SSSetProtected	C_ItemNm, -1, C_ItemNm 
				ggoSpread.SSSetProtected	C_Spec, -1, C_Spec
				ggoSpread.SSSetProtected	C_ProdtOrderNo, -1, C_Prodt_Order_No 
				ggoSpread.SSSetProtected	C_ProdtOrderPopup, -1, C_ProdtOrderPopup 
				ggoSpread.SSSetProtected	C_ProdtOrderQty, -1, C_ProdtOrderQty 
				ggoSpread.SSSetProtected	C_ProdtOrderLineSumQty, -1, C_ProdtOrderLineSumQty 
				
'			ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
			.vspdData1.ReDraw = True

		End If
			
		If pvSpdNo = "C" Then
		
			'--------------------------------
			'Grid 2
			'--------------------------------
			ggoSpread.Source = frm1.vspdData3
			.vspdData3.ReDraw = False
			ggoSpread.SpreadLock -1, -1
			.vspdData3.Redraw = True
		End If
    End With
End Sub

'============================= 2.2.5 SetSpreadColor() ===================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
	Dim i

    With frm1.vspdData1
    
		.Redraw = False
		ggoSpread.Source = frm1.vspdData1

		ggoSpread.SSSetProtected iDx,						pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_JobLineCd,				pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_JobLine,					pvStartRow, pvEndRow
'		ggoSpread.SSSetProtected C_ProdtOrderNo1,			pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_ProdtOrderNo,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemCd,					pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ItemNm,					pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Spec,					pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ProdtOrderQty,			pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_ProdtOrderLineSumQty,	pvStartRow, pvEndRow		
		
'		For i = 1  to C_GRIDCOUNT 
'			iDx = Cstr(((int(i)*5)+9))
'			.Col = iDx 
'			ggoSpread.SSSetProtected	iDx, pvStartRow, pvEndRow
'		Next
										
		.Redraw = True
    
    End With
   
End Sub

Sub SetSpreadLock2()
	Dim i
	Dim ii
	Dim value1
	
 with frm1.vspddata1
 .Redraw = False
  ggoSpread.Source = frm1.vspdData1
   			
   	   For  i = 1 to  .maxrows		
			 For ii = 1  to C_GRIDCOUNT * 5
      			Select Case ii 
				  Case "1" 
 					iDx = "10"                       
				  Case "2"	                        
  					iDx = "11"                       
				  Case "3"
  					iDx = "12"                       
  				  Case "4"	
  					iDx = "13"                      
				  Case "5"
					iDx = "14"
				  Case "6"
					iDx = "15"
				  Case "7"
				    iDx = "16"
				  Case "8"
					iDx = "17"
				  Case "9"
					iDx = "18"
				  Case "10"
					iDx = "19"
				  Case "11"
					iDx = "20"
				  Case "12"
					iDx = "21"
				  Case "13"
					iDx = "22"
				  Case "14"
					iDx = "23"
				  Case "15"
					iDx = "24"
				  Case "16"
					iDx = "25"
				  Case "17"
					iDx = "26"
				  Case "18"
					iDx = "27"
				  Case "19"
					iDx = "28"
				  Case "20"
					iDx = "29"
				  Case "21"
					iDx = "30"
				  Case "22"
					iDx = "31"
				  Case "23"
					iDx = "32"
				  Case "24"
					iDx = "33"
				  Case "25"
					iDx = "34"
 				  Case "26"
					iDx = "35"
  				  Case "27"
					iDx = "36"
				  Case "28"
					iDx = "37"
				  Case "29"
					iDx = "38"
				  Case "30"
					iDx = "39"
				  Case "31"
					iDx = "40"
				  Case "32"
					iDx = "41"
				  Case "33"
					iDx = "42"
				  Case "34"
					iDx = "43"
				  Case "35"
					iDx = "44"
				  Case "36"
					iDx = "45"
				  Case "37"
					iDx = "46"
				  Case "38"
					iDx = "47"
				  Case "39"
					iDx = "48"
				  Case "40"
					iDx = "49"
				  Case "41"
					iDx = "50"
				  Case "42"
					iDx = "51"
				  Case "43"
					iDx = "52"
				  Case "44"
					iDx = "53"	
				  Case "45"
					iDx = "54"		
				  Case "46"
					iDx = "55"	
				  Case "47"
					iDx = "56"
				  Case "48"
					iDx = "57"	
				  Case "49"
					iDx = "58"	
				  Case "50"
					iDx = "59"
				  Case "51"
					iDx = "60"
				  Case "52"
					iDx = "61"	
				  Case "53"
					iDx = "62"
				  Case "54"
					iDx = "63"		
				  Case "55"
					iDx = "64"
				  Case "56"
					iDx = "65"
				  Case "57"
					iDx = "66"		
				  Case "58"
					iDx = "67"	
				  Case "59"
					iDx = "68"
				  Case "60"
					iDx = "69"		
				  Case "61"
					iDx = "70"	
				  Case "62"
					iDx = "71"
				  Case "63"
					iDx = "72"
				  Case "64"
					iDx = "73"	
				  Case "65"
					iDx = "74"
				  Case "66"
					iDx = "75"		
				  Case "67"
					iDx = "76"	
				  Case "68"
					iDx = "77"
				  Case "69"
					iDx = "78"		
				  Case "70"
					iDx = "79"	
				  Case "71"
					iDx = "80"
				  Case "72"
					iDx = "81"		
				  Case "73"
					iDx = "82"
				  Case "74"
					iDx = "83"
				  Case "75"
					iDx = "84"
				  Case "76"
					iDx = "85"
				  Case "77"
					iDx = "86"
				  Case "78"
					iDx = "87"
				  Case "79"
					iDx = "88"
				  Case "80"
					iDx = "89"
				  Case "81"
					iDx = "90"
				  Case "82"
					iDx = "91"
				  Case "83"
					iDx = "92"
				  Case "84"
					iDx = "93"
				  Case "85"
					iDx = "94"
				  Case "86"
					iDx = "95"
				  Case "87"
					iDx = "96"
				  Case "88"
					iDx = "97"
				  Case "89"
					iDx = "98"
				  Case "90"
					iDx = "99"
				  Case "91"
					iDx = "100"
				  Case "92"
					iDx = "101"
				  Case "93"
					iDx = "102"
				  Case "94"
					iDx = "103"
				  Case "95"
					iDx = "104"
				  Case "96"
					iDx = "105"
				  Case "97"
					iDx = "106"
				  Case "98"
					iDx = "107"
				  Case "99"
					iDx = "108"
				  Case "100"
					iDx = "109"
				  Case "101"
					iDx = "110"
				  Case "102"
					iDx = "111"
				  Case "103"
					iDx = "112"	
				  Case "104"
					iDx = "113"
				  Case "105"
					iDx = "114"
				  Case "106"
					iDx = "115"
				  Case "107"
					iDx = "116"
				  Case "108"
					iDx = "117"
				  Case "109"
					iDx = "118"
				  Case "110"
					iDx = "119"
				  Case "111"
					iDx = "120"
				  Case "112"
					iDx = "121"
				  Case "113"
					iDx = "122"
				  Case "114"
					iDx = "123"
				  Case "115"
					iDx = "124"
				  Case "116"
					iDx = "125"
				  Case "117"
					iDx = "126"
				  Case "118"
					iDx = "127"
				  Case "119"
					iDx = "128"
				  Case "120"
					iDx = "129"									
                End Select
                
'				if (iDx = "11" or iDx = "16" or iDx = "21" or iDx = "26" or iDx = "31" or iDx = "36" or iDx = "41" or iDx = "46" or iDx = "51" or iDx = "56" or iDx = "61" or iDx = "66" or iDx = "71" or iDx = "76" or iDx = "81" or iDx = "86" or iDx = "91" or iDx = "96" or iDx = "101" or iDx = "106" or iDx = "111" or iDx = "116" or iDx = "121" or iDx = "126") then  
'					 .Col = iDx
'					 .Row = i
'					value1 = .value					
'				    if value1 = "" then
'						iDx = Cstr(int(iDx)+2)
'				        ggoSpread.SSSetProtected	iDx, i, i 
'				        iDx = Cstr(int(iDx)+1)
'				        ggoSpread.SSSetProtected	iDx, i, i 				        
'				    else 			            
'				    end if 
				if (iDx = "14" or iDx = "19" or iDx = "24" or iDx = "29" or iDx = "34" or iDx = "39" or iDx = "44" or iDx = "49" or iDx = "54" or iDx = "59" or iDx = "64" or iDx = "69" or iDx = "74" or iDx = "79" or iDx = "84" or iDx = "89" or iDx = "94" or iDx = "99" or iDx = "104" or iDx = "109" or iDx = "114" or iDx = "119" or iDx = "124" or iDx = "129") then
					.Col = iDx
					.Row = i
					Flag = .value
					if Flag = "1" then
						iDx = Cstr(int(iDx)-1)
						ggoSpread.SSSetProtected	iDx, i, i 
						iDx = Cstr(int(iDx)-1)
						ggoSpread.SSSetProtected	iDx, i, i 
					end if
				End if
			Next		
		  Next
	.Redraw = true   
  End With  
   
End Sub


'========================== 2.2.6 InitSpreadComboBox()  ========================================================
'	Name : InitSpreadComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitSpreadComboBox()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	Dim strWhere
	
	if frm1.txtPlantCd.value = "" then
	strWhere	=			 " plant_cd >= '" & frm1.txtPlantCd.value & "'"
	else
	strWhere	=			 " plant_cd = '" & frm1.txtPlantCd.value & "'"
	end if
	strWhere	= strWhere & " order by line_group, work_line "

	'****************************
	'List Minor code
	'****************************
	Call CommonQueryRs(" WORK_LINE, WORK_LINE_DESC ", " p_work_line_ko119 ", strWhere , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	
	ggoSpread.Source = frm1.vspdData1
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_JobLineCd
    ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_JobLine
End Sub

'==========================================  2.2.6 InitShiftCombo()  =======================================
'	Name : InitShiftCombo()
'	Description : Combo Display
'===========================================================================================================
Function InitShiftCombo()
	
	Dim strPlantCd
	Dim strShiftArr
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	
	strPlantCd = FilterVar(UCase(frm1.hPlantCd.value), "''", "S")
	
	'****************************
	'List Minor code(Reason Code)
	'****************************
    Call CommonQueryRs(" SHIFT_CD "," P_SHIFT_HEADER "," PLANT_CD = " & strPlantCd ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	ggoSpread.Source = frm1.vspdData2
    ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_ShiftId
    
    strShiftArr = Split(lgF0,Chr(11))
    
    lgShift = strShiftArr(0)
	
End Function

'==========================================  2.2.6 InitData()  =======================================
'	Name : InitData()
'	Description : Combo Display
'========================================================================================================= 
 Sub InitData()
	Dim intRow
	Dim intIndex
	
	With frm1.vspdData1
		For intRow = 1 to .MaxRows
			.Row = intRow
			.col = C_JobLineCd
			intIndex = .value
			.Col = C_JobLine
			.value = intindex
		Next	
	End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables() =================
' Function Name : InitSpreadPosVariables
' Function Desc : This method Assigns Sequential Number to spread sheet column 
'========================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	
	If pvSpdNo = "A" or pvSpdNo = "*" Then
		' Grid 1(vspdData1) - Results

		 C_JobLineCd					= 1
		 C_JobLine						= 2
		 C_ItemCd                       = 3
		 C_ItemNm                       = 4
		 C_Spec                         = 5
		 C_ProdtOrderNo				    = 6
		 
		 
	   Select Case C_GRIDCOUNT
		 Case "1"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10
		 Case "2"
		    C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
		 Case "3"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
		 Case "4"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
		Case "5"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
		Case "6"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
		Case "7"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
		Case "8"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
			C_JobPlanTime8              = 35
			C_JobOrderNo8				= 36
			C_JobQty8					= 37
			C_JobSeq8                   = 38 
		Case "9"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
			C_JobPlanTime8              = 35
			C_JobOrderNo8				= 36
			C_JobQty8					= 37
			C_JobSeq8                   = 38 
			C_JobPlanTime9              = 39
			C_JobOrderNo9				= 40
			C_JobQty9					= 41
			C_JobSeq9                   = 42 
		Case "10"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
			C_JobPlanTime8              = 35
			C_JobOrderNo8				= 36
			C_JobQty8					= 37
			C_JobSeq8                   = 38 
			C_JobPlanTime9              = 39
			C_JobOrderNo9				= 40
			C_JobQty9					= 41
			C_JobSeq9                   = 42 
			C_JobPlanTime10             = 43
			C_JobOrderNo10				= 44
			C_JobQty10					= 45
			C_JobSeq10                  = 46 
		Case "11"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
			C_JobPlanTime8              = 35
			C_JobOrderNo8				= 36
			C_JobQty8					= 37
			C_JobSeq8                   = 38 
			C_JobPlanTime9              = 39
			C_JobOrderNo9				= 40
			C_JobQty9					= 41
			C_JobSeq9                   = 42 
			C_JobPlanTime10             = 43
			C_JobOrderNo10				= 44
			C_JobQty10					= 45
			C_JobSeq10                  = 46 
			C_JobPlanTime11             = 47
			C_JobOrderNo11				= 48
			C_JobQty11					= 49
			C_JobSeq11                  = 50 
		Case "12"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
			C_JobPlanTime8              = 35
			C_JobOrderNo8				= 36
			C_JobQty8					= 37
			C_JobSeq8                   = 38 
			C_JobPlanTime9              = 39
			C_JobOrderNo9				= 40
			C_JobQty9					= 41
			C_JobSeq9                   = 42 
			C_JobPlanTime10             = 43
			C_JobOrderNo10				= 44
			C_JobQty10					= 45
			C_JobSeq10                  = 46 
			C_JobPlanTime11             = 47
			C_JobOrderNo11				= 48
			C_JobQty11					= 49
			C_JobSeq11                  = 50 
			C_JobPlanTime12             = 51
			C_JobOrderNo12				= 52
			C_JobQty12					= 53
			C_JobSeq12                  = 54 
		Case "13"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
			C_JobPlanTime8              = 35
			C_JobOrderNo8				= 36
			C_JobQty8					= 37
			C_JobSeq8                   = 38 
			C_JobPlanTime9              = 39
			C_JobOrderNo9				= 40
			C_JobQty9					= 41
			C_JobSeq9                   = 42 
			C_JobPlanTime10             = 43
			C_JobOrderNo10				= 44
			C_JobQty10					= 45
			C_JobSeq10                  = 46 
			C_JobPlanTime11             = 47
			C_JobOrderNo11				= 48
			C_JobQty11					= 49
			C_JobSeq11                  = 50 
			C_JobPlanTime12             = 51
			C_JobOrderNo12				= 52
			C_JobQty12					= 53
			C_JobSeq12                  = 54 
			C_JobPlanTime13             = 55
			C_JobOrderNo13				= 56
			C_JobQty13					= 57
			C_JobSeq13                  = 58
		Case "14"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
			C_JobPlanTime8              = 35
			C_JobOrderNo8				= 36
			C_JobQty8					= 37
			C_JobSeq8                   = 38 
			C_JobPlanTime9              = 39
			C_JobOrderNo9				= 40
			C_JobQty9					= 41
			C_JobSeq9                   = 42 
			C_JobPlanTime10             = 43
			C_JobOrderNo10				= 44
			C_JobQty10					= 45
			C_JobSeq10                  = 46 
			C_JobPlanTime11             = 47
			C_JobOrderNo11				= 48
			C_JobQty11					= 49
			C_JobSeq11                  = 50 
			C_JobPlanTime12             = 51
			C_JobOrderNo12				= 52
			C_JobQty12					= 53
			C_JobSeq12                  = 54 
			C_JobPlanTime13             = 55
			C_JobOrderNo13				= 56
			C_JobQty13					= 57
			C_JobSeq13                  = 58
			C_JobPlanTime14             = 59
			C_JobOrderNo14				= 60
			C_JobQty14					= 61
			C_JobSeq14                  = 62
		Case "15"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
			C_JobPlanTime8              = 35
			C_JobOrderNo8				= 36
			C_JobQty8					= 37
			C_JobSeq8                   = 38 
			C_JobPlanTime9              = 39
			C_JobOrderNo9				= 40
			C_JobQty9					= 41
			C_JobSeq9                   = 42 
			C_JobPlanTime10             = 43
			C_JobOrderNo10				= 44
			C_JobQty10					= 45
			C_JobSeq10                  = 46 
			C_JobPlanTime11             = 47
			C_JobOrderNo11				= 48
			C_JobQty11					= 49
			C_JobSeq11                  = 50 
			C_JobPlanTime12             = 51
			C_JobOrderNo12				= 52
			C_JobQty12					= 53
			C_JobSeq12                  = 54 
			C_JobPlanTime13             = 55
			C_JobOrderNo13				= 56
			C_JobQty13					= 57
			C_JobSeq13                  = 58
			C_JobPlanTime14             = 59
			C_JobOrderNo14				= 60
			C_JobQty14					= 61
			C_JobSeq14                  = 62
			C_JobPlanTime15             = 63
			C_JobOrderNo15				= 64
			C_JobQty15					= 65
			C_JobSeq15                  = 66
		Case "16"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
			C_JobPlanTime8              = 35
			C_JobOrderNo8				= 36
			C_JobQty8					= 37
			C_JobSeq8                   = 38 
			C_JobPlanTime9              = 39
			C_JobOrderNo9				= 40
			C_JobQty9					= 41
			C_JobSeq9                   = 42 
			C_JobPlanTime10             = 43
			C_JobOrderNo10				= 44
			C_JobQty10					= 45
			C_JobSeq10                  = 46 
			C_JobPlanTime11             = 47
			C_JobOrderNo11				= 48
			C_JobQty11					= 49
			C_JobSeq11                  = 50 
			C_JobPlanTime12             = 51
			C_JobOrderNo12				= 52
			C_JobQty12					= 53
			C_JobSeq12                  = 54 
			C_JobPlanTime13             = 55
			C_JobOrderNo13				= 56
			C_JobQty13					= 57
			C_JobSeq13                  = 58
			C_JobPlanTime14             = 59
			C_JobOrderNo14				= 60
			C_JobQty14					= 61
			C_JobSeq14                  = 62
			C_JobPlanTime15             = 63
			C_JobOrderNo15				= 64
			C_JobQty15					= 65
			C_JobSeq15                  = 66
			C_JobPlanTime16             = 67
			C_JobOrderNo16				= 68
			C_JobQty16					= 69
			C_JobSeq16                  = 70
		Case "17"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
			C_JobPlanTime8              = 35
			C_JobOrderNo8				= 36
			C_JobQty8					= 37
			C_JobSeq8                   = 38 
			C_JobPlanTime9              = 39
			C_JobOrderNo9				= 40
			C_JobQty9					= 41
			C_JobSeq9                   = 42 
			C_JobPlanTime10             = 43
			C_JobOrderNo10				= 44
			C_JobQty10					= 45
			C_JobSeq10                  = 46 
			C_JobPlanTime11             = 47
			C_JobOrderNo11				= 48
			C_JobQty11					= 49
			C_JobSeq11                  = 50 
			C_JobPlanTime12             = 51
			C_JobOrderNo12				= 52
			C_JobQty12					= 53
			C_JobSeq12                  = 54 
			C_JobPlanTime13             = 55
			C_JobOrderNo13				= 56
			C_JobQty13					= 57
			C_JobSeq13                  = 58
			C_JobPlanTime14             = 59
			C_JobOrderNo14				= 60
			C_JobQty14					= 61
			C_JobSeq14                  = 62
			C_JobPlanTime15             = 63
			C_JobOrderNo15				= 64
			C_JobQty15					= 65
			C_JobSeq15                  = 66
			C_JobPlanTime16             = 67
			C_JobOrderNo16				= 68
			C_JobQty16					= 69
			C_JobSeq16                  = 70
			C_JobPlanTime17             = 71
			C_JobOrderNo17				= 72
			C_JobQty17					= 73
			C_JobSeq17                  = 74
		Case "18"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
			C_JobPlanTime8              = 35
			C_JobOrderNo8				= 36
			C_JobQty8					= 37
			C_JobSeq8                   = 38 
			C_JobPlanTime9              = 39
			C_JobOrderNo9				= 40
			C_JobQty9					= 41
			C_JobSeq9                   = 42 
			C_JobPlanTime10             = 43
			C_JobOrderNo10				= 44
			C_JobQty10					= 45
			C_JobSeq10                  = 46 
			C_JobPlanTime11             = 47
			C_JobOrderNo11				= 48
			C_JobQty11					= 49
			C_JobSeq11                  = 50 
			C_JobPlanTime12             = 51
			C_JobOrderNo12				= 52
			C_JobQty12					= 53
			C_JobSeq12                  = 54 
			C_JobPlanTime13             = 55
			C_JobOrderNo13				= 56
			C_JobQty13					= 57
			C_JobSeq13                  = 58
			C_JobPlanTime14             = 59
			C_JobOrderNo14				= 60
			C_JobQty14					= 61
			C_JobSeq14                  = 62
			C_JobPlanTime15             = 63
			C_JobOrderNo15				= 64
			C_JobQty15					= 65
			C_JobSeq15                  = 66
			C_JobPlanTime16             = 67
			C_JobOrderNo16				= 68
			C_JobQty16					= 69
			C_JobSeq16                  = 70
			C_JobPlanTime17             = 71
			C_JobOrderNo17				= 72
			C_JobQty17					= 73
			C_JobSeq17                  = 74
			C_JobPlanTime18             = 75
			C_JobOrderNo18				= 76
			C_JobQty18					= 77
			C_JobSeq18                  = 78
		Case "19"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
			C_JobPlanTime8              = 35
			C_JobOrderNo8				= 36
			C_JobQty8					= 37
			C_JobSeq8                   = 38 
			C_JobPlanTime9              = 39
			C_JobOrderNo9				= 40
			C_JobQty9					= 41
			C_JobSeq9                   = 42 
			C_JobPlanTime10             = 43
			C_JobOrderNo10				= 44
			C_JobQty10					= 45
			C_JobSeq10                  = 46 
			C_JobPlanTime11             = 47
			C_JobOrderNo11				= 48
			C_JobQty11					= 49
			C_JobSeq11                  = 50 
			C_JobPlanTime12             = 51
			C_JobOrderNo12				= 52
			C_JobQty12					= 53
			C_JobSeq12                  = 54 
			C_JobPlanTime13             = 55
			C_JobOrderNo13				= 56
			C_JobQty13					= 57
			C_JobSeq13                  = 58
			C_JobPlanTime14             = 59
			C_JobOrderNo14				= 60
			C_JobQty14					= 61
			C_JobSeq14                  = 62
			C_JobPlanTime15             = 63
			C_JobOrderNo15				= 64
			C_JobQty15					= 65
			C_JobSeq15                  = 66
			C_JobPlanTime16             = 67
			C_JobOrderNo16				= 68
			C_JobQty16					= 69
			C_JobSeq16                  = 70
			C_JobPlanTime17             = 71
			C_JobOrderNo17				= 72
			C_JobQty17					= 73
			C_JobSeq17                  = 74
			C_JobPlanTime18             = 75
			C_JobOrderNo18				= 76
			C_JobQty18					= 77
			C_JobSeq18                  = 78
			C_JobPlanTime19             = 79
			C_JobOrderNo19				= 80
			C_JobQty19					= 81
			C_JobSeq19                  = 82
		Case "20"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
			C_JobPlanTime8              = 35
			C_JobOrderNo8				= 36
			C_JobQty8					= 37
			C_JobSeq8                   = 38 
			C_JobPlanTime9              = 39
			C_JobOrderNo9				= 40
			C_JobQty9					= 41
			C_JobSeq9                   = 42 
			C_JobPlanTime10             = 43
			C_JobOrderNo10				= 44
			C_JobQty10					= 45
			C_JobSeq10                  = 46 
			C_JobPlanTime11             = 47
			C_JobOrderNo11				= 48
			C_JobQty11					= 49
			C_JobSeq11                  = 50 
			C_JobPlanTime12             = 51
			C_JobOrderNo12				= 52
			C_JobQty12					= 53
			C_JobSeq12                  = 54 
			C_JobPlanTime13             = 55
			C_JobOrderNo13				= 56
			C_JobQty13					= 57
			C_JobSeq13                  = 58
			C_JobPlanTime14             = 59
			C_JobOrderNo14				= 60
			C_JobQty14					= 61
			C_JobSeq14                  = 62
			C_JobPlanTime15             = 63
			C_JobOrderNo15				= 64
			C_JobQty15					= 65
			C_JobSeq15                  = 66
			C_JobPlanTime16             = 67
			C_JobOrderNo16				= 68
			C_JobQty16					= 69
			C_JobSeq16                  = 70
			C_JobPlanTime17             = 71
			C_JobOrderNo17				= 72
			C_JobQty17					= 73
			C_JobSeq17                  = 74
			C_JobPlanTime18             = 75
			C_JobOrderNo18				= 76
			C_JobQty18					= 77
			C_JobSeq18                  = 78
			C_JobPlanTime19             = 79
			C_JobOrderNo19				= 80
			C_JobQty19					= 81
			C_JobSeq19                  = 82
			C_JobPlanTime20             = 83
			C_JobOrderNo20				= 84
			C_JobQty20					= 85
			C_JobSeq20                  = 86
		Case "21"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
			C_JobPlanTime8              = 35
			C_JobOrderNo8				= 36
			C_JobQty8					= 37
			C_JobSeq8                   = 38 
			C_JobPlanTime9              = 39
			C_JobOrderNo9				= 40
			C_JobQty9					= 41
			C_JobSeq9                   = 42 
			C_JobPlanTime10             = 43
			C_JobOrderNo10				= 44
			C_JobQty10					= 45
			C_JobSeq10                  = 46 
			C_JobPlanTime11             = 47
			C_JobOrderNo11				= 48
			C_JobQty11					= 49
			C_JobSeq11                  = 50 
			C_JobPlanTime12             = 51
			C_JobOrderNo12				= 52
			C_JobQty12					= 53
			C_JobSeq12                  = 54 
			C_JobPlanTime13             = 55
			C_JobOrderNo13				= 56
			C_JobQty13					= 57
			C_JobSeq13                  = 58
			C_JobPlanTime14             = 59
			C_JobOrderNo14				= 60
			C_JobQty14					= 61
			C_JobSeq14                  = 62
			C_JobPlanTime15             = 63
			C_JobOrderNo15				= 64
			C_JobQty15					= 65
			C_JobSeq15                  = 66
			C_JobPlanTime16             = 67
			C_JobOrderNo16				= 68
			C_JobQty16					= 69
			C_JobSeq16                  = 70
			C_JobPlanTime17             = 71
			C_JobOrderNo17				= 72
			C_JobQty17					= 73
			C_JobSeq17                  = 74
			C_JobPlanTime18             = 75
			C_JobOrderNo18				= 76
			C_JobQty18					= 77
			C_JobSeq18                  = 78
			C_JobPlanTime19             = 79
			C_JobOrderNo19				= 80
			C_JobQty19					= 81
			C_JobSeq19                  = 82
			C_JobPlanTime20             = 83
			C_JobOrderNo20				= 84
			C_JobQty20					= 85
			C_JobSeq20                  = 86
			C_JobPlanTime21             = 87
			C_JobOrderNo21				= 88
			C_JobQty21					= 89
			C_JobSeq21                  = 90
		Case "22"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
			C_JobPlanTime8              = 35
			C_JobOrderNo8				= 36
			C_JobQty8					= 37
			C_JobSeq8                   = 38 
			C_JobPlanTime9              = 39
			C_JobOrderNo9				= 40
			C_JobQty9					= 41
			C_JobSeq9                   = 42 
			C_JobPlanTime10             = 43
			C_JobOrderNo10				= 44
			C_JobQty10					= 45
			C_JobSeq10                  = 46 
			C_JobPlanTime11             = 47
			C_JobOrderNo11				= 48
			C_JobQty11					= 49
			C_JobSeq11                  = 50 
			C_JobPlanTime12             = 51
			C_JobOrderNo12				= 52
			C_JobQty12					= 53
			C_JobSeq12                  = 54 
			C_JobPlanTime13             = 55
			C_JobOrderNo13				= 56
			C_JobQty13					= 57
			C_JobSeq13                  = 58
			C_JobPlanTime14             = 59
			C_JobOrderNo14				= 60
			C_JobQty14					= 61
			C_JobSeq14                  = 62
			C_JobPlanTime15             = 63
			C_JobOrderNo15				= 64
			C_JobQty15					= 65
			C_JobSeq15                  = 66
			C_JobPlanTime16             = 67
			C_JobOrderNo16				= 68
			C_JobQty16					= 69
			C_JobSeq16                  = 70
			C_JobPlanTime17             = 71
			C_JobOrderNo17				= 72
			C_JobQty17					= 73
			C_JobSeq17                  = 74
			C_JobPlanTime18             = 75
			C_JobOrderNo18				= 76
			C_JobQty18					= 77
			C_JobSeq18                  = 78
			C_JobPlanTime19             = 79
			C_JobOrderNo19				= 80
			C_JobQty19					= 81
			C_JobSeq19                  = 82
			C_JobPlanTime20             = 83
			C_JobOrderNo20				= 84
			C_JobQty20					= 85
			C_JobSeq20                  = 86
			C_JobPlanTime21             = 87
			C_JobOrderNo21				= 88
			C_JobQty21					= 89
			C_JobSeq21                  = 90
			C_JobPlanTime22             = 91
			C_JobOrderNo22				= 92
			C_JobQty22					= 93
			C_JobSeq22                  = 94
		Case "23"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
			C_JobPlanTime8              = 35
			C_JobOrderNo8				= 36
			C_JobQty8					= 37
			C_JobSeq8                   = 38 
			C_JobPlanTime9              = 39
			C_JobOrderNo9				= 40
			C_JobQty9					= 41
			C_JobSeq9                   = 42 
			C_JobPlanTime10             = 43
			C_JobOrderNo10				= 44
			C_JobQty10					= 45
			C_JobSeq10                  = 46 
			C_JobPlanTime11             = 47
			C_JobOrderNo11				= 48
			C_JobQty11					= 49
			C_JobSeq11                  = 50 
			C_JobPlanTime12             = 51
			C_JobOrderNo12				= 52
			C_JobQty12					= 53
			C_JobSeq12                  = 54 
			C_JobPlanTime13             = 55
			C_JobOrderNo13				= 56
			C_JobQty13					= 57
			C_JobSeq13                  = 58
			C_JobPlanTime14             = 59
			C_JobOrderNo14				= 60
			C_JobQty14					= 61
			C_JobSeq14                  = 62
			C_JobPlanTime15             = 63
			C_JobOrderNo15				= 64
			C_JobQty15					= 65
			C_JobSeq15                  = 66
			C_JobPlanTime16             = 67
			C_JobOrderNo16				= 68
			C_JobQty16					= 69
			C_JobSeq16                  = 70
			C_JobPlanTime17             = 71
			C_JobOrderNo17				= 72
			C_JobQty17					= 73
			C_JobSeq17                  = 74
			C_JobPlanTime18             = 75
			C_JobOrderNo18				= 76
			C_JobQty18					= 77
			C_JobSeq18                  = 78
			C_JobPlanTime19             = 79
			C_JobOrderNo19				= 80
			C_JobQty19					= 81
			C_JobSeq19                  = 82
			C_JobPlanTime20             = 83
			C_JobOrderNo20				= 84
			C_JobQty20					= 85
			C_JobSeq20                  = 86
			C_JobPlanTime21             = 87
			C_JobOrderNo21				= 88
			C_JobQty21					= 89
			C_JobSeq21                  = 90
			C_JobPlanTime22             = 91
			C_JobOrderNo22				= 92
			C_JobQty22					= 93
			C_JobSeq22                  = 94
			C_JobPlanTime23             = 95
			C_JobOrderNo23				= 96
			C_JobQty23					= 97
			C_JobSeq23                  = 98
		Case "24"
			C_JobPlanTime1              = 7
			C_JobOrderNo1				= 8
			C_JobQty1					= 9
			C_JobSeq1                   = 10 
		    C_JobPlanTime2              = 11
			C_JobOrderNo2				= 12
			C_JobQty2					= 13
			C_JobSeq2                   = 14 
			C_JobPlanTime3              = 15
			C_JobOrderNo3				= 16
			C_JobQty3					= 17
			C_JobSeq3                   = 18 
			C_JobPlanTime4              = 19
			C_JobOrderNo4				= 20
			C_JobQty4					= 21
			C_JobSeq4                   = 22 
			C_JobPlanTime5              = 23
			C_JobOrderNo5				= 24
			C_JobQty5					= 25
			C_JobSeq5                   = 26 
			C_JobPlanTime6              = 27
			C_JobOrderNo6				= 28
			C_JobQty6					= 29
			C_JobSeq6                   = 30 
			C_JobPlanTime7              = 31
			C_JobOrderNo7				= 32
			C_JobQty7					= 33
			C_JobSeq7                   = 34 
			C_JobPlanTime8              = 35
			C_JobOrderNo8				= 36
			C_JobQty8					= 37
			C_JobSeq8                   = 38 
			C_JobPlanTime9              = 39
			C_JobOrderNo9				= 40
			C_JobQty9					= 41
			C_JobSeq9                   = 42 
			C_JobPlanTime10             = 43
			C_JobOrderNo10				= 44
			C_JobQty10					= 45
			C_JobSeq10                  = 46 
			C_JobPlanTime11             = 47
			C_JobOrderNo11				= 48
			C_JobQty11					= 49
			C_JobSeq11                  = 50 
			C_JobPlanTime12             = 51
			C_JobOrderNo12				= 52
			C_JobQty12					= 53
			C_JobSeq12                  = 54 
			C_JobPlanTime13             = 55
			C_JobOrderNo13				= 56
			C_JobQty13					= 57
			C_JobSeq13                  = 58
			C_JobPlanTime14             = 59
			C_JobOrderNo14				= 60
			C_JobQty14					= 61
			C_JobSeq14                  = 62
			C_JobPlanTime15             = 63
			C_JobOrderNo15				= 64
			C_JobQty15					= 65
			C_JobSeq15                  = 66
			C_JobPlanTime16             = 67
			C_JobOrderNo16				= 68
			C_JobQty16					= 69
			C_JobSeq16                  = 70
			C_JobPlanTime17             = 71
			C_JobOrderNo17				= 72
			C_JobQty17					= 73
			C_JobSeq17                  = 74
			C_JobPlanTime18             = 75
			C_JobOrderNo18				= 76
			C_JobQty18					= 77
			C_JobSeq18                  = 78
			C_JobPlanTime19             = 79
			C_JobOrderNo19				= 80
			C_JobQty19					= 81
			C_JobSeq19                  = 82
			C_JobPlanTime20             = 83
			C_JobOrderNo20				= 84
			C_JobQty20					= 85
			C_JobSeq20                  = 86
			C_JobPlanTime21             = 87
			C_JobOrderNo21				= 88
			C_JobQty21					= 89
			C_JobSeq21                  = 90
			C_JobPlanTime22             = 91
			C_JobOrderNo22				= 92
			C_JobQty22					= 93
			C_JobSeq22                  = 94
			C_JobPlanTime23             = 95
			C_JobOrderNo23				= 96
			C_JobQty23					= 97
			C_JobSeq23                  = 98
			C_JobPlanTime24             = 99
			C_JobOrderNo24				= 100
			C_JobQty24					= 101
			C_JobSeq24                  = 102
		End Select
	End If		 
	
'	If pvSpdNo = "C" Or pvSpdNo = "*" Then
'		' Grid 3(vspdData3) - Hidden
'		 C_ReportDt2				= 1
'		 C_ReportType2				= 2
'		 C_ShiftId2					= 3
'		 C_ProdQty2					= 4
'		 C_ReasonCd2				= 5
'		 C_ReasonDesc2				= 6
'		 C_Remark2					= 7
'		 C_LotNo2					= 8
'		 C_LotSubNo2				= 9
'		 C_RcptDocumentNo2			= 10
'		 C_IssueDocumentNo2			= 11
'		 C_InspReqNo2				= 12
'		 C_Insp_Good_Qty2			= 13
'		 C_Insp_Bad_Qty2			= 14
'		 C_Rcpt_Qty2				= 15
'		 C_ProdtOrderNo2			= 16
'		 C_OprNo2					= 17
'		 C_Sequence2				= 18
'		 C_MilestoneFlg2			= 19
'		 C_InsideFlag2				= 20
'		 C_AutoRcptFlg2				= 21
'		 C_LotReq2					= 22
'		 C_LotGenMthd2				= 23
'		 C_ProdInspReq2				= 24
'		 C_FinalInspReq2			= 25
'		 C_RoutOrder2				= 26
'	End If
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==========
' Function Name : GetSpreadColumnPos
' Function Desc : This method is used to get specific spreadsheet column position according to the arguement
'=================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)

    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
 			ggoSpread.Source = frm1.vspdData1
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			 C_ProdtOrderNo				= iCurColumnPos(1)
			 C_ItemCd					= iCurColumnPos(2)
			 C_ItemNm					= iCurColumnPos(3)
			 C_Spec						= iCurColumnPos(4)
			 C_SecItemCd				= iCurColumnPos(5)
			 C_RoutNo					= iCurColumnPos(6)
			 C_ProdtOrderQty			= iCurColumnPos(7)
			 C_ProdtOrderSumQty			= iCurColumnPos(8)
			 C_BaseUnit					= iCurColumnPos(9)
			 C_TrackingNo				= iCurColumnPos(10)
			 C_LampMaker				= iCurColumnPos(11)
			 C_InvertMaker				= iCurColumnPos(12)


		Case "B"

			ggoSpread.Source = frm1.vspdData2
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			' Grid 2(vspdData2) - Results
			 C_JobLine					=	iCurColumnPos(1)
			 C_ProdtOrderNo1			=   iCurColumnPos(2)
			 
		 Select Case C_GRIDCOUNT
			 Case "1"
				C_JobOrderNo1			=	iCurColumnPos(3)
				C_JobQty				=	iCurColumnPos(4)
			 Case "2"
				C_JobOrderNo1			=	iCurColumnPos(3)
				C_JobQty1				=	iCurColumnPos(4)
				C_JobOrderNo2			=	iCurColumnPos(5)
				C_JobQty2				=	iCurColumnPos(6)
		End Select
    End Select 
End Sub    

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'*********************************************************************************************************

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : 중복되어 있는 PopUp을 재정의, 재정의가 필요한 경우는 반드시 CommonPopUp.vbs 와 
'				  ManufactPopUp.vbs 에서 Copy하여 재정의한다.
'=========================================================================================================
'++++++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  OpenPlant()  ------------------------------------------------
'	Name : OpenPlant()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlant()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"				' 팝업 명칭 
	arrParam(1) = "B_PLANT"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "공장"					' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"					' Field명(0)
    arrField(1) = "PLANT_NM"					' Field명(1)
    
    arrHeader(0) = "공장"					' Header명(0)
    arrHeader(1) = "공장명"					' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
		Call initCombobox2()
	    Call initCombobox()
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenProdOrderNo()  ------------------------------------------
'	Name : OpenProdOrderNo()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProdOrderNo()

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName

	If IsOpenPop = True Or UCase(frm1.txtProdOrderNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	If frm1.txtProdFromDt.Text= "" Then
		Call displaymsgbox("971012","X", "계획일","X")
		frm1.txtProdFromDt.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4111PA1_ko119")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111PA1_ko119", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtProdFromDt.Text
	arrParam(2) = frm1.txtProdFromDt.Text
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	arrParam(5) = Trim(frm1.txtProdOrderNo.value)
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = ""	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetProdOrderNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtProdOrderNo.focus
	
End Function

'------------------------------------------  OpenProdOrderNo()  ------------------------------------------
'	Name : OpenProdOrderNo1()
'	Description : Condition Production Order PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenProdOrderNo1(byval strCon)

	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	Dim JobLineCd1

	If IsOpenPop = True  Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	If frm1.txtProdFromDt.Text= "" Then
		Call displaymsgbox("971012","X", "계획일","X")
		frm1.txtProdFromDt.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
'	With frm1.vspddata1			    
'	    .col = C_JobLineCd
'	    .row = .activerow	
'		JobLineCd1 =  .value
		
'		if JobLineCd1 = "" then
'			Call displaymsgbox("971012","X", "Line","X")
'			.focus 
'			Set gActiveElement = document.activeElement
'			Exit Function				    	
'		end if	

'	End With	
						
	
	iCalledAspName = AskPRAspName("p4111pa1_ko119")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "p4111pa1_ko119", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True	

	arrParam(0) = frm1.txtPlantCd.value
	arrParam(1) = frm1.txtProdFromDt.Text
	arrParam(2) = frm1.txtProdFromDt.Text
	arrParam(3) = "RL"
	arrParam(4) = "ST"
	arrParam(5) = Trim(strCon)
	arrParam(6) = ""
	arrParam(7) = ""
	arrParam(8) = ""	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		With frm1.vspdData1 
			.Row = .ActiveRow 
			.Col = C_ItemNm
			.text = arrRet(2)
			.Col = C_ItemCd
			.text = arrRet(1)
			.Col = C_Spec
			.text = arrRet(3) 
			.Col = C_ProdtOrderNo
			.text = arrRet(0) 
			Call SetFocusToDocument("M") 
			.focus
			Call vspdData1_Change(C_ProdtOrderNo,frm1.vspdData1.ActiveRow)
		End With 
	End If 		
End Function

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemInfo()

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
	arrField(0) = 1 '"ITEM_CD"					' Field명(0)
	arrField(1) = 2 '"ITEM_NM"					' Field명(1)
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'------------------------------------------  OpenItemGroup()  -------------------------------------------------
'	Name : OpenItemGroup()
'	Description : Item By Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"
	arrParam(1) = "B_ITEM_GROUP"
	arrParam(2) = Trim(UCase(frm1.txtItemGroupCd.Value))
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "
	arrParam(5) = "품목그룹"
	 
	arrField(0) = "ITEM_GROUP_CD"
	arrField(1) = "ITEM_GROUP_NM"
	    
	arrHeader(0) = "품목그룹"
	arrHeader(1) = "품목그룹명"
	    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If 
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd.focus
 
End Function


'------------------------------------------  OpenWcCd()  ------------------------------------------------
'	Name : OpenWcCd()
'	Description : Condition Work Center PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenWcCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	IsOpenPop = True

	arrParam(0) = "작업장팝업"											' 팝업 명칭 
	arrParam(1) = "P_WORK_CENTER"											' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtWCCd.Value)									' Code Condition
	arrParam(3) = ""'Trim(frm1.txtWCNm.Value)								' Name Cindition
	arrParam(4) = "PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")			' Where Condition
	arrParam(5) = "작업장"												' TextBox 명칭 
	
    arrField(0) = "WC_CD"													' Field명(0)
    arrField(1) = "WC_NM"													' Field명(1)
    
    arrHeader(0) = "작업장"												' Header명(0)
    arrHeader(1) = "작업장명"											' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetWcCd(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtWCCd.focus
	
End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTracking Info PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo()

	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName
	
	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function
	
	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = Trim(frm1.txtItemCd.value)
	arrParam(3) = frm1.txtProdFromDt.Text
'	arrParam(4) = frm1.txtProdFromDt.Text
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
	
End Function

'------------------------------------------  OpenPartRef()  ----------------------------------------------
'	Name : OpenPartRef()
'	Description : Part Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPartRef()
	Dim arrRet
	Dim arrParam(2)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4311RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4311RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	arrParam(0) = Trim(frm1.txtPlantCd.value)			'☆: 조회 조건 데이타 

	ggoSpread.Source = frm1.vspdData1

	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_ProdtOrderNo
		arrParam(1) = Trim(frm1.vspdData1.Text)				'☜: 조회 조건 데이타 
		frm1.vspdData1.Col = C_OprNo
		arrParam(2) = Trim(frm1.vspdData1.Text)				'☜: 조회 조건 데이타 
	End If	

	IsOpenPop = True
		
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2)), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function

'------------------------------------------  OpenOprRef()  -----------------------------------------------
'	Name : OpenOprRef()
'	Description : Operation Reference PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenOprRef()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4111RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4111RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)

	ggoSpread.Source = frm1.vspdData1

	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_ProdtOrderNo
		arrParam(1) = Trim(frm1.vspdData1.Text)				'☜: 조회 조건 데이타 
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenRcptRef()  ----------------------------------------------
'	Name : OpenRcptRef()
'	Description : Goods Receipt Reference
'---------------------------------------------------------------------------------------------------------
Function OpenRcptRef()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4511RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4511RA1", "X")
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)

	ggoSpread.Source = frm1.vspdData1

	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_ProdtOrderNo
		arrParam(1) = Trim(frm1.vspdData1.Text)				'☜: 조회 조건 데이타 
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
End Function

'------------------------------------------  OpenConsumRef()  --------------------------------------------
'	Name : OpenConsumRef()
'	Description : Part Consumption Reference
'---------------------------------------------------------------------------------------------------------
Function OpenConsumRef()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4412RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4412RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)

	ggoSpread.Source = frm1.vspdData1

	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_ProdtOrderNo
		arrParam(1) = Trim(frm1.vspdData1.Text)				'☜: 조회 조건 데이타 
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent,arrParam(0), arrParam(1)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenReworkRef()  --------------------------------------------
'	Name : OpenReworkRef()
'	Description : Rework Order History Reference
'---------------------------------------------------------------------------------------------------------
Function OpenReworkRef()

	Dim arrRet
	Dim arrParam(3)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call displaymsgbox("900002", "x", "x", "x")
		Exit Function
	End If

	If frm1.txtPlantCd.value= "" Then
		Call displaymsgbox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("P4413RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4413RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.hPlantCd.value)
	
	ggoSpread.Source = frm1.vspdData1

	If frm1.vspdData1.MaxRows > 0 Then
		frm1.vspdData1.Row = frm1.vspdData1.ActiveRow
		frm1.vspdData1.Col = C_ItemCd
		arrParam(1) = Trim(frm1.vspdData1.Text)				'☜: 조회 조건 데이타 
		frm1.vspdData1.Col = C_ProdtOrderNo
		arrParam(2) = Trim(frm1.vspdData1.Text)				'☜: 조회 조건 데이타 
		'opr_no
		arrParam(3) = ""									'☜: 조회 조건 데이타 
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam(0), arrParam(1), arrParam(2), arrParam(3)), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'------------------------------------------  OpenBackFlushRef()  -----------------------------------------
'	Name : OpenBackFlushRef()
'	Description : BackFlush Simmulation Reference
'---------------------------------------------------------------------------------------------------------
Function OpenBackFlushRef()
	
	Dim arrRet
	Dim IntRows
	Dim strVal
	Dim iCalledAspName
	Dim strFlag
	
	If IsOpenPop = True Then Exit Function
	
	strVal = ""
		
	With frm1.vspdData3
		For IntRows = 1 To .MaxRows
			.Row = IntRows
			.Col = 0
			strFlag = .Text
			.Col = C_ProdQty2		' Produced Qty
			
			If UNICDbl(.Text) > CDbl(0) and strFlag = ggoSpread.InsertFlag Then

				strVal = strVal & frm1.hPlantCd.value & parent.gColSep
				.Col = C_ProdtOrderNo2			
				strVal = strVal & UCase(Trim(.Text)) & parent.gColSep
				.Col = C_OprNo2
				strVal = strVal & UCase(Trim(.Text)) & parent.gColSep
				.Col = C_ProdQty2
				strVal = strVal & UniConvNum(.Text,0) & parent.gRowSep
			End If
		Next
	End With

	iCalledAspName = AskPRAspName("P4400RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4400RA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, strVal), _
		"dialogWidth=720px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
End Function

'===============================================================================
' Function Name : JumpReworkRun
' Function Desc : Jump to (E)Production Order(Single)
'========================================================================================
Function JumpReworkRun()
	
	Dim strProdtOrdNo, strOprNo
	Dim strItemCd
	Dim DblJumpQty, DblInspBadQty, DblBadQty, DblReworkQty
	Dim strTrackingNo
	
	If lgIntFlgMode = parent.OPMD_CMODE Then		
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	If Trim(frm1.txtPlantCd.value) = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
	With frm1.vspdData1 
		.Row = .ActiveRow
		.Col = C_InspBadQtyInOrderUnit
		DblInspBadQty = UNICDbl(.Text)
		.Col = C_BadQtyInOrderUnit	
		DblBadQty = UNICDbl(.Text)
		.Col = C_ReworkPrevQty	
		DblReworkQty = UNICDbl(.Text)
		
		DblJumpQty = DblInspBadQty + DblBadQty - DblReworkQty
		'Error Check -  Whether Defect Qty is greater than zero
		
		If DblInspBadQty + DblBadQty = Cdbl(0) Then
			Call DisplayMsgBox("189247", "x", "x", "x")
			Exit Function 
		End If
		
		If DblJumpQty <= 0 Then
			Call DisplayMsgBox("189248", "x", "x", "x")
			Exit Function 
		End If
		
		.Col = C_ProdtOrderNo
		strProdtOrdNo = UCase(Trim(.Text))
		.Col = C_OprNo
		strOprNo = UCase(Trim(.Text))
		.Col = C_ItemCd
		strItemCd = UCase(Trim(.Text))
		.Col = C_TrackingNo
		strTrackingNo = UCase(Trim(.Text))
		
	End With
	
End Function

'========================================================================================
' Function Name : JumpOrdRscComptRun
' Function Desc : Jump to (E)Production Order(Single)
'========================================================================================
Function JumpOrdRscComptRun()
	
	Dim strProdtOrdNo, strOprNo
	
	If lgIntFlgMode = parent.OPMD_CMODE Then		
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
	With frm1.vspdData1 
		.Row = .ActiveRow
		.Col = C_ProdtOrderNo
		strProdtOrdNo = UCase(Trim(.Text))
'		.Col = C_OprNo
'		strOprNo = UCase(Trim(.Text))
		
	End With	
		
	WriteCookie "txtPlantCd", UCase(Trim(frm1.hPlantCd.value))
	WriteCookie "txtPlantNm", UCase(Trim(frm1.txtPlantNm.value))
	WriteCookie "txtProdOrderNo", strProdtOrdNo
'	WriteCookie "txtOprNo", strOprNo
	
	PgmJump(BIZ_PGM_JUMPORDRSCCOMPT_ID)
	
End Function


'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp에서 전송된 값을 특정 Tag Object에 지정 
'=========================================================================================================
'++++++++++++++++++++  Insert Your Code for PopUp(Open)  +++++++++++++++++++++++++++++++++++++++++++++++++

'------------------------------------------  SetPlant()  -------------------------------------------------
'	Name : SetPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetProdOrderNo()  -------------------------------------------
'	Name : SetProdOrderNo()
'	Description : Production Order Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetProdOrderNo(byval arrRet)
	frm1.txtProdOrderNo.Value    = arrRet(0)		
End Function

'------------------------------------------  SetItemInfo()  ----------------------------------------------
'	Name : SetItemInfo()
'	Description : Item Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetItemInfo(Byval arrRet)

    With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
    End With

End Function

'------------------------------------------  SetItemInfo()  -------------------------------------------
'	Name : SetItemGroup()
'	Description : Item Group Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function

'------------------------------------------  SetWcCd()  -------------------------------------------------
'	Name : SetWcCd()
'	Description : Work Center Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetWcCd(byval arrRet)
	frm1.txtWCCd.Value    = arrRet(0)		
	frm1.txtWCNm.Value    = arrRet(1)		
End Function

'------------------------------------------  SetTrackingNo()  ----------------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup에서 Return되는 값 setting
'-----------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
	
	frm1.txtTrackingNo.Value = arrRet(0)
	
End Function

Sub cboLine_onclick()

	if frm1.txtPlantCd.value = "" then
	 Call DisplayMsgBox("169901","X", "공장","X")
	 exit sub
	end if 
    
End Sub


'------------------------------------------  txtPlantCd_OnChange -----------------------------------------
'	Name : txtPlantCd_OnChange()
'	Description : Plant Popup에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Sub txtPlantCd_OnChange()
	Call initCombobox2()
	Call InitComboBox
End Sub

'------------------------------------------  txtProdFromDt_KeyDown ----------------------------------------
'	Name : txtProdFromDt_KeyDown
'	Description : Plant Popup에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Sub txtProdFromDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'------------------------------------------  txtProdToDt_KeyDown ------------------------------------------
'	Name : txtProdToDt_KeyDown
'	Description : Plant Popup에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------
Sub txtProdToDt_KeyDown(keycode, shift)
	If Keycode = 13 Then
		Call MainQuery()
	End If
End Sub	

'=======================================================================================================
'   Event Name : txtPlanStartDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtProdFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtProdFromDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtProdFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtPlanStartDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtProdToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtProdToDt.Action = 7
        Call SetFocusToDocument("M")
		Frm1.txtProdToDt.Focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtReportDT_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
'Sub txtReportDT_DblClick(Button)
'    If Button = 1 Then
'        frm1.txtReportDT.Action = 7
'        Call SetFocusToDocument("M")
'		Frm1.txtReportDT.Focus
'    End If
'End Sub
'------------------------------------------  txtReportDT_OnChange -----------------------------------------
'	Name : txtReportDT_OnChange()
'	Description : vspddata2의 ReportDt 업데이트 
'----------------------------------------------------------------------------------------------------------
'Sub txtReportDT_Change()
'	dim intRows
'	if frm1.txtReportDt.text = "" then
'	else
'		with frm1.vspdData2
'		for intRows = 1 to .maxRows    
'			.Row = intRows
'			.Col = 0
'			if .text = ggoSpread.InsertFlag then
'				.Col = C_ReportDt
'				.text = frm1.txtReportDt.text
'			End if
'		next
'		end with
'		with frm1.vspdData3
'		for intRows = 1 to .maxRows
'			.Row = intRows
'			.Col = 0
'			if .text = ggoSpread.InsertFlag then
'				.Col = C_ReportDt2
'				.text = frm1.txtReportDt.text
'			End if
'		next 
'		end with
'	
'	End if
'End Sub
'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'========================================================================================
' Function Name : JumpReworkRun
' Function Desc : Jump to (E)Production Order(Single)
'========================================================================================
Function JumpReworkRun()
	
	Dim strProdtOrdNo, strOprNo
	Dim strItemCd
	Dim DblJumpQty, DblInspBadQty, DblBadQty, DblReworkQty
	Dim strTrackingNo
	
	If lgIntFlgMode = parent.OPMD_CMODE Then		
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False
		Exit Function
	End If
	
	With frm1.vspdData1 
		.Row = .ActiveRow
'		.Col = C_InspBadQtyInOrderUnit
'		DblInspBadQty = UNICDbl(.Text)
'		.Col = C_BadQtyInOrderUnit	
'		DblBadQty = UNICDbl(.Text)
'		.Col = C_ReworkPrevQty	
'		DblReworkQty = UNICDbl(.Text)
		
'		DblJumpQty = DblInspBadQty + DblBadQty - DblReworkQty
		'Error Check -  Whether Defect Qty is greater than zero
		
'		If DblInspBadQty + DblBadQty = Cdbl(0) Then
'			Call DisplayMsgBox("189247", "x", "x", "x")
'			Exit Function 
'		End If
		
'		If DblJumpQty <= 0 Then
'			Call DisplayMsgBox("189248", "x", "x", "x")
'			Exit Function 
'		End If
		
'		.Col = C_ProdtOrderNo
'		strProdtOrdNo = UCase(Trim(.Text))
'		.Col = C_OprNo
'		strOprNo = UCase(Trim(.Text))
'		.Col = C_ItemCd
'		strItemCd = UCase(Trim(.Text))
'		.Col = C_TrackingNo
'		strTrackingNo = UCase(Trim(.Text))
		
	End With	
		
	WriteCookie "txtPlantCd", UCase(Trim(frm1.hPlantCd.value))
	WriteCookie "txtPlantNm", UCase(Trim(frm1.txtPlantNm.value))
	WriteCookie "txtItemCd", strItemCd
	WriteCookie "txtProdOrderNo", strProdtOrdNo
	WriteCookie "txtOprNo", strOprNo
	WriteCookie "txtJumpQty", UniFormatNumber(DblJumpQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)
	WriteCookie "txtTrackingNo", strTrackingNo
	
	PgmJump(BIZ_PGM_JUMPREWORKRUN_ID)
	
End Function


'#########################################################################################################
'												3. Event부 
'	기능: Event 함수에 관한 처리 
'	설명: Window처리, Single처리, Grid처리 작업.
'         여기서 Validation Check, Calcuration 작업이 가능한 Event가 발생.
'         각 Object단위로 Grouping한다.
'##########################################################################################################
'******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'**********************************************************************************************************

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'**************************  3.2 HTML Form Element & Object Event처리  *********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'*******************************************************************************************************

'******************************  3.2.1 Object Tag 처리  ************************************************
'	Window에 발생 하는 모든 Even 처리	
'*******************************************************************************************************

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
'Sub vspdData1_Click(ByVal Col , ByVal Row )
    '---------------------- 
	'Column Split
	'----------------------
'	gMouseClickStatus = "SPC"
	
'	Set gActiveSpdSheet = frm1.vspdData1
    
'  	Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
  	
' 	If frm1.vspdData1.MaxRows = 0 Then
' 		Exit Sub
' 	End If
 	
' 	If Row <= 0 Then
' 		ggoSpread.Source = frm1.vspdData1 
' 		If lgSortKey1 = 1 Then
' 			ggoSpread.SSSort Col					'Sort in Ascending
' 			lgSortKey1 = 2
' 		Else
' 			ggoSpread.SSSort Col, lgSortKey1		'Sort in Descending
' 			lgSortKey1 = 1
' 		End If
 		
' 		frm1.vspdData1.MaxRows = 0
			
'		If DbDtlQuery = False Then
'			Call RestoreToolBar()
'			Exit Sub
'		End If
 		
'	Else
 		'------ Developer Coding part (Start)
' 		If lgOldRow <> Row Then
	
'			frm1.vspdData1.Col = 1
'			frm1.vspdData1.Row = row
			
'			lgOldRow = Row

'			frm1.vspdData2.MaxRows = 0
			
'			If DbDtlQuery = False Then	
'				Call RestoreToolBar()
'				Exit Sub
'			End If	

'		End If
	 	'------ Developer Coding part (End)
	
' 	End If
 	
	
'End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Cell을 벗어날시 무조건발생 이벤트 
'==========================================================================================
Sub vspdData1_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
    Dim strInsideFlag
    
    If Row = NewRow Then
        Exit Sub
    End If

	If NewRow <= 0 Or NewCol < 0 Then
		Exit Sub
	End If

	frm1.vspdData1.Row = NewRow
	
'	If SetBlnInsertRow(frm1.vspdData1.Row) = True Then
'		Call SetToolBar("11001101000111")										'⊙: 버튼 툴바 제어 
'	ElseIf SetBlnInsertRow(frm1.vspdData1.Row) = False Then
'		Call SetToolBar("11001000000111")										'⊙: 버튼 툴바 제어 
'	End If
    
End Sub

'==========================================================================================
'   Event Name : vspdData1_Click
'   Event Desc :
'==========================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP2C"
	
	Set gActiveSpdSheet = frm1.vspdData1
	
'	If SetBlnInsertRow(frm1.vspdData1.ActiveRow) = True Then
		Call SetPopupMenuItemInf("1101101100")         '화면별 설정 
'	ElseIf SetBlnInsertRow(frm1.vspdData1.ActiveRow) = False Then	
'		Call SetPopupMenuItemInf("0000111111")         '화면별 설정 
'    End If
    
    
    
 	If frm1.vspdData1.MaxRows = 0 Then
 		Exit Sub
 	End If
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData1 
 		If lgSortKey2 = 1 Then
 			ggoSpread.SSSort Col					'Sort in Ascending
 			lgSortKey2 = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey2		'Sort in Descending
 			lgSortKey2 = 1
 		End If
	Else
 			
 	End If

End Sub

'========================================================================================================
'   Event Name : vspdData1_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData1_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData2_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
End Sub

'==========================================================================================
'   Event Name : vspdData1_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

'Sub vspdData1_MouseDown(Button,Shift,x,y)
		
'	If Button <> "1" And gMouseClickStatus = "SPC" Then
'		gMouseClickStatus = "SPCR"
'	End If

'End Sub

'==========================================================================================
'   Event Name : vspdData2_MouseDown(Button,Shift,x,y)
'   Event Desc :
'==========================================================================================

Sub vspdData1_MouseDown(Button,Shift,x,y)
		
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub

'=======================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData1_Change(ByVal Col, ByVal Row)

	Dim strItemCd
	Dim strHndItemCd, strHndOprNo
	Dim i
	Dim strReqDt, strEndDt
	Dim strProdtOrderNo, strOprNo
	Dim LngFindRow
	Dim lRow
	Dim PQtyU, PQtyD 
	Dim ProdtOrderNo1
	Dim ProdtOrderQty1
	Dim MinorCd, MinorNm, Spec1, strWhere
	Dim JobLineCd1, JobQtySum1
    
    PQtyU = 0
	PQtyD = 0	

	ggoSpread.Source = frm1.vspdData1
	ggoSpread.UpdateRow Row

	Frm1.vspdData1.Row = Row
	Frm1.vspdData1.Col = Col
	
	Call CheckMinNumSpread(frm1.vspdData1, Col, Row)        '  <------변경된 표준 라인 
	
	With frm1.vspdData1
	
		.Row = Row
		.Col = Col

		Select Case Col
		
			Case C_ProdtOrderNo						   	

				.col = C_ProdtOrderNo	
                 ProdtOrderNo1 = Trim(.text)
                 
				if ProdtOrderNo1 = "" then
					.Col = C_ProdtOrderQty
					.value = cdbl(0)
					
					.Col = C_ProdtOrderLineSumQty
					.value = cdbl(0)
					
					.Col = C_ItemCd
					.value = ""
						
					.Col = C_ItemNm
					.value = ""
						
					.Col = C_Spec
					.value = ""
					
     			else
     			    If frm1.txtPlantCd.value= "" Then
						Call displaymsgbox("971012","X", "공장","X")
						.Col = C_ProdtOrderNo
						.Value = ""			
						frm1.txtPlantCd.focus 
						Set gActiveElement = document.activeElement
						Exit Sub
					End If
	
					If frm1.txtProdFromDt.Text= "" Then
						Call displaymsgbox("971012","X", "계획일","X")
						.Col = C_ProdtOrderNo
						.Value = ""			
						frm1.txtProdFromDt.focus 
						Set gActiveElement = document.activeElement
						Exit Sub
					End If
     			    
'     			    .col = C_JobLineCd
'    			    .Row = row	
'					JobLineCd1 =  .Text

'					if JobLineCd1 = "" then
'						Call displaymsgbox("971012","X", "Line","X")
'						.Col = C_ProdtOrderNo
'						.Value = ""						
'						.focus 
'						Set gActiveElement = document.activeElement
'						Exit Sub					    	
'					end if
					
					strWhere    = ""
					strWhere	=	strWhere & " a.item_cd = b.item_cd "						    		       		    
					strWhere	=	strWhere & " and a.plant_cd = " & FilterVar(trim(frm1.txtplantcd.value), "''", "S")
					strWhere	=	strWhere & " and a.plan_start_dt >=  '" & uniconvdate(FilterVar(frm1.txtProdFromDt.text,"","SNM")) & "'"
					strWhere	=	strWhere & " and a.prodt_order_no  = '" & ProdtOrderNo1 & "'"

					IF  CommonQueryRs(" a.item_cd, b.item_nm, b.spec ", " p_production_order_header a (nolock), b_item b ", strWhere , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) THEN    
						MinorCd    = Replace(lgF0,Chr(11),"")
						MinorNm    = Replace(lgF1,Chr(11),"")
						Spec1      = Replace(lgF2,Chr(11),"")
								
						.Col = C_ItemCd
						.value = MinorCd
						
						.Col = C_ItemNm
						.value = MinorNm
						
						.Col = C_Spec
						.value = Spec1
																																			
					ELSE
						Call DisplayMsgBox("XX1129","X", MinorCd,"X")
						.Col = C_ProdtOrderNo
						.text = ""
						
						.Col = C_ItemCd
						.value = ""
						
						.Col = C_ItemNm
						.value = ""
						
						.Col = C_Spec
						.value = ""
																					
					END IF										
					
     				strWhere    = ""	    		       		    
					strWhere	=	strWhere & " plant_cd = " & FilterVar(trim(frm1.txtplantcd.value), "''", "S")
					strWhere	=	strWhere & " and prodt_order_no = '" & FilterVar(ProdtOrderNo1,"","SNM") & "'"

					IF  CommonQueryRs(" prodt_order_no,  isnull(prodt_order_qty,0) prodt_order_qty ", " p_production_order_header (nolock) ", strWhere , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) THEN    
						MinorCd    = Replace(lgF0,Chr(11),"")
						ProdtOrderQty1 = cdbl(Replace(lgF1,Chr(11),""))
								
						.Col = C_ProdtOrderQty
						.value = ProdtOrderQty1																													
					ELSE
						Call DisplayMsgBox("XX1129","X", MinorCd,"X")
						.Col = C_ProdtOrderNo
						.text = ""									
					END IF
					
'					strWhere    = ""	    		       		    
'					strWhere	=	strWhere & " plant_cd = " & FilterVar(trim(frm1.txtplantcd.value), "''", "S")
'					strWhere	=	strWhere & " and prodt_order_no = '" & FilterVar(ProdtOrderNo1,"","SNM") & "'"
'					strWhere	=	strWhere & " and job_plan_dt = '" & Uniconvdate(Trim(frm1.txtProdFromDt.text)) & "'"
'					strWhere	=	strWhere & " and job_line = '" & JobLineCd1 & "'"

'					IF  CommonQueryRs(" isnull(sum(job_qty),0) as Prodt_Order_SumQty ", " p_prod_time_order_ko119 (nolock) ", strWhere , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) THEN    
'						JobQtySum1    = cdbl(Replace(lgF0,Chr(11),""))
								
'						.Col = C_ProdtOrderLineSumQty
'						.value = JobQtySum1
						
'						if JobQtySum1 = 0 then
'							For i = 1  to C_GRIDCOUNT 
'								iDx = Cstr(((int(i)*5)+9) -2)
'								.Col = iDx 
'								.Row = Row 
'								PQtyD = PQtyD + cdbl(Trim(.value))
'							Next
						
'							.Col = C_ProdtOrderLineSumQty
'							.value = PQtyD
'						end if																													
'					ELSE
'						For i = 1  to C_GRIDCOUNT 
'							iDx = Cstr(((int(i)*5)+9) -2)
'							.Col = iDx 
'							.Row = Row 
'							PQtyD = PQtyD + cdbl(Trim(.value))
'						Next
						
'						.Col = C_ProdtOrderLineSumQty
'						.value = PQtyD
'					END IF																																							
     			end if	     	 			                  			

		    Case "12", "13", "17", "18", "22", "23", "27", "28",  "32", "33",  "37", "38", "42", "43", "47", "48", "52", "53", "57", "58", "62", "63", "67", "68", "72", "73", "77", "78", "82", "83", "87", "88", "92", "93", "97", "98", "102", "103", "107", "108", "112", "113", "117", "118", "122", "123", "127", "128"     				                                                                                                	
				
				Call NumberCheck(Row, Col)
				
				.Row =  Row 
				.Col =  C_ProdtOrderNo
				strProdtOrderNo = .value
				   
				.Col = C_ProdtOrderQty
    			PQtyU	= cdbl(Trim(.value))
    			
				For i = 1  to C_GRIDCOUNT 
				  iDx = Cstr(((int(i)*5)+9) -2)
				  .Col = iDx 
				  .Row = Row 
				  PQtyD = PQtyD + cdbl(Trim(.value))
				Next
				
				.Col = C_ProdtOrderLineSumQty
				.value = PQtyD																														    
			Case "14", "19", "24", "29", "34", "39", "44", "49", "54", "59", "64", "69", "74", "79", "84", "89", "94", "99", "104", "109", "114", "119", "124", "129"			     

'				.Col = 0
'				.Row = Row
'				if frm1.vspdData1.Text = ggoSpread.InsertFlag or frm1.vspdData1.Text = ggoSpread.DeletetFlag then
'				else
'					ggoSpread.UpdateRow Row
'				end if
		End Select

	End With

End Sub

'==========================================================================================
'   Event Name :vspddata1_ComboSelChange                                                          
'   Event Desc :Combo Change Event                                                                           
'==========================================================================================
Sub vspdData1_ComboSelChange(ByVal Col, ByVal Row)

	Dim intIndex
	Dim strLotReq
	Dim strAutoRcptFlg
	Dim strRoutOrder
	Dim strProdtOrderNo, strOprNo
	Dim LngFindRow
	Dim ProdtOrderNo1				    				

	With frm1.vspdData1

		.Row = Row
		Select Case Col
		
			Case  C_JobLineCd
				.Col = Col
				intIndex = .Value
				.Col = C_JobLine
				.Value = intIndex
			Case  C_JobLine
				.Col = Col
				intIndex = .Value
				.Col = C_JobLineCd
				.Value = intIndex				
		End Select
		
		.Col = C_ProdtOrderNo
		ProdtOrderNo1 = .value
		
		if ProdtOrderNo1 <> "" then
			Call vspdData1_Change(C_ProdtOrderNo,frm1.vspdData1.ActiveRow)
		end if
		
    End With

End Sub

'==========================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc : check button clicked
'==========================================================================================
Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, ByVal ButtonDown)

	With frm1.vspdData1 
		ggoSpread.Source = frm1.vspdData1
		If Row > 0 And Col = C_ProdtOrderPopup Then
			.Col = C_ProdtOrderNo
			.Row = Row
			Call OpenProdOrderNo1(.text)
		Else
		End if 
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub

Sub vspdData2_GotFocus()
    ggoSpread.Source = frm1.vspdData2
End Sub

'==========================================================================================
'   Event Name : vspdData1_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
             Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey <> "" and lgStrPrevKey1 <> "" and lgStrPrevKey2 <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'==========================================================================================
'   Event Name : vspdData2_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then			'⊙: 다른 비즈니스 로직이 진행 중이면 더 이상 업무로직ASP를 호출하지 않음 
             Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
'    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
'		If lgIntPrevKey <> 0 Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
'			Call LayerShowHide(1)
'			If DbDtlQuery = False Then	
'				Call RestoreToolBar()
'				Exit Sub
'			End If
'		End If     
'   End if
    
End Sub

'#########################################################################################################
'												4. Common Function부 
'	기능: Common Function
'	설명: 환율처리함수, VAT 처리 함수 
'#########################################################################################################


'#########################################################################################################
'												5. Interface부 
'	기능: Interface
'	설명: 각각의 Toolbar에 대한 처리를 행한다. 
'	      Toolbar의 위치순서대로 기술하는 것으로 한다. 
'	<< 공통변수 정의 부분 >>
' 	공통변수 : Global Variables는 아니지만 각각의 Sub나 Function에서 자주 사용하는 변수로 변수명은 
'				통일하도록 한다.
' 	1. 공통컨트롤을 Call하는 변수 
'    	   ADF (ADS, ADC, ADF는 그대로 사용)
'    	   - ADF는 Set하고 사용한 뒤 바로 Nothing 하도록 한다.
' 	2. 공통컨트롤에서 Return된 값을 받는 변수 
'    		strRetMsg
'#########################################################################################################
'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 

    Dim IntRetCD 
    
    FncQuery = False											'⊙: Processing is NG
    
    Err.Clear													'☜: Protect system from crashing

    ggoSpread.Source = frm1.vspdData1							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
        IntRetCD = displaymsgbox("900013", parent.VB_YES_NO, "x", "x")	'⊙: Display Message
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
'    If ValidDateCheck(frm1.txtProdFromDt, frm1.txtProdTODt) = False Then Exit Function
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						'⊙: Clear Contents  Field

    Call InitVariables
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData


	'-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then	
		Call RestoreToolBar()
		Exit Function
	End If														'☜: Query db data
	
    FncQuery = True												'⊙: Processing is OK
   
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	On Error Resume Next
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	On Error Resume Next    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 

    Dim IntRetCD 
    Dim	LngRows
    Dim lRow, IRow2
    Dim PQtyU, PQtyD, Flag, iDx, i
    Dim strProdtOrderNo
    Dim strProdtOrderNo1
	
	PQtyU = 0
	PQtyD = 0
    
    FncSave = False												'⊙: Processing is NG
    
    Err.Clear                                                   '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData1							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = False Then						'⊙: Check If data is chaged
        IntRetCD = displaymsgbox("900001", "x", "x", "x")		'⊙: Display Message(There is no changed data.)
        Exit Function
    End If
    
    If Not ggoSpread.SSDefaultCheck Then						'⊙: Check required field(Multi area)
       Exit Function
    End If
    If Not chkfield(Document, "1") Then					'⊙: Check required field(Single area)
       Exit Function
    End If

   	With frm1.vspddata1
   	    For lRow = 1 To .MaxRows
   			.Row = lRow
   			.Col =  C_ProdtOrderNo
			strProdtOrderNo = .value
			.Col = C_ProdtOrderQty
   			PQtyU	= cdbl(.value)
   	      
   			PQtyD = 0
            For IRow2 = 1 to .MaxRows
                .Row = IRow2
   				.Col =  C_ProdtOrderNo
				strProdtOrderNo1 = .value
				
				if strProdtOrderNo = strProdtOrderNo1 then    	      
   						For i = 1  to C_GRIDCOUNT 
							iDx = Cstr(((int(i)*5)+9) -2)
							.Col = iDx 
							.Row = IRow2 
							PQtyD = PQtyD + cdbl(Trim(.value))
						Next
				else		
				end if	
			Next

			If cdbl(PQtyU) >=  cdbl(PQtyD)  then
			Else
				Call DIsplayMsgBox("XX1010", vbOKOnly, "x", "x")
				Exit Function
			End if	
			
		Next	
		    
	End With    
          		   			            
    '-----------------------
    'Save function call area
    '-----------------------
    If DbSave = False Then Exit Function						'☜: Save db data
     
    FncSave = True												'⊙: Processing is OK

End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
        
	If frm1.vspdData2.MaxRows < 1 Then Exit Function	
        
    frm1.vspdData2.focus
    Set gActiveElement = document.activeElement 
	frm1.vspdData2.EditMode = True
	frm1.vspdData2.ReDraw = False
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.CopyRow
'    ggoSpread.Source = frm1.vspdData3
'    ggoSpread.CopyRow
    frm1.vspdData2.ReDraw = True
    
    SetSpreadColor frm1.vspdData2.ActiveRow, frm1.vspdData2.ActiveRow
   
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 

Dim Row
Dim strMode
Dim strChangeFlag
Dim LngRow, LngFindRow

	If frm1.vspdData1.MaxRows < 1 Then Exit Function	

    ggoSpread.Source = frm1.vspdData1	
    Row = frm1.vspdData1.ActiveRow
'    frm1.vspdData1.Row = Row
'    frm1.vspdData1.Col = 0
'    strMode = frm1.vspdData1.Text

     ggoSpread.EditUndo                                                  '☜: Protect system from crashing

	strChangeFlag = "N"
	
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt) 

	Dim IntRetCD
	Dim imRow
	Dim pvRow
	Dim strProdtOrderNo
	Dim strSelect
	Dim strFrom
	Dim strWhere
	Dim arrVal1, arrVal2
	Dim strMinorCd, strMinorNm
	Dim ii, jj, iii, i
	Dim Time1, Time2, Time3, Time4, Time5, Time6, Time7, Time8, Time9, Time10, Time11, Time12
	Dim Time13, Time14, Time15, Time16, Time17, Time18, Time19, Time20, Time21, Time22, Time23, Time24 

'	On Error Resume Next
	
	FncInsertRow = False

'	If SetBlnInsertRow(frm1.vspdData1.ActiveRow) = False Then Exit Function

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()

		If imRow = "" Then
			Exit Function
		End If
	End If

		  	strSelect	=			 " a.minor_cd, a.minor_nm "
			strFrom		=			 " b_minor a (NOLOCK), b_configuration b (nolock) "
			strWhere	=			 " a.major_cd = b.major_cd and a.major_cd = 'M2110' and b.seq_no = 99 and a.minor_cd = b.minor_cd "
			strWhere	= strWhere & " order by b.reference "

			If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) Then 

				arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))
				jj = Ubound(arrVal1,1) 

				For ii = 0 To jj - 1 
					i = ii + 1
					arrVal2			= Split(arrVal1(ii), chr(11))
					strMinorCd		= Ucase(Trim(arrVal2(1)))
					strMinorNm		= Trim(arrVal2(2))

				select case i 
					case "1"
						Time1 = strMinorCd
					case "2"
						Time2 = strMinorCd
					case "3"
						Time3 = strMinorCd
					case "4"
						Time4 = strMinorCd
					case "5"
						Time5 = strMinorCd
					case "6"
						Time6 = strMinorCd
					case "7"
						Time7 = strMinorCd
					case "8"
						Time8 = strMinorCd
					case "9"
						Time9 = strMinorCd
					case "10"
						Time10 = strMinorCd
					case "11"
						Time11 = strMinorCd
					case "12"
						Time12 = strMinorCd
					case "13"
						Time13 = strMinorCd
					case "14"
						Time14 = strMinorCd
					case "15"
						Time15 = strMinorCd
					case "16"
						Time16 = strMinorCd
					case "17"
						Time17 = strMinorCd
					case "18"
						Time18 = strMinorCd
					case "19"
						Time19 = strMinorCd
					case "20"
						Time20 = strMinorCd
					case "21"
						Time21 = strMinorCd
					case "22"
						Time22 = strMinorCd
					case "23"
						Time23 = strMinorCd
					case "24"
						Time24 = strMinorCd
				End Select		
				Next	
         End if		

	With frm1
		.vspdData1.Row = .vspdData1.ActiveRow	    
		' Get Production Order No.
'		.vspdData1.Col = C_ProdtOrderNo
'		strProdtOrderNo = Trim(.vspdData1.value)
'		.vspdData2.focus
'		Set gActiveElement = document.activeElement 
		ggoSpread.Source = .vspdData1
		
'		For LngRowCnt = 1 To .vspdData2.MaxRows
'			.vspdData2.Row = LngRowCnt
			
'		    .vspdData2.Col = C_Sequence
			
'			If CInt(LngCompSeq) < CInt(.vspdData2.value) Then
'				LngCompSeq = CInt(.vspdData2.value)
'			End If
'		Next
		
		.vspdData1.ReDraw = False
		ggoSpread.InsertRow .vspdData1.ActiveRow, imRow
    	
    	For pvRow = .vspdData1.ActiveRow To .vspdData1.ActiveRow + imRow -1
			
			.vspdData1.Row = pvRow
'			.vspdData2.Col = C_ProdtOrderNo1
'			.vspdData2.value = strProdtOrderNo
           
           For iii = 1  to C_GRIDCOUNT * 5 
             i2 = iii + 9
'            if i2 = "10" or i2 = "15" or i2 = "20" or i2 = "25" or i2 = "30" or i2 = "35" or i2 = "40" or i2 = "45" or i2 = "50" or i2 = "55" or i2 = "60" or i2 = "65" or i2 = "70" or i2 = "75" or i2 = "80" or i2 = "85" or i2 = "90" or i2 = "95" or i2 = "100" or i2 = "105" or i2 = "110" or i2 = "115" or i2 = "120" or i2 = "125" then             
				.vspdData1.Col = i2
			   select case i2
			     case "10"	
					.vspdData1.value = Time1
				 case "13"	
					.vspdData1.value = "0"
				 case "15"	
					.vspdData1.value = Time2
				 case "18"
					.vspdData1.value = "0"	
				 case "20"	
					.vspdData1.value = Time3	
				 case "23"
					.vspdData1.value = "0"	
				 case "25"	
					.vspdData1.value = Time4	
				 case "28"
					.vspdData1.value = "0"
				 case "30"	
					.vspdData1.value = Time5	
				 case "33"
					.vspdData1.value = "0"	
				 case "35"	
					.vspdData1.value = Time6
				 case "38"
					.vspdData1.value = "0"	
				 case "40"	
					.vspdData1.value = Time7		
				 case "43"
					.vspdData1.value = "0"	
				 case "45"	
					.vspdData1.value = Time8
				 case "48"
				    .vspdData1.value = "0"	
				 case "50"	
					.vspdData1.value = Time9
				 case "53"
				    .vspdData1.value = "0"		
				 case "55"	
					.vspdData1.value = Time10		
				 case "58"
					.vspdData1.value = "0"	
				 case "60"	
					.vspdData1.value = Time11		
				 case "63"
					.vspdData1.value = "0"	
				 case "65"	
					.vspdData1.value = Time12
				 case "68"
					.vspdData1.value = "0"		
				 case "70"	
					.vspdData1.value = Time13
				 case "73"
					.vspdData1.value = "0"	
				 case "75"	
					.vspdData1.value = Time14
				 case "78"
					.vspdData1.value = "0"		
				 case "80"	
					.vspdData1.value = Time15
				 case "83"
					.vspdData1.value = "0"		
				 case "85"	
					.vspdData1.value = Time16
				 case "88"
					.vspdData1.value = "0"		
				 case "90"	
					.vspdData1.value = Time17
				 case "93"
					.vspdData1.value = "0"	
				 case "95"	
					.vspdData1.value = Time18
				 case "98"
					.vspdData1.value = "0"	
				 case "100"	
					.vspdData1.value = Time19
				 case "103"
					.vspdData1.value = "0"	
				 case "105"	
					.vspdData1.value = Time20
				 case "108"
					.vspdData1.value = "0"					
				 case "110"	
					.vspdData1.value = Time21
				 case "113"
					.vspdData1.value = "0"	
				 case "115"	
					.vspdData1.value = Time22
				 case "118"
					.vspdData1.value = "0"			
				 case "120"	
					.vspdData1.value = Time23
				 case "123"	
					.vspdData1.value = "0"	
				 case "125"	
					.vspdData1.value = Time24
				 case "128"
					.vspdData1.value = "0"	
				End select
'			end if
		   Next	
		Next
		
		.vspdData1.ReDraw = True
		
		SetSpreadColor frm1.vspdData1.ActiveRow, frm1.vspdData1.ActiveRow + imRow -1
				
		Set gActiveElement = document.ActiveElement
		
		If Err.number = 0 Then FncInsertRow = True
		
	End With
    
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
function FncDeleteRow() 
    Dim lDelRows
    
    if frm1.vspdData1.maxrows < 1 then exit function 
	   
    
    With frm1.vspdData1
    	.focus
    	ggoSpread.Source = frm1.vspdData1 
    	lDelRows = ggoSpread.DeleteRow
    End With
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.fncPrint()
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLEMULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()

	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData1							'⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then						'⊙: Check If data is chaged
		IntRetCD = displaymsgbox("900016", parent.VB_YES_NO, "x", "x")	'⊙: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    FncExit = True
    
End Function

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  ******************************
'	설명 : 
'*********************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

Dim strVal
Dim lngRows
Dim strPlantCd

	Call MakeKeyStream(pDirect)

	DbQuery = False
	
    With frm1
       
'        frm1.vspdData1.MaxRows = 0

		strPlantCd = frm1.txtplantcd.value	
		Call initspreadsheet3(strPlantCd)

		Call LayerShowHide(1)       

		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001						'☜: 
			strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'☆: 조회 조건 데이타 
    		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
    		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
    		strVal = strVal & "&cboLine=" & Trim(.hcboLine.Value)
			strVal = strVal & "&txtProdOrderNo=" & Trim(.hProdOrderNo.value)
			strVal = strVal & "&txtItemCd=" & Trim(.htxtItemCd.Value)
			strVal = strVal & "&txtProdFromDt=" & Trim(.hProdFromDt.Value)
			strVal = strVal & "&txtMaxCount=" & Trim(MaxCount)
			strVal = strVal & "&GridColCount=" & Trim(C_GRIDCOUNT)
			strVal = strVal & "&txtKeyStream=" & lgKeyStream2         '☜: Query Key
			strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
		Else
			strVal = BIZ_PGM_QRY1_ID & "?txtMode=" & parent.UID_M0001						'☜: 
			strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
    		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
			strVal = strVal & "&cboLine=" & Trim(.cboLine.value)
			strVal = strVal & "&txtProdOrderNo=" & Trim(.txtProdOrderNo.value)
			strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.Value)
			strVal = strVal & "&txtProdFromDt=" & Trim(.txtProdFromDt.Text)
			strVal = strVal & "&GridColCount=" & Trim(C_GRIDCOUNT)
			strVal = strVal & "&txtMaxCount=" & Trim(MaxCount)
			strVal = strVal & "&txtKeyStream=" & lgKeyStream2         '☜: Query Key
			strVal = strVal & "&txtMaxRows=" & .vspdData1.MaxRows
		End If

		Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

    End With

    DbQuery = True

End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk(ByVal LngMaxRow)

	Dim strInsideFlag

'	Call InitData()
	
'	frm1.vspdData1.Col = 1
'	frm1.vspdData1.Row = 1

'	lgOldRow = 1	

'	If lgIntFlgMode <> parent.OPMD_UMODE Then
'		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
'		Set gActiveElement = document.activeElement

'		If DbDtlQuery = False Then	
'			Call RestoreToolBar()
'			Exit Function
'		End If
'	End If
	Call InitSpreadComboBox()	
	Call InitData()
	Call SetToolbar("110011110001111")
	
	lgIntFlgMode = parent.OPMD_UMODE
	
	Call SetSpreadlock2

End Function

'========================================================================================
' Function Name : DbQueryNotOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryNotOk()

'	Call SetToolBar("11000000000011")										'⊙: 버튼 툴바 제어 
'	Call SetFieldColor(False)
    
End Function

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery() 

Dim strVal
Dim boolExist
Dim lngRows

	Call MakeKeyStream(pDirect)

	boolExist = False
    With frm1

       
        frm1.vspdData2.MaxRows = 0
        
		Call initspreadsheet3(strProdtOrderNo)
		
		DbDtlQuery = False   
    
		.vspdData1.Row = .vspdData1.ActiveRow

		Call LayerShowHide(1)       

		If lgIntFlgMode = parent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'☜: 
			strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&lgIntPrevKey=" & lgIntPrevKey
			strVal = strVal & "&txtProdOrderNo=" & Trim(strProdtOrderNo)
			strVal = strVal & "&txtMaxCount=" & Trim(MaxCount)
			strVal = strVal & "&GridColCount=" & Trim(C_GRIDCOUNT)
			strVal = strVal & "&txtKeyStream=" & lgKeyStream2         '☜: Query Key
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		Else
			strVal = BIZ_PGM_QRY2_ID & "?txtMode=" & parent.UID_M0001						'☜: 
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)				'☆: 조회 조건 데이타 
			strVal = strVal & "&lgIntPrevKey=" & lgIntPrevKey
			strVal = strVal & "&txtProdOrderNo=" & Trim(strProdtOrderNo)
			strVal = strVal & "&GridColCount=" & Trim(C_GRIDCOUNT)
			strVal = strVal & "&txtMaxCount=" & Trim(MaxCount)
			strVal = strVal & "&txtKeyStream=" & lgKeyStream2         '☜: Query Key
			strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
		End If

		Call RunMyBizASP(MyBizASP, strVal)											'☜: 비지니스 ASP 를 가동 

    End With

    DbDtlQuery = True

End Function


Function DbDtlQueryOk(ByVal LngMaxRow)												'☆: 조회 성공후 실행로직 

	Dim LngRow

    '-----------------------
    'Reset variables area
    '-----------------------
	frm1.vspdData2.ReDraw = Fales

	Call InitData()
   
    lgIntFlgMode = parent.OPMD_UMODE

'    If frm1.vspdData2.MaxRows > 0 Then 
'		Call SetToolBar("11001111000111")										'⊙: 버튼 툴바 제어 	
'	End if

	frm1.vspdData2.ReDraw = True

End Function

'=======================================================================================================
'   Function Name : FindData
'   Function Desc : 
'=======================================================================================================
Function FindData()

Dim strProdtOrderNo, strOprNo, strSequence
Dim strHndProdtOrderNo, strHndOprNo, strHndSequence
Dim lRows

    FindData = 0

    With frm1
        
        For lRows = 1 To .vspdData3.MaxRows
        
            .vspdData3.Row = lRows
            .vspdData3.Col = C_ProdtOrderNo2
            strHndProdtOrderNo = .vspdData3.Text
            .vspdData3.Col = C_OprNo2
            strHndOprNo = .vspdData3.Text
            .vspdData3.Col = C_Sequence2
            strHndSequence = .vspdData3.Text

            .vspdData2.Row = frm1.vspdData2.ActiveRow
            .vspdData2.Col = C_ProdtOrderNo1
            strProdtOrderNo = .vspdData2.Text
            .vspdData2.Col = C_OprNo1
            strOprNo = .vspdData2.Text
            .vspdData2.Col = C_Sequence
            strSequence = .vspdData2.Text
            
            If strHndProdtOrderNo = strProdtOrderNo and strHndOprNo = strOprNo and strHndSequence = strSequence Then
				FindData = lRows
				Exit Function
            End If    
        Next
        
    End With        
    
End Function


'=======================================================================================================
'   Function Name : CopyFromHSheet
'   Function Desc : 
'=======================================================================================================
Function CopyFromHSheet(ByVal strProdtOrderNo, strOprNo)

Dim lngRows
Dim boolExist
Dim iCols
Dim strHdnProdtOrderNo
Dim strHdnOprNo
Dim strStatus
Dim strLotReq
Dim strLotGenMthd
Dim strProdInspReq
Dim strFinalInspReq
Dim strAutoRcptFlg
Dim strInsideFlg
Dim strRoutOrder
Dim iCurColumnPos

	ggoSpread.Source = frm1.vspdData2
 	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

    boolExist = False
    
    CopyFromHSheet = boolExist
    
    With frm1

        Call SortHSheet()
        '------------------------------------
        ' Find First Row
        '------------------------------------
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = C_ProdtOrderNo2
            strHdnProdtOrderNo = .vspdData3.Text
            .vspdData3.Col = C_OprNo2
            strHdnOprNo = .vspdData3.Text
			
            If Trim(strProdtOrderNo) = Trim(strHdnProdtOrderNo) and Trim(strOprNo) = Trim(strHdnOprNo) Then
                boolExist = True
                Exit For
            End If    
        Next

	    '------------------------------------
        ' Show Data
        '------------------------------------
         .vspdData3.Row = lngRows
        
        If boolExist = True Then
            
            frm1.vspdData2.Redraw = False
            
            While lngRows <= .vspdData3.MaxRows

	             .vspdData3.Row = lngRows
                
                .vspdData3.Col = C_ProdtOrderNo2
				strHdnProdtOrderNo = .vspdData3.Text
				.vspdData3.Col = C_OprNo2
				strHdnOprNo = .vspdData3.Text

                If Trim(strProdtOrderNo) = Trim(strHdnProdtOrderNo) and Trim(strOprNo) = Trim(strHdnOprNo) Then
					If Trim(strProdtOrderNo) = Trim(strHdnProdtOrderNo) Then
						.vspdData2.MaxRows = .vspdData2.MaxRows + 1
						.vspdData2.Row = .vspdData2.MaxRows
						.vspdData2.Col = 0
						.vspdData3.Col = 0
						.vspdData2.Text = .vspdData3.Text
						
						For iCols = 1 To .vspdData3.MaxCols
						    .vspdData2.Col = iCurColumnPos(iCols)
						    .vspdData3.Col = iCols
						    .vspdData2.Text = .vspdData3.Text
						Next
						
						.vspdData3.Col = 0

						If .vspdData3.Text = ggoSpread.InsertFlag Then
						
							.vspdData3.Col = C_AutoRcptFlg2
							strAutoRcptFlg = .vspdData3.Text
							.vspdData3.Col = C_LotReq2
							strLotReq = .vspdData3.Text
							.vspdData3.Col = C_LotGenMthd2
							strLotGenMthd = .vspdData3.Text
							.vspdData3.Col = C_ProdInspReq2
							strProdInspReq = .vspdData3.Text
							.vspdData3.Col = C_FinalInspReq2
							strFinalInspReq = .vspdData3.Text
							.vspdData3.Col = C_InsideFlag2
							strInsideFlg = .vspdData3.Text
							.vspdData3.Col = C_RoutOrder2
							strRoutOrder = .vspdData3.Text
							
							Call SetSpreadColor(.vspdData2.Row, .vspdData2.Row, strLotReq, strLotGenMthd, strProdInspReq, strFinalInspReq, strAutoRcptFlg, strInsideFlg, strRoutOrder)
						Else

						End If

					End If
				Else
					lngRows = .vspdData3.MaxRows + 1
                End If   
                
                lngRows = lngRows + 1
                
            Wend
            frm1.vspdData2.Redraw = True

        End If
            
    End With        
    
    CopyFromHSheet = boolExist
   
End Function

'=======================================================================================================
'   Function Name : CopyToHSheet
'   Function Desc : 
'=======================================================================================================
Sub CopyToHSheet(ByVal Row)
Dim lRow
Dim iCols
Dim LngCurRow
Dim iCurColumnPos

	ggoSpread.Source = frm1.vspdData2
 	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

	With frm1 
        
	    lRow = FindData

	    If lRow > 0 Then
			LngCurRow = lRow
            .vspdData3.Row = lRow
            .vspdData2.Row = Row
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text
            For iCols = 1 To 26 'vspdData2 의 데이타만 변경한다.
                .vspdData2.Col = iCurColumnPos(iCols)
                .vspdData3.Col = iCols
                .vspdData3.Text = .vspdData2.Text
            Next
        Else
			.vspdData3.MaxRows = .vspdData3.MaxRows + 1
			LngCurRow = .vspdData3.MaxRows
            .vspdData3.Row = .vspdData3.MaxRows
            .vspdData2.Row = Row
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text
       
       
            For iCols = 1 To 26 'vspdData2 의 데이타만 변경한다.
                .vspdData2.Col = iCurColumnPos(iCols)
                .vspdData3.Col = iCols
                .vspdData3.Text = .vspdData2.Text
            Next
        
        End If

		With .vspdData3

			.Redraw = False
			ggoSpread.Source = frm1.vspdData3
			ggoSpread.SSSetRequired C_ReportDt2,				LngCurRow, LngCurRow
			ggoSpread.SSSetRequired C_ReportType2,				LngCurRow, LngCurRow
			ggoSpread.SSSetRequired C_ShiftId2,					LngCurRow, LngCurRow
			ggoSpread.SSSetRequired C_ProdQty2,					LngCurRow, LngCurRow
			
			.Col = C_ReportType2
			.Row = LngCurRow
			
			If Trim(.Text) = "G" Then
				.Col = C_ReasonCd2
				.Text = ""
				.Col = C_ReasonDesc2
				.Text = ""					

				ggoSpread.SpreadLock C_ReasonCd2, LngCurRow, C_ReasonCd2, LngCurRow
				ggoSpread.SpreadLock C_ReasonDesc2, LngCurRow, C_ReasonDesc2, LngCurRow
				ggoSpread.SSSetProtected C_ReasonCd2, LngCurRow, LngCurRow
				ggoSpread.SSSetProtected C_ReasonDesc2, LngCurRow, LngCurRow
				
				If UCase(Trim(GetSpreadText(frm1.vspdData3,C_AutoRcptFlg2,LngCurRow,"X","X"))) = "Y" _
						And UCase(Trim(GetSpreadText(frm1.vspdData3,C_LotReq2,LngCurRow,"X","X"))) = "Y" _
						And UCase(Trim(GetSpreadText(frm1.vspdData3,C_LotGenMthd2,LngCurRow,"X","X"))) = "M" _
						And (UCase(Trim(GetSpreadText(frm1.vspdData3,C_RoutOrder2,LngCurRow,"X","X"))) = "L" _
						Or UCase(Trim(GetSpreadText(frm1.vspdData3,C_RoutOrder2,LngCurRow,"X","X"))) = "S") Then
					ggoSpread.SpreadUnLock C_LotNo2,LngCurRow,C_LotNo2,LngCurRow
					ggoSpread.SpreadUnLock C_LotSubNo2,LngCurRow,C_LotSubNo2,LngCurRow
					ggoSpread.SSSetRequired C_LotNo2,					LngCurRow, LngCurRow
					ggoSpread.SSSetRequired C_LotSubNo2,					LngCurRow, LngCurRow
				Else
					ggoSpread.SpreadUnLock C_LotNo2,LngCurRow,C_LotNo2,LngCurRow
					ggoSpread.SpreadUnLock C_LotSubNo2,LngCurRow,C_LotSubNo2,LngCurRow	
				End If
			Else
				ggoSpread.SpreadUnLock C_ReasonCd2, LngCurRow, C_ReasonCd2, LngCurRow
				ggoSpread.SpreadUnLock C_ReasonDesc2, LngCurRow, C_ReasonDesc2, LngCurRow
				ggoSpread.SSSetRequired C_ReasonCd2, LngCurRow, LngCurRow
				ggoSpread.SSSetRequired C_ReasonDesc2, LngCurRow, LngCurRow
				ggoSpread.SSSetProtected C_LotNo2, LngCurRow, LngCurRow
				ggoSpread.SSSetProtected C_LotSubNo2, LngCurRow, LngCurRow
			End If
    
		End With

	End With
	
End Sub

'=======================================================================================================
'   Function Name : DeleteHSheet
'   Function Desc : 
'=======================================================================================================
Function DeleteHSheet(ByVal strProdtOrderNo, Byval strOprNo, Byval strSequence)

Dim boolExist
Dim lngRows
Dim StrHndProdtOrderNo, strHndOprNo, strHndSequence
 
    DeleteHSheet = False
    boolExist = False
    
    With frm1
    
        Call SortHSheet()
        
	    '------------------------------------
        ' Find First Row
        '------------------------------------
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows

            .vspdData3.Col = C_ProdtOrderNo2
			StrHndProdtOrderNo = .vspdData3.Text
            .vspdData3.Col = C_OprNo2
			strHndOprNo = .vspdData3.Text
            .vspdData3.Col = C_Sequence2
			strHndSequence = .vspdData3.Text

            If strProdtOrderNo = StrHndProdtOrderNo and strHndOprNo = strOprNo and strSequence = strHndSequence Then
                boolExist = True
                Exit For
            End If    
        Next
       
		'------------------------------------
        ' Data Delete
        '------------------------------------
         .vspdData3.Row = lngRows
        
        If boolExist = True Then
            
            While lngRows <= .vspdData3.MaxRows

                .vspdData3.Row = lngRows
				.vspdData3.Col = C_ProdtOrderNo2
				StrHndProdtOrderNo = .vspdData3.Text
				.vspdData3.Col = C_OprNo2
				strHndOprNo = .vspdData3.Text
				.vspdData3.Col = C_Sequence2
				strHndSequence = .vspdData3.Text
                
                If (strProdtOrderNo <> StrHndProdtOrderNo) or (strOprNo <> strHndOprNo) or (strSequence <> strHndSequence) Then
                    lngRows = .vspdData3.MaxRows + 1
                Else
                    .vspdData3.Action = 5
                    .vspdData3.MaxRows = .vspdData3.MaxRows - 1
                End If   

            Wend
            
            ggoSpread.Source = frm1.vspdData2
            
            frm1.vspdData2.Row = lgCurrRow
            frm1.vspdData2.Col = frm1.vspdData2.MaxCols
            ggoSpread.Source = frm1.vspdData2

            frm1.vspdData2.Redraw = True

        End If

    End With

    DeleteHSheet = True
End Function    

'======================================================================================================
' Function Name : SortHSheet
' Function Desc : This function set Muti spread Flag
'======================================================================================================
Function SortHSheet()
    
    With frm1
    
        .vspdData3.BlockMode = True
        .vspdData3.Col = 0
        .vspdData3.Col2 = .vspdData3.MaxCols
        .vspdData3.Row = 1
        .vspdData3.Row2 = .vspdData3.MaxRows
        .vspdData3.SortBy = 0 'SS_SORT_BY_ROW
        
        .vspdData3.SortKey(1) = C_ProdtOrderNo2	' Production Order No
        .vspdData3.SortKey(2) = C_OprNo2		' Operation No        
        .vspdData3.SortKey(3) = C_Sequence2		' Sequence
        
        .vspdData3.SortKeyOrder(1) = 1 'SS_SORT_ORDER_ASCENDING
        .vspdData3.SortKeyOrder(2) = 1 'SS_SORT_ORDER_ASCENDING
        .vspdData3.SortKeyOrder(3) = 1 'SS_SORT_ORDER_ASCENDING
        
        .vspdData3.Col = 0
        .vspdData3.Col2 = .vspdData3.MaxCols
        .vspdData3.Row = 1
        .vspdData3.Row2 = .vspdData3.MaxRows
        .vspdData3.Action = 25 'SS_ACTION_SORT
        .vspdData3.BlockMode = False
        
    End With        
    
End Function

'=======================================================================================================
'   Function Name : FindRow
'   Function Desc : 
'=======================================================================================================
Function FindRow(Byval strProdtOrderNo)

Dim lRows
Dim CompProdtOrderNo, CompOprNo

    FindRow = 0

    With frm1
        
        For lRows = 1 To .vspdData1.MaxRows
        
            .vspdData1.Row = lRows
            .vspdData1.Col = C_ProdtOrderNo
            CompProdtOrderNo = .vspdData1.Text
'            .vspdData1.Col = C_OprNo
'            CompOprNo = .vspdData1.Text
            If CompProdtOrderNo = strProdtOrderNo  Then
				FindRow = lRows
				Exit Function
            End If    
        Next
        
    End With
    
End Function

'========================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 

    Dim lRow        
    Dim lGrpCnt     
   	Dim strVal
	Dim strDel
	Dim Loop2
	Dim strRoutNo, strItemCd, strBaseUnit, strTrackingNo
	Dim StartDate2, EndDate2
	
    DbSave = False                                                          '⊙: Processing is NG
    
'    If Not CntMaxRows(0) Then Exit Function

    StartDate2 = frm1.txtProdFromDt.text
	EndDate2 = UNIDateAdd("D", 1, StartDate2, parent.gDateFormat)	'☆: 초기화면에 뿌려지는 마지막 날짜


    LayerShowHide(1) 
		
    'On Error Resume Next                                                   '☜: Protect system from crashing
	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		.txtFlgMode.value = lgIntFlgMode

    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""
    '-----------------------
    'Data manipulate area
    '-----------------------
'    ggoSpread.Source = frm1.vspdData1 
    For lRow = 1 To .vspdData1.MaxRows
    
        .vspdData1.Row = lRow
        .vspdData1.Col = 0
        
        Select Case .vspdData1.Text

            Case ggoSpread.InsertFlag												'☜: 신규 
			
			
			For Loop2 = 1  to  C_GRIDCOUNT 
				
				strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep & Loop2 &  parent.gColSep					'☜: C=Create

				.vspdData1.Col = C_JobLineCd	'3
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                
                .vspdData1.Col = C_ItemCd	'4
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                
                .vspdData1.Col = C_ProdtOrderNo	'5
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
             
                strVal = strVal & UNIConvDate(Trim(StartDate2)) & parent.gColSep    '6
                strVal = strVal & UNIConvDate(Trim(EndDate2)) & parent.gColSep		'7
                
                .vspdData1.Col =  ((Loop2*5) + 9) - 4   	 '작업계획시간          '8
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                
                .vspdData1.Col = ((Loop2*5) + 9 ) -3    	 '작업지시번호           '9
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                
                .vspdData1.Col = ((Loop2*5) + 9 ) - 2	    '작업지시수량					'10
                strVal = strVal & UNIConvNum(Trim(.vspdData1.Text),0) & parent.gColSep
				
               .vspdData1.Col = ((Loop2*5) + 9 )  - 1   	 '우선순위                '11
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                
                .vspdData1.Col = ((Loop2*5) + 9 )     	 '확정여부                '12
                strVal = strVal & Trim(.vspdData1.Text) & parent.gRowSep
           
           Next     
                                
           lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.UpdateFlag
 
              For Loop2 = 1  to  C_GRIDCOUNT

				strVal = strVal & "U" & parent.gColSep	& lRow & parent.gColSep	& Loop2 &  parent.gColSep				'☜: U=Update
				
  				.vspdData1.Col = C_JobLineCd	'3
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                
                .vspdData1.Col = C_ItemCd	'4
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                
                .vspdData1.Col = C_ProdtOrderNo	'5
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
             
                strVal = strVal & UNIConvDate(Trim(StartDate2)) & parent.gColSep    '6
                strVal = strVal & UNIConvDate(Trim(EndDate2)) & parent.gColSep		'7
                
                .vspdData1.Col =  ((Loop2*5) + 9) - 4   	 '작업계획시간          '8
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                
                .vspdData1.Col = ((Loop2*5) + 9 ) -3    	 '작업지시번호           '9
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                
                .vspdData1.Col = ((Loop2*5) + 9 ) - 2	    '작업지시수량					'10
                strVal = strVal & UNIConvNum(Trim(.vspdData1.Text),0) & parent.gColSep
				
               .vspdData1.Col = ((Loop2*5) + 9 )  - 1   	 '우선순위                '11
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                
                .vspdData1.Col = ((Loop2*5) + 9 )     	 '확정여부                '12
                strVal = strVal & Trim(.vspdData1.Text) & parent.gRowSep
 
             Next   
                                               
                lGrpCnt = lGrpCnt + 1
                
            Case ggoSpread.DeleteFlag												'☜: 삭제 
             
              For Loop2 = 1  to  C_GRIDCOUNT

				strVal = strVal & "D" & parent.gColSep	& lRow & parent.gColSep & Loop2 &  parent.gColSep
				
   				.vspdData1.Col = C_JobLineCd	'3
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                
                .vspdData1.Col = C_ItemCd	'4
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                
                .vspdData1.Col = C_ProdtOrderNo	'5
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
             
                strVal = strVal & UNIConvDate(Trim(StartDate2)) & parent.gColSep    '6
                strVal = strVal & UNIConvDate(Trim(EndDate2)) & parent.gColSep		'7
                
                .vspdData1.Col =  ((Loop2*5) + 9) - 4   	 '작업계획시간          '8
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                
                .vspdData1.Col = ((Loop2*5) + 9 ) -3    	 '작업지시번호           '9
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                
                .vspdData1.Col = ((Loop2*5) + 9 ) - 2	    '작업지시수량					'10
                strVal = strVal & UNIConvNum(Trim(.vspdData1.Text),0) & parent.gColSep
				
               .vspdData1.Col = ((Loop2*5) + 9 )  - 1   	 '우선순위                '11
                strVal = strVal & Trim(.vspdData1.Text) & parent.gColSep
                
                .vspdData1.Col = ((Loop2*5) + 9 )     	 '확정여부                '12
                strVal = strVal & Trim(.vspdData1.Text) & parent.gRowSep

            Next  
                                
                lGrpCnt = lGrpCnt + 1
        End Select	
                
    Next

	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value =  strVal

	Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)										'☜: 비지니스 ASP 를 가동 
	
	End With
	
    DbSave = True																	'⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()
   
    lgIntPrevKey = 0
    lgLngCurRows = 0

	ggoSpread.source = frm1.vspddata1
    frm1.vspdData1.MaxRows = 0
'	ggoSpread.source = frm1.vspddata2
'    frm1.vspdData2.MaxRows = 0
	ggoSpread.source = frm1.vspddata3
    frm1.vspdData3.MaxRows = 0
	
	lgIntFlgMode = parent.OPMD_CMODE
	
	Call RemovedivTextArea
	Call MainQuery
	
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
End Function

'========================================================================================
' Function Name : RemovedivTextArea
' Function Desc : 저장후, 동적으로 생성된 HTML 객체(TEXTAREA)를 Clear시켜 준다.
'========================================================================================
Function RemovedivTextArea()

	Dim ii
		
	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next

End Function



'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
'==============================================================================
' Function : GetHiddenFocus
' Description : 에러발생시 Hidden Spread Sheet를 찾아 SheetFocus에 값을 넘겨줌.
'==============================================================================
Function GetHiddenFocus(lRow, lCol)

	Dim lRows1, lRows2						'Quantity of the Hidden Data Keys Referenced by FindData Function
	Dim strHdnOrdNo, strHdnOprNo, strHdnSeqNo			'Variable of Hidden Keys
	Dim strSeqNo					'Variable of Visible Sheet Keys		
	
	If Trim(lCol) = "" Then
		lCol = C_ReportDt					'If Value of Column is not passed, Assign Value of the First Column in Second Spread Sheet
	End If
	'Find Key Datas in Hidden Spread Sheet
	With frm1.vspdData3
		.Row = lRow
		.Col = C_ProdtOrderNo2			
		strHdnOrdNo = Trim(.Text)
		.Col = C_OprNo2			
		strHdnOprNo = Trim(.Text)
		.Col = C_Sequence2				
		strHdnSeqNo = Trim(.Text)
	End With
	'Compare Key Datas to Visible Spread Sheets
	With frm1
		For lRows1 = 1 To .vspdData1.MaxRows
			.vspdData1.Row = lRows1
			.vspdData1.Col = C_ProdtOrderNo			
			If Trim(.vspdData1.Text) = strHdnOrdNo Then
				.vspdData1.Col = C_OprNo			
				If Trim(.vspdData1.Text) = strHdnOprNo Then
					.vspdData1.Col = C_ProdtOrderNo	
					.vspdData1.focus
					.vspdData1.Action = 0
					lgOldRow = lRows1			'※ If this line is omitted, program could not query Data When errors occur
					ggoSpread.Source = .vspdData2
					.vspdData2.MaxRows = 0
					If CopyFromHSheet(strHdnOrdNo, strHdnOprNo) = True Then
					    For lRows2 = 1 To .vspdData2.MaxRows
							.vspdData2.Row = lRows2
							.vspdData2.Col = C_Sequence
							strSeqNo = .vspdData2.Text
							'Find Key Datas in Second Sheet and then Focus the Cell 
							If Trim(strHdnSeqNo) = Trim(strSeqNo) Then
								Call SheetFocus(lRows2, lCol)
								Exit Function
							End If
					    Next
					End If
				End If	
			End If
		Next
	End With
End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow, lCol)

	frm1.vspdData2.focus
	frm1.vspdData2.Row = lRow
	frm1.vspdData2.Col = lCol
	frm1.vspdData2.Action = 0
	frm1.vspdData2.SelStart = 0
	frm1.vspdData2.SelLength = len(frm1.vspdData2.Text)	
End Function

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : 그리드 현상태를 저장한다.
'========================================================================================
'Sub PopSaveSpreadColumnInf()
'   ggoSpread.Source = gActiveSpdSheet
'    Call ggoSpread.SaveSpreadColumnInf()
'End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : 그리드를 예전 상태로 복원한다.
'========================================================================================
'Sub PopRestoreSpreadColumnInf()
   
'	Dim LngRow
'	Dim strProdtOrderNo
'	Dim strOprNo

'    ggoSpread.Source = gActiveSpdSheet
    
'    If gActiveSpdSheet.Id = "B" Then
'		frm1.vspdData2.Row = frm1.vspdData2.ActiveRow
'		frm1.vspdData2.Col = C_ProdtOrderNo1
'		strProdtOrderNo = Trim(frm1.vspdData2.Text)
'		frm1.vspdData2.Col = C_OprNo1
'		strOprNo = Trim(frm1.vspdData2.Text)
'	End If
    
'    Call ggoSpread.RestoreSpreadInf()
'    Call InitSpreadSheet(gActiveSpdSheet.Id)
	
'	If gActiveSpdSheet.Id = "A" Then
'		Call ggoSpread.ReOrderingSpreadData()
'	ElseIf gActiveSpdSheet.Id = "B" Then
'		Call InitSpreadComboBox
'		Call InitShiftCombo
		
'	    ggoSpread.Source = frm1.vspdData3
'       Call ggoSpread.RestoreSpreadInf()

'		ggoSpread.Source = frm1.vspdData3
'		Call InitSpreadSheet("C")
'		Call ggoSpread.ReOrderingSpreadData()
'		
'		Call CopyFromHsheet(strProdtOrderNo,strOprNo)
				
'		Call InitData()
'	End If
   
'End Sub 

'========================================================================================
' Function Name : SetBlnInsertRow
' Function Desc : Boolean Value whether insert row will be enabled, was passed in this function 
'========================================================================================
Function SetBlnInsertRow(ByVal pvRow)
	
	Dim strMileStoneFlg, strInsideFlag
	
	frm1.vspdData1.Row = pvRow
	frm1.vspdData1.Col = C_ProdtOrderNo
	strInsideFlag = Trim(frm1.vspdData1.value)
	frm1.vspdData1.Col = C_ItemCd
	strMileStoneFlg = Trim(frm1.vspdData1.value)
	
	If strMileStoneFlg <> "" and  strInsideFlag <> "" Then
		SetBlnInsertRow = True
	Else
		SetBlnInsertRow = False	
	End If
	
End Function

'==============================================================================
' Function : SetFieldColor
' Description : 중간 입력 필드의 Color를 맞춤. 
'==============================================================================
Function SetFieldColor(ByVal BlnQueryOk) 

	If BlnQueryOk  = True Then
		Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
		If UCase(Trim(GetSpreadText(frm1.vspdData1,C_AutoRcptFlg,1,"X","X"))) = "Y" Then
			Call ggoOper.SetReqAttr(frm1.txtRcptNo,"N")
			Call ggoOper.SetReqAttr(frm1.txtRcptNo,"D")
		Else
			Call ggoOper.SetReqAttr(frm1.txtRcptNo,"Q")
		End If	
	
'		frm1.txtReportDt.text	= LocSvrDate
'		frm1.txtRcptNo.value = ""
	Else
		Call ggoOper.LockField(Document, "Q")                                   '⊙: Lock  Suitable  Field
	
'		frm1.txtReportDt.text	= ""
'		frm1.txtRcptNo.value = ""
	End If
End Function


Function NumberCheck(Row, Col)
    Dim JobSeq
    
    NumberCheck = False

	frm1.vspdData1.Row = Row
	frm1.vspdData1.Col = Col
	JobSeq = Trim(frm1.vspdData1.Text)
		if isnumeric(JobSeq) then
		else
		  Call displaymsgbox("211434","X", "X","X")
			frm1.vspdData1.Text = 0
		  Exit Function
		End if  
	 NumberCheck = True
End Function
<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name  : Production
'*  2. Function Name : 
'*  3. Program ID  : i2511ma1.asp
'*  4. Program Name  : LOT Tracing 조회 
'*  5. Program Desc  :
'*  6. Comproxy List : +B19029LookupNumericFormat
'                         +B25011ManagePlant
'                         +B25011ManagePlant
'                         +B25018ListPlant
'                         +B25019LookUpPlant
'*  7. Modified date(First) : 2000/04/18
'*  8. Modified date(Last)  : 2003/05/16 
'*  9. Modifier (First)  : Im Hyun Soo
'* 10. Modifier (Last)  :  Lee Seung Wook	
'* 11. Comment  :
'* 12. Common Coding Guide : this mark(☜) means that "Do not change"
'*                                this mark(⊙) Means that "may  change"
'*                                this mark(☆) Means that "must change"
'* 13. History              :
'*                            
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'     1. 선 언 부 
'##########################################################################################################-->
<!--'******************************************  1.1 Inc 선언   **********************************************
' 기능: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->     
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css"> 

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                           

'==========================================  1.2.1 Global 상수 선언  ======================================
'========================================================================================================== 
Const BIZ_PGM_QRY_ID    = "i2511mb1.asp"  
Const BIZ_PGM_LOOKUPHDR_ID   = "i2519mb1.asp"
Const BIZ_PGM_LOOKUPITEMBYPLANT_ID = "p1401mb7.asp" 

Const C_Sep  = "/"
Const C_PROD  = "PROD"
Const C_MATL  = "MATL"

Const C_IMG_PROD = "../../../CShared/image/product.gif"
Const C_IMG_MATL = "../../../CShared/image/material.gif"
Const tvwChild = 4

'==========================================  1.2.2 Global 변수 선언  =====================================
' 1. 변수 표준에 따름. prefix로 g를 사용함.
' 2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'=========================================================================================================  
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgBlnFlgConChg    '☜: Condition 변경 Flag

Dim lgNextNo     
Dim lgPrevNo     
'=========================================================================================================  
'----------------  공통 Global 변수값 정의  -----------------------------------------------------------  

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++  
Dim IsOpenPop
Dim lgBlnBizLoadMenu
Dim lgProcType
'==========================================  2.1.1 InitVariables()  ======================================
' Name : InitVariables()
' Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'====================================================================================================  
Sub InitVariables()

    lgIntFlgMode		= Parent.OPMD_UMODE   
    lgBlnFlgChgValue	= False   
    lgIntGrpCount		= 0     
    '----------  Coding part  ------------------------------------------------------------- 
    IsOpenPop = False            
 
End Sub

'========================================= 2.1.2 LoadInfTB19029() ==================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'===================================================================================================  
Sub LoadInfTB19029()
 <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
 <% Call loadInfTB19029A("Q", "I","NOCOOKIE","MA") %>
End Sub

'========================================  2.2.1 SetDefaultVal()  ======================================
' Name : SetDefaultVal()
' Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'===================================================================================================  
Sub SetDefaultVal()
 frm1.rdoSrchType1.checked = True
End Sub


'------------------------------------------  OpenCondPlant()  -------------------------------------------------
' Name : OpenCondPlant()
' Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------  
Function OpenConPlant()
 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

 IsOpenPop = True

 arrParam(0) = "공장팝업"    
 arrParam(1) = "B_PLANT"       
 arrParam(2) = Trim(frm1.txtPlantCd.Value) 
 arrParam(3) = ""        
 arrParam(4) = ""        
 arrParam(5) = "공장"    
 
    arrField(0) = "PLANT_CD" 
    arrField(1) = "PLANT_NM" 
    
    arrHeader(0) = "공장"
    arrHeader(1) = "공장명"   
    
 arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 IsOpenPop = False

 If arrRet(0) = "" Then
	frm1.txtPlantCd.focus
	Exit Function
 Else
  Call SetConPlant(arrRet)
 End If 
 
End Function
'------------------------------------------  OpenItemCd()  -------------------------------------------------
' Name : OpenItemCd()
' Description : Item PopUp
'---------------------------------------------------------------------------------------------------------  
Function OpenItemCd()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim arrParam0, arrParam1
	 
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("169901","X", "X", "X")   '공장정보가 필요합니다  
		frm1.txtPlantCd.focus
		Exit Function
	End If
	 
	'-----------------------
	'Check Plant CODE  '공장코드가 있는 지 체크 
	'-----------------------
	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.value = ""
		frm1.txtPlantCd.focus
		Exit function
	End If

	lgF0 = Split(lgF0,Chr(11))
	frm1.txtPlantNm.value = lgF0(0)

	If IsOpenPop = True Then Exit Function  
	 
	IsOpenPop = True
	 
	arrParam0 = Trim(frm1.txtPlantCd.value)   
	arrParam1 = Trim(frm1.txtItemCd.value)    

	iCalledAspName = AskPRAspName("I2512PA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I2512PA1","x")
		IsOpenPop = False
		Exit Function
	End If
	 
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam0, arrParam1), _
	"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		Call SetItemCd(arrRet)
	End If 
End Function
'------------------------------------------  OpenLotNo()  -------------------------------------------------
' Name : OpenLotNo()
' Description : Condition BomNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenLotNo()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim Param1, Param2, Param3 , Param4 , Param5
	 
	If IsOpenPop = True Then Exit Function
	 
	If frm1.txtPlantCd.value = "" Then 
		Call DisplayMsgBox("169901","X", "X", "X")   <% '공장정보가 필요합니다 %>
		frm1.txtPlantCd.focus
		Exit Function
	End If  
	 
	 
	If frm1.txtItemCd.value = "" Then
		Call DisplayMsgBox("169915","X", "X", "X")   <% '품목코드를 입력하십시오 %>
		frm1.txtItemCd.focus
		Exit Function
	End If  
	 
	'-----------------------
	'Check Plant CODE  '공장코드가 있는 지 체크 
	'-----------------------
	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
				      lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.value = ""
		frm1.txtPlantCd.focus
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	frm1.txtPlantNm.value = lgF0(0)

	'-----------------------
	'Check ItemCD CODE     '품목코드가 있는 지 체크 
	'-----------------------
	If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S"), _
					  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("122600","X","X","X")
		frm1.txtItemNm.value = ""
		frm1.txtItemCd.focus
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	frm1.txtItemNm.value = lgF0(0)  
	 
	'-----------------------
	'Check ItemCD CODE     '공장코드별 품목코드가 있는 지 체크 
	'-----------------------
	If  CommonQueryRs(" ITEM_CD "," B_ITEM_BY_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S"), _
					  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("122700","X","X","X")
		frm1.txtItemCd.focus
		Exit function
	End If
	 
	IsOpenPop = True

	Param1 = Trim(frm1.txtPlantCd.value)
	Param2 = Trim(frm1.txtItemCd.value) 
	Param3 = Trim(frm1.txtLotNo.value)
	Param4 = Trim(frm1.txtLotSubNo.value)
	Param5 = Trim(frm1.txtItemNm.value)
	 
	iCalledAspName = AskPRAspName("I2511PA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I2511PA1","x")
		IsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1,Param2, Param3,Param4,Param5), _
			"dialogWidth=655px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		frm1.txtLotNo.focus
		Exit Function
	Else
		Call SetLotNo(arrRet)
	End If 
End Function

'------------------------------------------  OpenOrdReltdRef()  -------------------------------------------------
' Name : OpenOrdReltdRef()
' Description : Condition BomNo PopUp
'---------------------------------------------------------------------------------------------------------  

Function OpenOrdReltdRef()
	Dim iCalledAspName
	Dim IntRetCD

	Dim NodX    
	Dim iPos1
	Dim iPos2
	Dim iPos3
	 
	Dim txtLotNo
	Dim txtLotSubNo
	Dim txtItemCd 
	Dim intLevel 
	Dim prntNode
	Dim SelIndex
	 
	Dim arrRet
	Dim Param1, Param2, Param3 , Param4 , Param5 , Param6
	 
Err.Clear   
	 
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("169901","X", "X", "X")   
		frm1.txtPlantCd.focus
		Exit Function
	End If  
	 
	If frm1.txtItemCd.value = "" Then
		Call DisplayMsgBox("169915","X", "X", "X")   
		frm1.txtItemCd.focus
		Exit Function
	End If   

	'-----------------------
	'Check Plant CODE  '공장코드가 있는 지 체크 
	'-----------------------
	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
					  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.value = ""
		frm1.txtPlantCd.focus
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	frm1.txtPlantNm.value = lgF0(0)

	'-----------------------
	'Check ItemCD CODE     '품목코드가 있는 지 체크 
	'-----------------------
	If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S"), _
					  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("122600","X","X","X")
		frm1.txtItemNm.value = ""
		frm1.txtItemCd.focus
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	frm1.txtItemNm.value = lgF0(0)
	  
	'-----------------------
	'Check ItemCD CODE     '공장코드별 품목코드가 있는 지 체크 
	'-----------------------
	If  CommonQueryRs(" ITEM_CD "," B_ITEM_BY_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S"), _
					  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("122700","X","X","X")
		frm1.txtItemCd.focus
		Exit function
	End If
	                                           
	With frm1
		Set NodX = .uniTree1.SelectedItem
		    
		If Not NodX Is Nothing Then 

			'-------------------------------------
			'Hidden Value Init
			'--------------------------------------- 
			Set PrntNode = NodX.Parent
			  
			If PrntNode is Nothing Then 
				iPos1 = InStr(1,NodX.Key, "|^|^|")            
				iPos2 = Instr(iPos1+5,NodX.Key, "|^|^|")        
				iPos3 = Instr(iPos2+5,NodX.Key, "|^|^|")
				txtItemCd   = Trim(Mid(NodX.Key,1,iPos1-1))   
				txtLotNo	= Trim(Mid(NodX.Key,iPos1+5,iPos2-iPos1-5))
				txtLotSubNo = CInt(Trim(Right(NodX.Key,3)))
				    
				IsOpenPop = True
				 
				Param1 = Trim(frm1.txtPlantCd.value)
				Param2 = txtItemCd
				Param3 = txtLotNo
				Param4 = txtLotSubNo
				Param5 = Trim(frm1.txtPlantNm.value)
				Param6 = Trim(frm1.txtItemNm.value)

				iCalledAspName = AskPRAspName("I2511RA1")
				If Trim(iCalledAspName) = "" Then
					IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I2511RA1","x")
					IsOpenPop = False
					Exit Function
				End If

				arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1,Param2, Param3,Param4,Param5,Param6), _
				"dialogWidth=705px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
				 
				IsOpenPop = False
				   
				If arrRet(0) = "" Then
					frm1.txtLotNo.focus
					Exit Function
				End if
		   
			Else
		      	'SelIndex = NodX.Index   
					      
				iPos1 = InStr(1,NodX.Key, "|^|^|")             
				iPos2 = Instr(iPos1+5,NodX.Key, "|^|^|")       
				iPos3 = Instr(iPos2+5,NodX.Key, "|^|^|")
				txtItemCd        = Trim(Mid(NodX.Key,1,iPos1-1))   
				txtLotNo   = Trim(Mid(NodX.Key,iPos1+5,iPos2-iPos1-5))
				txtLotSubNo      = CInt(Trim(Mid(NodX.Key,iPos2+5,iPos3-iPos2-5)))
					   
				IsOpenPop = True
					 
				Param1 = Trim(frm1.txtPlantCd.value)
				Param2 = txtItemCd
				Param3 = txtLotNo
				Param4 = txtLotSubNo
				Param5 = Trim(frm1.txtPlantNm.value)
				Param6 = Trim(frm1.txtItemNm.value)

				iCalledAspName = AskPRAspName("I2511RA1")
				If Trim(iCalledAspName) = "" Then
					IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I2511RA1","x")
					IsOpenPop = False
					Exit Function
				End If
					    
				arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1,Param2, Param3,Param4,Param5,Param6), _
				"dialogWidth=705px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
					 
				IsOpenPop = False
					   
				If arrRet(0) = "" Then
					frm1.txtLotNo.focus
					Exit Function
				End if
			End IF
		Else 
			Call DisplayMsgBox("169925","X", "X", "X")
			Exit function
		End If
		    
		Set NodX = Nothing
		Set PrntNode = Nothing
	End With
End Function


'------------------------------------------  OpenOnhandRef()  -------------------------------------------------
' Name : OpenOnhandRef()
' Description : Condition OnhandRef 
'---------------------------------------------------------------------------------------------------------  

Function OpenOnhandRef()
	Dim iCalledAspName
	Dim IntRetCD

	Dim NodX    
	Dim iPos1
	Dim iPos2
	Dim iPos3
	 
	Dim txtLotNo
	Dim txtLotSubNo
	Dim txtItemCd 
	Dim intLevel 
	Dim prntNode
	Dim SelIndex
	 
	Dim arrRet
	Dim Param1, Param2, Param3 , Param4 , Param5 , Param6
	 
Err.Clear                                                           
	   
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("169901","X", "X", "X")   '공장정보가 필요합니다  
		frm1.txtPlantCd.focus
		Exit Function
	End If  
	 
	If frm1.txtItemCd.value = "" Then
		Call DisplayMsgBox("169915","X", "X", "X")   '품목코드를 입력하십시오  
		frm1.txtItemCd.focus
		Exit Function
	End If   

	'-----------------------
	'Check Plant CODE  '공장코드가 있는 지 체크 
	'-----------------------
	If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
				      lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.value = ""
		frm1.txtPlantCd.focus
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	frm1.txtPlantNm.value = lgF0(0)
	 
	'-----------------------
	'Check ItemCD CODE     '품목코드가 있는 지 체크 
	'-----------------------
	If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S"), _
				 	  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	   
		Call DisplayMsgBox("122600","X","X","X")
		frm1.txtItemNm.value = ""
		frm1.txtItemCd.focus
		Exit function
	End If
	lgF0 = Split(lgF0,Chr(11))
	frm1.txtItemNm.value = lgF0(0)
	 
	'-----------------------
	'Check ItemCD CODE    
	'-----------------------
	If  CommonQueryRs(" ITEM_CD "," B_ITEM_BY_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S"), _
					  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

		Call DisplayMsgBox("122700","X","X","X")
		frm1.txtItemCd.focus
		Exit function
	End If

	'-----------------------
	'Check txtLotSubNo CODE    
	'-----------------------
	If  CommonQueryRs(" LOT_NO "," I_LOT_MASTER ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S") & _
					  " AND LOT_NO = " & FilterVar(frm1.txtLotNo.Value, "''", "S") & " AND LOT_SUB_NO = " & Trim(frm1.txtLotSubNo.Value), _
					  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
	     
		Call DisplayMsgBox("161101","X","X","X")
		frm1.txtLotNo.focus
		Exit function
	End If

	With frm1
		Set NodX = .uniTree1.SelectedItem
		    
		If Not NodX Is Nothing Then 
			'-------------------------------------
			'Hidden Value Init
			'--------------------------------------- 
			  
			Set PrntNode = NodX.Parent
			  
			If PrntNode is Nothing Then 
				iPos1 = InStr(1,NodX.Key, "|^|^|")           
				iPos2 = Instr(iPos1+5,NodX.Key, "|^|^|")     
				iPos3 = Instr(iPos2+5,NodX.Key, "|^|^|")
				txtItemCd       = Trim(Mid(NodX.Key,1,iPos1-1))   
				txtLotNo		= Trim(Mid(NodX.Key,iPos1+5,iPos2-iPos1-5))
				txtLotSubNo     = CInt(Trim(Right(NodX.Key,3)))
				    
				IsOpenPop = True
				 
				Param1 = Trim(frm1.txtPlantCd.value)
				Param2 = txtItemCd
				Param3 = txtLotNo
				Param4 = txtLotSubNo
				Param5 = Trim(frm1.txtPlantNm.value)
				Param6 = Trim(frm1.txtItemNm.value)
				
				iCalledAspName = AskPRAspName("I2511RA2")
				If Trim(iCalledAspName) = "" Then
					IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I2511RA2","x")
					IsOpenPop = False
					Exit Function
				End If
				
				arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1,Param2, Param3,Param4,Param5,Param6), _
				"dialogWidth=705px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
				 
				IsOpenPop = False
				   
				If arrRet(0) = "" Then
					frm1.txtLotNo.focus
					Exit Function
				End if
			   
			Else
				      
				iPos1 = InStr(1,NodX.Key, "|^|^|")            
				iPos2 = Instr(iPos1+5,NodX.Key, "|^|^|")      
				iPos3 = Instr(iPos2+5,NodX.Key, "|^|^|")
				txtItemCd        = Trim(Mid(NodX.Key,1,iPos1-1))   
				txtLotNo   = Trim(Mid(NodX.Key,iPos1+5,iPos2-iPos1-5))
				txtLotSubNo      = CInt(Trim(Mid(NodX.Key,iPos2+5,iPos3-iPos2-5)))
				   
				IsOpenPop = True
				 
				Param1 = Trim(frm1.txtPlantCd.value)
				Param2 = txtItemCd
				Param3 = txtLotNo
				Param4 = txtLotSubNo
				Param5 = Trim(frm1.txtPlantNm.value)
				Param6 = Trim(frm1.txtItemNm.value)

				iCalledAspName = AskPRAspName("I2511RA2")
				If Trim(iCalledAspName) = "" Then
					IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I2511RA2","x")
					IsOpenPop = False
					Exit Function
				End If
				
				arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1,Param2, Param3,Param4,Param5,Param6), _
				"dialogWidth=705px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
				 
				IsOpenPop = False
				   
				If arrRet(0) = "" Then
					frm1.txtLotNo.focus
					Exit Function
				End if
			End IF
		else 
			Call DisplayMsgBox("169925","X", "X", "X")
			exit function
		End If
		    
		Set NodX = Nothing
		Set PrntNode = Nothing
	End With
End Function

'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++  
'------------------------------------------  SetItemCd()  --------------------------------------------------
' Name : SetItemCd()
' Description : Item Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------  
Function SetItemCd(byval arrRet)
	frm1.txtItemCd.Value    = arrRet(0)  
	frm1.txtItemNm.Value    = arrRet(1)
	frm1.txtItemCd.focus
End Function

'------------------------------------------  SetConPlant()  --------------------------------------------------
' Name : SetConPlant()
' Description : Condition Plant Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------  
Function SetConPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)  
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus  
End Function

'------------------------------------------  SetBomNo()  --------------------------------------------------
' Name : SetBomNo()
' Description : Bom No Popup에서 return된 값 
'---------------------------------------------------------------------------------------------------------  
Function SetLotNo(byval arrRet)
	frm1.txtLotNo.Value    = arrRet(0)  
	frm1.txtLotSubNo.Value = arrRet(1)
    frm1.txtLotNo.focus 
End Function

'==========================================================================================
'   Function Name :LookUpHdr
'   Function Desc :선택한 품목의 Lot Header Data를 읽는다.
'==========================================================================================
 

Sub LookUpHdr(ByVal txtItemCd,ByVal txtLotNo, ByVal txtLotSubNo)

 Dim strVal 
 Call ggoOper.ClearField(Document, "2")
 Call ggoOper.FormatField(Document,"2", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
 Call ggoOper.LockField(Document, "Q") 
 If Trim(frm1.txtSrchType.value) = "2" Then 
  'Call ggoOper.SetReqAttr(frm1.txtLotNo,"D")
 End If
 
 Call LayerShowHide(1)               
 '------------------------------
 ' Server Logic Call
 '------------------------------
 strVal = BIZ_PGM_LOOKUPHDR_ID & "?txtMode=" & Parent.UID_M0001    
 strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)  
 strVal = strVal & "&txtItemCd="   & txtItemCd       
 strVal = strVal & "&txtLotNo="    & txtLotNo       
 strVal = strVal & "&txtLotSubNo=" & txtLotSubNo    
 
 Call RunMyBizASP(MyBizASP, strVal)        

End Sub

'==========================================================================================
'   Function Name :LookUpItemByPlant
'   Function Desc :선택한 품목의 Item Acct를 읽는다.
'==========================================================================================
 
Sub LookUpItemByPlant(ByVal str,ByVal iWhere)
    
	Err.Clear              
	    
	Dim strVal

	'frm1.txtHdnItemAcct.value = ""
	  
	Call LayerShowHide(1)
	       
	strVal = BIZ_PGM_LOOKUPITEMBYPLANT_ID	& "?txtMode=" & Parent.UID_M0001   
	strVal = strVal & "&txtPlantCd="		& Trim(frm1.txtPlantCd.value)  
	strVal = strVal & "&txtItemCd="			& Trim(str) 
	strVal = strVal & "&iPos="				& iWhere
	strVal = strVal & "&CurDate="			& UniConvDateAToB(GetSvrDate, Parent.gServerDateFormat, Parent.gDateFormat)     
	Call RunMyBizASP(MyBizASP, strVal)         

End Sub

'========================================================================================
' Function Name : InitTreeImage
' Function Desc : 이미지 초기화 
'========================================================================================
 
Function InitTreeImage()
 Dim NodX, lHwnd
 
 With frm1

 .uniTree1.SetAddImageCount = 2
 .uniTree1.Indentation = "200" 
 .uniTree1.AddImage C_IMG_PROD, C_PROD, 0          
 .uniTree1.AddImage C_IMG_MATL, C_MATL, 0

 .uniTree1.OLEDragMode = 0             
 .uniTree1.OLEDropMode = 0
 
 End With

End Function

'==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================  
Sub Form_Load()

	Call InitVariables              
	Call LoadInfTB19029             
	Call ggoOper.FormatField(Document, "2", CInt(ggAmtOfMoney.DecPoint), CInt(ggQty.DecPoint), _ 
	                    CInt(ggUnitCost.DecPoint), CInt(ggExchRate.DecPoint), Parent.gDateFormat)
	 
	Call ggoOper.LockField(Document, "N")        
	Call SetToolbar("11000000000011")
	Call SetDefaultVal
	Call InitTreeImage 
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtItemCd.focus    
	Else
		frm1.txtPlantCd.focus 
	End If

End Sub

'==========================================================================================
'   Event Name : rdoSrchType1_OnClick
'   Event Desc : 정전개 선택시 
'==========================================================================================
 
Sub rdoSrchType1_OnClick()
	Call ggoOper.SetReqAttr(frm1.txtLotNo,"N")
End Sub

'==========================================================================================
'   Event Name : uniTree1_NodeClick
'   Event Desc : Node Click시 Look Up Call
'==========================================================================================
Sub uniTree1_NodeClick(ByVal Node)
    Dim NodX
    
 Dim iPos1
 Dim iPos2
 Dim iPos3
 
 Dim txtLotNo
 Dim txtLotSubNo
 Dim txtItemCd 
 Dim intLevel 
 Dim prntNode
 Dim SelIndex
 
 Err.Clear                                                               
   
 With frm1
 
    Set NodX = .uniTree1.SelectedItem
    
    If Not NodX Is Nothing Then 

  '-------------------------------------
  'Hidden Value Init
  '--------------------------------------- 
  
	Set PrntNode = NodX.Parent
  
		If PrntNode is Nothing Then 
			iPos1 = InStr(1,NodX.Key, "|^|^|")            
			iPos2 = Instr(iPos1+5,NodX.Key, "|^|^|")      
			iPos3 = Instr(iPos2+5,NodX.Key, "|^|^|")
			txtItemCd        = Trim(Mid(NodX.Key,1,iPos1-1))   
			txtLotNo   = Trim(Mid(NodX.Key,iPos1+5,iPos2-iPos1-5))
			txtLotSubNo      = CInt(Trim(Right(NodX.Key,3)))
			 
			Call LookUpHdr(txtItemCd ,txtLotNo,txtLotSubNo) 
   
		Else
			iPos1		= InStr(1,NodX.Key, "|^|^|")            
			iPos2		= Instr(iPos1+5,NodX.Key, "|^|^|")      
			iPos3		= Instr(iPos2+5,NodX.Key, "|^|^|")
			txtItemCd    = Trim(Mid(NodX.Key,1,iPos1-1))   
			txtLotNo		= Trim(Mid(NodX.Key,iPos1+5,iPos2-iPos1-5))
			txtLotSubNo  = CInt(Trim(Mid(NodX.Key,iPos2+5,iPos3-iPos2-5)))
   
			Call LookUpHdr(txtItemCd ,txtLotNo,txtLotSubNo) 
   
		End IF
	End If
    
    Set NodX = Nothing
    Set PrntNode = Nothing
    
End With

End Sub
'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
 
Function FncQuery() 
    Dim IntRetCD 
    FncQuery = False                                                     
    Err.Clear                                                            

    '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X", "X")    
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then        
       Call SetDefaultVal  
       Exit Function
    End If

   '-----------------------
   'Erase contents area
   '----------------------- 
 frm1.uniTree1.Nodes.Clear            
    Call ggoOper.ClearField(Document, "2")      
    Call InitVariables               

 '-----------------------
 'Check Plant CODE  '공장코드가 있는 지 체크 
 '-----------------------
 If  CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
   
  Call DisplayMsgBox("125000","X","X","X")
  frm1.txtPlantNm.value = ""
  frm1.txtPlantCd.focus
  Exit function
 End If
 lgF0 = Split(lgF0,Chr(11))
 frm1.txtPlantNm.value = lgF0(0)
 
 
 '-----------------------
 'Check ItemCD CODE     '품목코드가 있는 지 체크 
 '-----------------------
 If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S"), _
  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
   
  Call DisplayMsgBox("122600","X","X","X")
  frm1.txtItemNm.value = ""
  frm1.txtItemCd.focus
  Exit function
 End If
 lgF0 = Split(lgF0,Chr(11))
 frm1.txtItemNm.value = lgF0(0)
  
 '-----------------------
 'Check ItemCD CODE     '공장코드별 품목코드가 있는 지 체크 
 '-----------------------
 If  CommonQueryRs(" ITEM_CD "," B_ITEM_BY_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S"), _
  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
   
  Call DisplayMsgBox("122700","X","X","X")
  frm1.txtItemCd.focus
  Exit function
 End If
 
    
   '-----------------------
   ' Check txtLotSubNo 
   '----------------------- 
 If isNumeric(frm1.txtLotSubNo.value) = False Then
        Call DisplayMsgBox("700119","X",frm1.txtLotSubNo.ALT ,"X") 
        frm1.txtLotSubNo.focus()
        Exit Function
    End If

 '-----------------------
 'Check txtLotSubNo CODE  
 '-----------------------
 If  CommonQueryRs(" LOT_NO "," I_LOT_MASTER ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND ITEM_CD= " & FilterVar(frm1.txtItemCd.value, "''", "S") & _
               " AND LOT_NO = " & FilterVar(frm1.txtLotNo.Value, "''", "S") & " AND LOT_SUB_NO = " & Trim(frm1.txtLotSubNo.Value), _
  lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
     
  Call DisplayMsgBox("161101","X","X","X")
  frm1.txtLotNo.focus
  Exit function
 End If
    
 If DbQuery = False Then
	Exit Function
 End if
       
    FncQuery = True                
        
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
 
Function FncPrint() 
    On Error Resume Next   
    Call Parent.FncPrint()                                                
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
 
Function FncPrev() 
    Dim strVal
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                   
        Call DisplayMsgBox("900002","X", "X", "X")                             
        Exit Function
    ElseIf lgPrevNo = "" Then
		Call DisplayMsgBox("900011","X", "X", "X")
		Exit Function
    End If

    strVal = BIZ_PGM_ID & "?txtMode="		& Parent.UID_M0001       
    strVal = strVal		& "&txtPlantCd="	& lgPrevNo       
    
 Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
 
Function FncNext() 
    Dim strVal

    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                  
        Call DisplayMsgBox("900002","X", "X", "X")                             
        Exit Function
    ElseIf lgNextNo = "" Then
		Call DisplayMsgBox("900011","X", "X", "X")
		Exit Function    
    End If
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001       
    strVal = strVal & "&txtPlantCd=" & lgNextNo      
    
 Call RunMyBizASP(MyBizASP, strVal)

End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
 
Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)           
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
 
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE, False)                                  
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
 
Function FncExit() 
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
 
Function DbQuery() 
    Dim PrntKey
    Dim strVal
    Dim LotSubNo
    Err.Clear            
    
    DbQuery = False            
    
    Call LayerShowHide(1)               
    
    
    frm1.txtUpdtUserId.value= Parent.gUsrID    
    
    strVal = BIZ_PGM_QRY_ID & "?txtMode="	& Parent.UID_M0001    
    strVal = strVal & "&txtPlantCd="		& Trim(frm1.txtPlantCd.value) 
    strVal = strVal & "&txtItemCd="			& Trim(frm1.txtItemCd.value) 
    strVal = strVal & "&txtLotNo="			& Trim(frm1.txtLotNo.value)
      
    if  Len(frm1.txtLotSubNo.value) = 1 then
		LotSubNo = "00"&frm1.txtLotSubNo.value
    Elseif Len(frm1.txtLotSubNo.value) = 2 then
		LotSubNo = "0"&frm1.txtLotSubNo.value
    Elseif Len(frm1.txtLotSubNo.value) = 3 then
		LotSubNo = frm1.txtLotSubNo.value
    end if
    
	strVal = strVal & "&txtLotSubNo="		& Trim(LotSubNo)  
    strVal = strVal & "&txtHdnItemAcct="	& Trim(frm1.txtHdnItemAcct.value)
    strVal = strVal & "&txtUpdtUserId="		& Trim(frm1.txtUpdtUserId.value)
    
    If frm1.rdoSrchType1.checked = True Then
		strVal = strval & "&rdoSrchType=" & frm1.rdoSrchType1.value   
		frm1.txtSrchType.value = 1
    ElseIf frm1.rdoSrchType2.checked = True Then
		strVal = strval & "&rdoSrchType=" & frm1.rdoSrchType2.value   
		frm1.txtSrchType.value = 2
    End If          
    
    Call RunMyBizASP(MyBizASP, strVal)       
 
    DbQuery = True              

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
 
Function DbQueryOk()            

    lgIntFlgMode = Parent.OPMD_UMODE         
    
    Call ggoOper.LockField(Document, "Q")    

    Call SetToolbar("11000000000111")
    
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
	<TABLE  <%=LR_SPACE_TYPE_00%>>
		<TR>
			<TD <%=HEIGHT_TYPE_00%> >
			</TD>
		</TR>
		<TR HEIGHT=23>
			<TD WIDTH=100%>
				<TABLE  <%=LR_SPACE_TYPE_10%>>
					<TR>
						<TD WIDTH=10>&nbsp;</TD>
						<TD CLASS="CLSMTABP">
							<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Lot Tracing</font></TD>
									<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
								</TR>
							</TABLE>
						</TD>
						<TD WIDTH=* align=right><A href="vbscript:OpenOnhandRef()">현재고정보</A> | <A href="vbscript:OpenOrdReltdRef()">Lot출고정보</A></TD>
						<TD WIDTH=10>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR HEIGHT=*>
			<TD WIDTH=100% CLASS="Tab11">
				<TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD <%=HEIGHT_TYPE_02%> >
						</TD>
					</TR>
					<TR>
						<TD HEIGHT=20 WIDTH=100%>
							<FIELDSET CLASS="CLSFLD">
								<TABLE <%=LR_SPACE_TYPE_40%>>
									<TR> 
										<TD CLASS=TD5 NOWRAP>공장</TD>
										<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=37 tag="14"></TD>
										<TD CLASS=TD5 NOWRAP>전개구분</TD>
										<TD CLASS=TD6 NOWRAP><SPAN STYLE="width:70;"><INPUT TYPE="RADIO" NAME="rdoSrchType" ID="rdoSrchType1" CLASS="RADIO" tag="1X" Value="FM" CHECKED><LABEL FOR="rdoSrchType1">정전개</LABEL></SPAN>
										                     <SPAN STYLE="width:70;"><INPUT TYPE="RADIO" NAME="rdoSrchType" ID="rdoSrchType2" CLASS="RADIO" tag="1X" Value="BM"><LABEL FOR="rdoSrchType2">역전개</LABEL></SPAN></TD>
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>품목</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="12XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd()" >&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=28 tag="14"></TD>
										<TD CLASS=TD5 NOWRAP>LOT번호</TD>
										<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLotNo" SIZE=20 MAXLENGTH=25 tag="12XXXU" ALT="LOT번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnLotNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenLotNo()">&nbsp;
										      <INPUT TYPE=TEXT NAME="txtLotSubNo" SIZE=5 MAXLENGTH=3 tag="12" ALT="LOT NO 순번"></TD>
									</TR> 
								</TABLE>
							</FIELDSET>
						</TD>
					</TR>
					<TR>
						<TD <%=HEIGHT_TYPE_03%> WIDTH=100%>
						</TD>
					</TR>
					<TR>
						<TD>
							<TABLE CLASS="BasicTB" CELLSPACING=0>
								<TR>
									<TD WIDTH=50% HEIGHT=* valign=top>
									<script language =javascript src='./js/i2511ma1_uniTree1_N368630582.js'></script>                  
									</TD>
									<TD WIDTH=50% HEIGHT=* valign=top>
										<FIELDSET>
											<TABLE CLASS="BasicTB" CELLSPACING=0>
												<TR>
													<TD CLASS=TD5 NOWRAP>품목</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd2" SIZE=18 MAXLENGTH=18  tag="24" ALT="품목"></TD>
												</TR>
												<TR>
													 <TD CLASS=TD5 NOWRAP>품목명</TD>
													 <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemNm2" SIZE=35 MAXLENGTH=40  tag="24" ALT="품목명"></TD>
												</TR>            
												<TR>             
													<TD CLASS=TD5 NOWRAP>Order번호</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOrdNo" SIZE=15 MAXLENGTH=18  tag="24" ALT="Order번호">&nbsp;<INPUT TYPE=TEXT NAME="txtOrdSubNo" SIZE=5 MAXLENGTH=4  tag="24" ALT="ORDSUBNO"></TD>
												</TR>
												<TR>             
													<TD CLASS=TD5 NOWRAP>ORDER유형</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOrdType" SIZE=4 MAXLENGTH=4  tag="24" ALT="ORDER유형"></TD>
												</TR>
											</TABLE> 
										</FIELDSET> 
										<FIELDSET>
											<TABLE CLASS="BasicTB" CELLSPACING=0>
												<TR>
													<TD CLASS=TD5 NOWRAP>LOT생성일</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLotGenDt" SIZE=10  tag="24" ALT="LOT생성일"></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>단위</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemUnit" SIZE=4 tag="24" ALT="단위"></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>입고수량</TD>
													<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/i2511ma1_fpDoubleSingle1_txtRcptQty.js'></script></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>Tracking No</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=15 MAXLENGTH=25  tag="24" ALT="Tracking No"></TD>
												</TR>                                                            
											</TABLE>
										</FIELDSET>
									</TD>
								</TR>
							</TABLE>
						</TD>         
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD <%=HEIGHT_TYPE_01%> >
			</TD>
		</TR>
		<TR HEIGHT=20 >
			<TD>
				<TABLE <%=LR_SPACE_TYPE_30%> >
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH="100%" HEIGHT=<%=BizSize%>>
				<IFRAME NAME="MyBizASP" SRC="../../Blank.htm" WIDTH="100%" HEIGHT=20 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
			</TD>
		</TR>
	</TABLE>
		<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtHdnItemAcct" tag="14" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtSrchType" tag="14" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               


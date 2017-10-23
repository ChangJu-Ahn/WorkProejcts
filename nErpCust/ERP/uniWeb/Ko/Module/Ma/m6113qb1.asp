<%'======================================================================================================
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2003-05-26
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Kim Jin Ha
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
Option Explicit
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                          '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'☜ : DBAgent Parameter 선언 
Dim rs1, rs2, rs3, rs4, rs5, rs6							'☜ : DBAgent Parameter 선언 
Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim iTotstrData
Dim lgStrPrevKey                                            '☜ : 이전 값 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT

Dim lgDataExist
Dim lgPageNo

Dim ICount  		                                        '   Count for column index
Dim arrRsVal(5)											'* : 화면에 조회해온 Name을 담아놓기 위해 만든 Array	
Dim iFrPoint
iFrPoint=0

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "M", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q", "M", "NOCOOKIE", "QB")
	
     Call HideStatusWnd 
     
     lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
     lgSelectList     = Request("lgSelectList")
     lgTailList       = Request("lgTailList")
     lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 

     Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
     call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query


'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
 Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100            

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       iFrPoint     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo) 
    End If
    
    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
    Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
           PvArr(iLoopCount) = lgstrData	
		   lgstrData = ""
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop
	
	iTotstrData = Join(PvArr, "")
	
    If iLoopCount < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
 Sub FixUNISQLData()
	Dim iStrSql
    Redim UNISqlId(6)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(6,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 
    UNISqlId(0) = "M6113QA101"     
    UNISqlId(1) = "M5111QA102"								              '사업장명 
    UNISqlId(2) = "M6111QA102"								              '경비항목명 
    UNISqlId(3) = "M6111QA103"								              '진행구분명     	
    UNISqlId(4) = "M6111QA105"								              '지급처명 
    UNISqlId(5) = "M2111QA302"											  '공장 
	UNISqlId(6) = "M2111QA303"								              '품목     
    
	   																  'Reusage is Recommended
    UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 
	 
	iStrSql = ""
    '---사업장 
    If Len(Trim(Request("txtBizArea"))) Then
    	iStrSql = iStrSql & " AND B.BIZ_AREA =  " & FilterVar(UCase(Request("txtBizArea")), "''", "S") & " "
    End If
    '---경비항목 
    If Len(Trim(Request("txtChargeType"))) Then
    	iStrSql = iStrSql & " AND B.CHARGE_TYPE =  " & FilterVar(UCase(Request("txtChargeType")), "''", "S") & " "
    End If
    '---진행구분 
    If Len(Trim(Request("txtProcessStep"))) Then
    	iStrSql = iStrSql & " AND B.PROCESS_STEP =  " & FilterVar(UCase(Request("txtProcessStep")), "''", "S") & " "
    End If
     '---지급처 
    If Len(Trim(Request("txtBpCd"))) Then
    	iStrSql = iStrSql & " AND B.PAYEE_CD =  " & FilterVar(UCase(Request("txtBpCd")), "''", "S") & " "
    End If
    '---공장 
    If Len(Trim(Request("txtPlantCd"))) Then
	    iStrSql = iStrSql & " AND A.PLANT_CD  =  " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	End IF
	
     '---품목 
    If Len(Trim(Request("txtItemCd"))) Then
    	iStrSql = iStrSql & " AND A.ITEM_CD =  " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "
    End If    
     '---발생일자 
    If Len(Trim(Request("txtChargeFrDt"))) Then
    	iStrSql = iStrSql & " AND B.CHARGE_DT >=  " & FilterVar(UNIConvDate(Trim(Request("txtChargeFrDt"))), "''", "S") & ""
    Else
    	iStrSql = iStrSql & " AND B.CHARGE_DT >= " & FilterVar("1900-01-01", "''", "S") & ""
    End If
    If Len(Trim(Request("txtChargeToDt"))) Then
    	iStrSql = iStrSql & " AND B.CHARGE_DT <=  " & FilterVar(UNIConvDate(Trim(Request("txtChargeToDt"))), "''", "S") & ""
    Else
    	iStrSql = iStrSql & " AND B.CHARGE_DT <= " & FilterVar("2999-12-30", "''", "S") & ""
    End If    
    '---발주번호 
    If Len(Trim(Request("txtPoNo"))) Then
    	iStrSql = iStrSql & " AND A.PO_NO =  " & FilterVar(UCase(Request("txtPoNo")), "''", "S") & " "
    End If
    
    UNIValue(0,1)  = iStrSql
    UNIValue(0,UBound(UNIValue,2)    ) = Trim(lgTailList)	'---Order By 조건 
    
    UNIValue(1,0)  = " " & FilterVar(UCase(Request("txtBizArea")), "''", "S") & " "
    UNIValue(2,0)  = " " & FilterVar(UCase(Request("txtChargeType")), "''", "S") & " "
    UNIValue(3,0)  = " " & FilterVar(UCase(Request("txtProcessStep")), "''", "S") & " "
    UNIValue(4,0)  = " " & FilterVar(UCase(Request("txtBpCd")), "''", "S") & " "
    UNIValue(5,0)  = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
    UNIValue(6,0)  = " " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
    UNIValue(6,1)  = " " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "
     
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    On Error Resume Next
    SetConditionData = false
    
	If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing        
        If Len(Request("txtBizArea")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "사업장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
	       Call SetFoucs("txtBizArea")
		   Exit Function
	    End If
    Else    
		arrRsVal(0) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
         If Len(Request("txtChargeType")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "경비항목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True
	       Call SetFoucs("txtChargeType")
		   Exit Function	
		End If
    Else    
		arrRsVal(1) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If
	
	If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
         If Len(Request("txtProcessStep")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "진행구분", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True
	       Call SetFoucs("txtProcessStep")
		   Exit Function		
		End If
    Else    
		arrRsVal(2) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If
    
    If  rs4.EOF And rs4.BOF Then
        rs4.Close
        Set rs4 = Nothing
         If Len(Request("txtBpCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "지급처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True
	       Call SetFoucs("txtBpCd")
		   Exit Function			
		End If
    Else    
		arrRsVal(3) = rs4(1)
        rs4.Close
        Set rs4 = Nothing
    End If
	
	If  rs5.EOF And rs5.BOF Then
        rs5.Close
        Set rs5 = Nothing
        If Len(Request("txtPlantCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(4) = rs5(1)
        rs5.Close
        Set rs5 = Nothing
    End If
    
    If  rs6.EOF And rs6.BOF Then
        rs6.Close
        Set rs6 = Nothing
         If Len(Request("txtItemCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(5) = rs6(1)
        rs6.Close
        Set rs6 = Nothing
    End If

    SetConditionData = True
End Function

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
 Sub QueryData()
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1,rs2,rs3,rs4,rs5,rs6)			
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If  
    
    Dim FalsechkFlg
    FalsechkFlg = False      
    
    'If SetConditionData = False Then Exit Sub
    
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing        
        If Len(Request("txtBizArea")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "사업장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
	       Call SetFoucs("txtBizArea")
		   Exit Sub
	    End If
    Else    
		arrRsVal(0) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing
         If Len(Request("txtChargeType")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "경비항목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True
	       Call SetFoucs("txtChargeType")
		   Exit Sub	
		End If
    Else    
		arrRsVal(1) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If
	
	If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing
         If Len(Request("txtProcessStep")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "진행구분", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True
	       Call SetFoucs("txtProcessStep")
		   Exit Sub		
		End If
    Else    
		arrRsVal(2) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If
    
    If  rs4.EOF And rs4.BOF Then
        rs4.Close
        Set rs4 = Nothing
         If Len(Request("txtBpCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "지급처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True
	       Call SetFoucs("txtBpCd")
		   Exit Sub			
		End If
    Else    
		arrRsVal(3) = rs4(1)
        rs4.Close
        Set rs4 = Nothing
    End If
	
	If  rs5.EOF And rs5.BOF Then
        rs5.Close
        Set rs5 = Nothing
        If Len(Request("txtPlantCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(4) = rs5(1)
        rs5.Close
        Set rs5 = Nothing
    End If
    
    If  rs6.EOF And rs6.BOF Then
        rs6.Close
        Set rs6 = Nothing
         If Len(Request("txtItemCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
		End If
    Else    
		arrRsVal(5) = rs6(1)
        rs6.Close
        Set rs6 = Nothing
    End If
        
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetFoucs
' Desc : 
'----------------------------------------------------------------------------------------------------------
Sub SetFoucs(ByVal Opt)
	Response.Write "<Script Language=vbscript>"								& vbCr
	Response.Write "With parent.frm1"										& vbCr										
	Response.Write "	If  """ & Opt & """ = ""txtBizArea"" Then "			& vbCr
	Response.Write "		.txtBizArea.focus() "							& vbCr	
	Response.Write "	Elseif  """ & Opt & """ = ""txtChargeType"" Then "		& vbCr
	Response.Write "		.txtChargeType.focus() "							& vbCr
	Response.Write "	Elseif  """ & Opt & """ = ""txtProcessStep"" Then "		& vbCr
	Response.Write "		.txtProcessStep.focus() "							& vbCr	
	Response.Write "	Else  "								& vbCr
	Response.Write "		.txtBpCd.focus() "				& vbCr	
	Response.Write "	End If"								& vbCr
	Response.Write "End With"								& vbCr
	Response.Write "</Script>"								& vbCr
End Sub

%>

<Script Language=vbscript>
    
    With Parent
         .ggoSpread.Source  = .frm1.vspdData
         .frm1.vspdData.Redraw = False
         .ggoSpread.SSShowData "<%=iTotstrData%>","F"                  '☜ : Display data
        
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows, .Parent.gCurrency,.GetKeyPos("A",2),"A","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows, .Parent.gCurrency,.GetKeyPos("A",9),"A","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows, .Parent.gCurrency,.GetKeyPos("A",10),"A","Q","X","X")
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows, .Parent.gCurrency,.GetKeyPos("A",11),"A","Q","X","X")
		
         .lgPageNo			=  "<%=lgPageNo%>"               '☜ : Next next data tag
         
         .frm1.hdnBizArea.value		= "<%=ConvSPChars(Request("txtBizArea"))%>"
         .frm1.hdnChargeType.value	= "<%=ConvSPChars(Request("txtChargeType"))%>"
         .frm1.hdnBpCd.value		= "<%=ConvSPChars(Request("txtBpCd"))%>"
         .frm1.hdnChargeFrDt.value	= "<%=ConvSPChars(Request("txtChargeFrDt"))%>"
         .frm1.hdnChargeToDt.value	= "<%=ConvSPChars(Request("txtChargeToDt"))%>"
         .frm1.hdnPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
         .frm1.hdnItemCd.value		= "<%=ConvSPChars(Request("txtItemCd"))%>"
         .frm1.hdnProcessStep.value	= "<%=ConvSPChars(Request("txtProcessStep"))%>"
         .frm1.hdnPoNo.value	    = "<%=ConvSPChars(Request("txtPoNo"))%>"
         
         .frm1.txtBizAreaNm.value			=  "<%=ConvSPChars(arrRsVal(0))%>"
         .frm1.txtChargeTypeNm.value		=  "<%=ConvSPChars(arrRsVal(1))%>" 	
  		 .frm1.txtProcessStepNm.value		=  "<%=ConvSPChars(arrRsVal(2))%>"
  		 .frm1.txtBpNm.value				=  "<%=ConvSPChars(arrRsVal(3))%>"
  		 .frm1.txtPlantNm.value				=  "<%=ConvSPChars(arrRsVal(4))%>" 	
  		 .frm1.txtItemNm.value				=  "<%=ConvSPChars(arrRsVal(5))%>" 	
  		 
         .DbQueryOk
         .frm1.vspdData.Redraw = True
	End with
</Script>	

<%
    Response.End												'☜: 비지니스 로직 처리를 종료함 
%>

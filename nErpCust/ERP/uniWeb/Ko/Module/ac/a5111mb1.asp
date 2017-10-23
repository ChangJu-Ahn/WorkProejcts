<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("Q","*", "COOKIE", "QB")
Call LoadBNumericFormatB("Q", "*","NOCOOKIE","QB")

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3 , rs4, rs5    '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim txtFromGlDt		'format : YYYY-MM-DD
Dim txtToGlDt		'format : YYYY-MM-DD
Dim strFrGlDts		'format : YYYYMMDD
Dim strToGlDts		'format : YYYYMMDD

Dim txtAcctCd
Dim txtBizAreaCd
Dim txtBizAreaCd1
Dim txtSubLedger1
Dim txtSubLedger2
Dim txtMajorCd1
Dim txtMajorCd2
Dim Fiscyyyy,Fiscmm,Fiscdd
Dim Fiscyyyymm00, Fiscyyyymm01
Dim strGlDtYr, strGlDtMnth, strGlDtDt
Dim strGlDtYr1, strGlDtMnth1, strGlDtDt1

Dim txtTDrAmt
Dim txtTCrAmt
Dim txtTumAmt
Dim txtNDrAmt
Dim txtNCrAmt
Dim txtNSumAmt
Dim txtSDrAmt
Dim txtSCrAmt
Dim txtSSumAmt
Dim txtTSumAmt
Dim txtSumAmt

Dim lgAcctCd
Dim lgAcctNm
Dim lgBalFg

Dim lgBizAreaCd
Dim lgBizAreaNm
Dim lgBizAreaCd1
Dim lgBizAreaNm1
Dim lgGridDataExists
Dim iPrevEndRow
Dim iEndRow

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					

Const C_SHEETMAXROWS_D  = 100

    Call HideStatusWnd 

	txtFromGlDt		= Trim(Request("txtFromGlDt"))
	txtToGlDt		= Trim(Request("txtToGlDt"))	
	
	txtAcctCd		= Trim(Request("txtAcctCd"))
	txtBizAreaCd	= Trim(Request("txtBizAreaCd"))
	txtBizAreaCd1	= Trim(Request("txtBizAreaCd1"))
	txtSubLedger1	= Trim(Request("txtSubLedger1"))
	txtSubLedger2	= Trim(Request("txtSubLedger2"))
	txtMajorCd1		= Trim(Request("txtMajorCd1"))
	txtMajorCd2		= Trim(Request("txtMajorCd2"))

	'권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
    lgGridDataExists = "No"
    iPrevEndRow = 0
    iEndRow = 0
    
    gFiscStart = GetGlobalInf("gFiscStart")
        
    Call ExtractDateFrom(gFiscStart,gServerDateFormat,gServerDateType,Fiscyyyy,Fiscmm,Fiscdd)

    Call ExtractDateFrom(txtFromGlDt,gServerDateFormat,gServerDateType,strGlDtYr,strGlDtMnth,strGlDtDt)		
	strFrGlDts = 	strGlDtYr +  strGlDtMnth + strGlDtDt
	
	Call ExtractDateFrom(txtToGlDt,gServerDateFormat,gServerDateType,strGlDtYr1,strGlDtMnth1,strGlDtDt1)		
	strToGlDts = 	strGlDtYr1 +  strGlDtMnth1 + strGlDtDt1	
    
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
	On Error Resume Next
	Err.Clear
  
    If lgGridDataExists = "Yes" Then
		lgDataExist    = "Yes"
		lgstrData      = ""

		If CDbl(lgPageNo) > 0 Then
			iPrevEndRow = CDbl(C_SHEETMAXROWS_D) * CDbl(lgPageNo)    
			rs0.Move= iPrevEndRow                   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo	
		End If

		iLoopCount = -1
    
		Do While Not (rs0.EOF Or rs0.BOF)
		    iLoopCount =  iLoopCount + 1
		    iRowStr = ""
			For ColCnt = 0 To UBound(lgSelectListDT) - 1 
		        iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
			Next
 
		    If  iLoopCount < C_SHEETMAXROWS_D Then
		        lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
		    Else
		        lgPageNo = lgPageNo + 1
		        Exit Do
		    End If
		    rs0.MoveNext
		Loop

		If  iLoopCount < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
		    lgPageNo = ""                                                  '☜: 다음 데이타 없다.
            iEndRow = iPrevEndRow + iLoopCount + 1
		Else
			iEndRow = iPrevEndRow + iLoopCount            
		End If
  	
		rs0.Close
		Set rs0 = Nothing 
    End If

    If Not( rs1.EOF Or rs1.BOF) Then		
		txtTDrAmt = rs1(0)
		txtTCrAmt = rs1(1)		
	Else
		txtTDrAmt = 0
		txtTCrAmt = 0		
    End If
    
    rs1.Close
    Set rs1 = Nothing 
    
    If Not( rs2.EOF Or rs2.BOF) Then		
		txtNDrAmt = rs2(0)
		txtNCrAmt = rs2(1)		
	Else
		txtNDrAmt = 0
		txtNCrAmt = 0		
    End If
    
    rs2.Close
    Set rs2 = Nothing 
    
    If Not( rs3.EOF Or rs3.BOF) Then		
		lgAcctCd = rs3(0)
		lgAcctNm = rs3(1)
		lgBalFg	 = rs3(2)		
    End If
       
    rs3.Close
    Set rs3 = Nothing

    If Not( rs4.EOF Or rs4.BOF) Then		
		lgBizAreaCd = rs4(0)
		lgBizAreaNm = rs4(1)
    End If
    
    rs4.Close
    Set rs4= Nothing
    
    If Not( rs5.EOF Or rs5.BOF) Then		
		lgBizAreaCd1 = rs5(0)
		lgBizAreaNm1 = rs5(1)
    End If
    
    rs5.Close
    Set rs5= Nothing
End Sub


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	Redim UNISqlId(5)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(5,8)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "a5111MA101"
    UNISqlId(1) = "a5111MA102"
    UNISqlId(2) = "a5111MA103"   
    UNISqlId(3) = "A_GETACCT"
    UNISqlId(4) = "A_GETBIZ"
    UNISqlId(5) = "A_GETBIZ"
    
	Fiscyyyy =  strGlDtYr
	
	If Cint(Fiscmm) > Cint(strGlDtMnth)  Then                         ' 조회시작월이 당기 시작월보다작은 경우 전기 일자계산 
	   Fiscyyyymm00	= Cstr(Cint(strGlDtYr) - 1) & Fiscmm & "00"	
	   Fiscyyyymm01 = Cstr(Cint(strGlDtYr) - 1) & Fiscmm & "01"	  
	Else
	   Fiscyyyymm00	= Fiscyyyy & Fiscmm & "00"	 
	   Fiscyyyymm01	= Fiscyyyy & Fiscmm & "01"	   
	End If
   
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
	UNIValue(0,1)  = strFrGlDts
	UNIValue(0,2)  = strToGlDts
	UNIValue(0,3)  = FilterVar(txtAcctCd, "''", "S") 
	
	UNIValue(1,0)  = Fiscyyyymm00
	UNIValue(1,1)  = Fiscyyyymm01
	UNIValue(1,2)  = strFrGlDts
	UNIValue(1,3)  = FilterVar(txtAcctCd, "''", "S") 
	
	UNIValue(2,0)  = strFrGlDts
	UNIValue(2,1)  = strToGlDts
	UNIValue(2,2)  = FilterVar(txtAcctCd, "''", "S") 
	
	UNIValue(3,0)  = FilterVar(txtAcctCd, "''", "S") 
	
	If txtBizAreaCd = "" Then
	 	UNIValue(0,4)  = " and c.biz_area_cd >= " & FilterVar("0", "''", "S") & " "
	 	UNIValue(1,4)  = " and a.biz_area_cd >= " & FilterVar("0", "''", "S") & " "
	 	UNIValue(2,3)  = " and a.biz_area_cd >= " & FilterVar("0", "''", "S") & " "
	 	UNIValue(4,0)  = FilterVar("", "''", "S")  	 	
	Else
		UNIValue(0,4)  = " and c.biz_area_cd >= " & FilterVar(txtBizAreaCd, "''", "S") 
		UNIValue(1,4)  = " and a.biz_area_cd >= " & FilterVar(txtBizAreaCd, "''", "S") 
		UNIValue(2,3)  = " and a.biz_area_cd >= " & FilterVar(txtBizAreaCd, "''", "S")
		UNIValue(4,0)  =  FilterVar(txtBizAreaCd, "''", "S")  
	End If
	
	If txtBizAreaCd1 = "" Then
	 	UNIValue(0,7)  = "and c.biz_area_cd <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	 	UNIValue(1,7)  = "and a.biz_area_cd <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	 	UNIValue(2,6)  = "and a.biz_area_cd <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "	 	
	 	UNIValue(5,0)  = FilterVar("", "''", "S")  	 		 	
	Else
		UNIValue(0,7)  = "and c.biz_area_cd <= " & FilterVar(txtBizAreaCd1, "''", "S") 
		UNIValue(1,7)  = "and a.biz_area_cd <= " & FilterVar(txtBizAreaCd1, "''", "S") 
		UNIValue(2,6)  = "and a.biz_area_cd <= " & FilterVar(txtBizAreaCd1, "''", "S")
		UNIValue(5,0)  =  FilterVar(txtBizAreaCd1, "''", "S")  
	End If

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then			
	 	UNIValue(0,4)  = UNIValue(0,4) & " and c.biz_area_cd >= " & FilterVar(lgAuthBizAreaCd, "''", "S") & " "
	 	UNIValue(1,4)  = UNIValue(1,4) & " and a.biz_area_cd >= " & FilterVar(lgAuthBizAreaCd, "''", "S") & " "
	 	UNIValue(2,3)  = UNIValue(2,3) & " and a.biz_area_cd >= " & FilterVar(lgAuthBizAreaCd, "''", "S") & " "
	 	UNIValue(4,0)  = UNIValue(4,0) & " AND BIZ_AREA_CD LIKE " & FilterVar(lgAuthBizAreaCd, "''", "S")
	 		
		UNIValue(0,7)  = UNIValue(0,7) & " and c.biz_area_cd <= " & FilterVar(lgAuthBizAreaCd, "''", "S") 
		UNIValue(1,7)  = UNIValue(1,7) & " and a.biz_area_cd <= " & FilterVar(lgAuthBizAreaCd, "''", "S") 
		UNIValue(2,6)  = UNIValue(2,6) & " and a.biz_area_cd <= " & FilterVar(lgAuthBizAreaCd, "''", "S")
		UNIValue(5,0)  = UNIValue(5,0) & " AND BIZ_AREA_CD LIKE " & FilterVar(lgAuthBizAreaCd, "''", "S")  
	End If			
	
	If txtSubLedger1 = "" Then
	 	UNIValue(0,5)  = ""
	 	UNIValue(1,5)  = ""
	 	UNIValue(2,4)  = ""
	Else
		UNIValue(0,5)  = "and c.ctrl_val1 = " & FilterVar(txtSubLedger1, "''", "S")  
		UNIValue(1,5)  = "and a.ctrl_val1 = " & FilterVar(txtSubLedger1, "''", "S")  
		UNIValue(2,4)  = "and a.ctrl_val1 = " & FilterVar(txtSubLedger1, "''", "S")  
	End If

	If txtSubLedger2 = "" Then
	 	UNIValue(0,6)  = ""
	 	UNIValue(1,6)  = ""
	 	UNIValue(2,5)  = ""
	Else
		UNIValue(0,6)  = "and c.ctrl_val2 = " & FilterVar(txtSubLedger2, "''", "S") 
		UNIValue(1,6)  = "and a.ctrl_val2 = " & FilterVar(txtSubLedger2, "''", "S") 
		UNIValue(2,5)  = "and a.ctrl_val2 = " & FilterVar(txtSubLedger2, "''", "S") 
	End If
    
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))    
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
    
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
	Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 

	On Error Resume Next
	Err.Clear

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
		
        rs0.Close
        Set rs0 = Nothing	
        lgGridDataExists = "No"	
		Call  MakeSpreadSheetData()
    Else  
		lgGridDataExists = "Yes" 
		
        Call  MakeSpreadSheetData()
    End If
End Sub

%>

<Script Language=vbscript>
	Dim txtTDrAmt
	
	If "<%=lgDataExist%>" = "Yes" Then
		With Parent
		   'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			   .Frm1.hFromGlDt.value		= "<%=UNIDateClientFormat(txtFromGlDt)%>"
			   .Frm1.hToGlDt.value			= "<%=UNIDateClientFormat(txtToGlDt)%>"
			   .Frm1.hAcctCd.value			= "<%=ConvSPChars(txtAcctCd)%>"
			   .Frm1.hBizAreaCd.value		= "<%=ConvSPChars(txtBizAreaCd)%>"
			   .Frm1.hBizAreaCd1.value		= "<%=ConvSPChars(txtBizAreaCd1)%>"
			   .Frm1.hSubLedger1.value		= "<%=ConvSPChars(txtSubLedger1)%>"
			   .Frm1.hSubLedger2.value		= "<%=ConvSPChars(txtSubLedger2)%>"
			   .Frm1.hMajorCd1.value		= "<%=ConvSPChars(txtMajorCd1)%>"
			   .Frm1.hMajorCd2.value		= "<%=ConvSPChars(txtMajorCd2)%>"
			'msgbox Parent.Frm1.hFromGlDt.value
			End If
		   
		   'Show multi spreadsheet data from this line
		    .ggoSpread.Source  = Parent.frm1.vspdData
			.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '☜ : Display data       
			Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,"<%=gCurrency%>" ,.GetKeyPos("A",1),"A","Q","X","X")
			Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,"<%=gCurrency%>" ,.GetKeyPos("A",2),"A","Q","X","X")
			Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData ,<%=iPrevEndRow+1%>,<%=iEndRow%>,.GetKeyPos("A",6),.GetKeyPos("A",3),"A","Q","X","X")
			Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData ,<%=iPrevEndRow+1%>,<%=iEndRow%>,.GetKeyPos("A",6),.GetKeyPos("A",4),"A","Q","X","X")
			Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData ,<%=iPrevEndRow+1%>,<%=iEndRow%>,.GetKeyPos("A",6),.GetKeyPos("A",5),"D","Q","X","X")
			.frm1.vspdData.Redraw = True
		    .lgPageNo      =  "<%=lgPageNo%>"						'☜ : Next next data tag
		    .DbQueryOk
		    	
			.frm1.txtTDrAmt.text		= "<%=UNINumClientFormat(txtTDrAmt, ggAmtOfMoney.DecPoint,0)%>" 
			.frm1.txtTCrAmt.text		= "<%=UNINumClientFormat(txtTCrAmt, ggAmtOfMoney.DecPoint,0)%>"
					
			.frm1.txtNDrAmt.text		= "<%=UNINumClientFormat(txtNDrAmt, ggAmtOfMoney.DecPoint,0)%>"
			.frm1.txtNCrAmt.text		= "<%=UNINumClientFormat(txtNCrAmt, ggAmtOfMoney.DecPoint,0)%>"
			'text = 0인경우에 대비하자		
			<%
				txtSDrAmt	= Cdbl(txtTDrAmt) + Cdbl(txtNDrAmt)
				txtSCrAmt	= Cdbl(txtTCrAmt) + Cdbl(txtNCrAmt)
			%>

			.frm1.txtSDrAmt.text		= "<%=UNINumClientFormat(txtSDrAmt, ggAmtOfMoney.DecPoint,0)%>" 
			.frm1.txtSCrAmt.text		= "<%=UNINumClientFormat(txtSCrAmt, ggAmtOfMoney.DecPoint,0)%>" 
					
			If UCase(Trim("<%=lgBalFg%>")) = "DR" Then	
				<%
					txtTSumAmt	= Cdbl(txtTDrAmt) - Cdbl(txtTCrAmt)
					txtNSumAmt	= Cdbl(txtNDrAmt) - Cdbl(txtNCrAmt)
					txtSumAmt	= Cdbl(txtSDrAmt) - Cdbl(txtSCrAmt)
				%>

				.frm1.txtTSumAmt.text	= "<%=UNINumClientFormat(txtTSumAmt, ggAmtOfMoney.DecPoint,0)%>" 
				.frm1.txtNSumAmt.text	= "<%=UNINumClientFormat(txtNSumAmt, ggAmtOfMoney.DecPoint,0)%>" 
				.frm1.txtSSumAmt.text	= "<%=UNINumClientFormat(txtSumAmt, ggAmtOfMoney.DecPoint,0)%>" 
			Else
				<%
					txtTSumAmt	= Cdbl(txtTCrAmt) - Cdbl(txtTDrAmt)
					txtNSumAmt	= Cdbl(txtNCrAmt) - Cdbl(txtNDrAmt)
					txtSumAmt	= Cdbl(txtSCrAmt) - Cdbl(txtSDrAmt)
				%>
				.frm1.txtTSumAmt.text	= "<%=UNINumClientFormat(txtTSumAmt, ggAmtOfMoney.DecPoint,0)%>" 
				.frm1.txtNSumAmt.text	= "<%=UNINumClientFormat(txtNSumAmt, ggAmtOfMoney.DecPoint,0)%>" 
				.frm1.txtSSumAmt.text	= "<%=UNINumClientFormat(txtSumAmt, ggAmtOfMoney.DecPoint,0)%>"
			End If
  
			.frm1.txtAcctNm.value = "<%=ConvSPChars(lgAcctNm)%>"
			.frm1.txtBizAreaNm.value = "<%=ConvSPChars(lgBizAreaNm)%>"
			.frm1.txtBizAreaNm1.value = "<%=ConvSPChars(lgBizAreaNm1)%>"
		End With	
	End if
	
	'당기에 대한 데이타 존재하지 않아도(grid에 setting할 데이타 존재하지 않아도)이월금액에 대한 값은 존재할수 있으므로 
	'이월금액을 ma단에 넘겨준다.
	If "<%=lgGridDataExists%>" = "No" Then
		With Parent
			.frm1.txtTDrAmt.text		= "<%=UNINumClientFormat(txtTDrAmt, ggAmtOfMoney.DecPoint,0)%>" 
			.frm1.txtTCrAmt.text		= "<%=UNINumClientFormat(txtTCrAmt, ggAmtOfMoney.DecPoint,0)%>"
					
			.frm1.txtNDrAmt.text		= "<%=UNINumClientFormat(txtNDrAmt, ggAmtOfMoney.DecPoint,0)%>"
			.frm1.txtNCrAmt.text		= "<%=UNINumClientFormat(txtNCrAmt, ggAmtOfMoney.DecPoint,0)%>"
			'text = 0인경우에 대비하자		
		
			<%
				txtSDrAmt	= Cdbl(txtTDrAmt) + Cdbl(txtNDrAmt)
				txtSCrAmt	= Cdbl(txtTCrAmt) + Cdbl(txtNCrAmt)
			%>
		
			.frm1.txtSDrAmt.text		= "<%=UNINumClientFormat(txtSDrAmt, ggAmtOfMoney.DecPoint,0)%>" 
			.frm1.txtSCrAmt.text		= "<%=UNINumClientFormat(txtSCrAmt, ggAmtOfMoney.DecPoint,0)%>" 
					
			If UCase(Trim("<%=lgBalFg%>")) = "DR" Then	
				<%
					txtTSumAmt	= Cdbl(txtTDrAmt) - Cdbl(txtTCrAmt)
					txtNSumAmt	= Cdbl(txtNDrAmt) - Cdbl(txtNCrAmt)
					txtSumAmt	= Cdbl(txtSDrAmt) - Cdbl(txtSCrAmt)
				
				%>
								
				.frm1.txtTSumAmt.text	= "<%=UNINumClientFormat(txtTSumAmt, ggAmtOfMoney.DecPoint,0)%>" 
				.frm1.txtNSumAmt.text	= "<%=UNINumClientFormat(txtNSumAmt, ggAmtOfMoney.DecPoint,0)%>" 
				.frm1.txtSSumAmt.text	= "<%=UNINumClientFormat(txtSumAmt, ggAmtOfMoney.DecPoint,0)%>" 
			Else
		
				<%
					txtTSumAmt	= Cdbl(txtTCrAmt) - Cdbl(txtTDrAmt)
					txtNSumAmt	= Cdbl(txtNCrAmt) - Cdbl(txtNDrAmt)
					txtSumAmt	= Cdbl(txtSCrAmt) - Cdbl(txtSDrAmt)
				
				%>
				
				.frm1.txtTSumAmt.text	= "<%=UNINumClientFormat(txtTSumAmt, ggAmtOfMoney.DecPoint,0)%>" 
				.frm1.txtNSumAmt.text	= "<%=UNINumClientFormat(txtNSumAmt, ggAmtOfMoney.DecPoint,0)%>" 
				.frm1.txtSSumAmt.text	= "<%=UNINumClientFormat(txtSumAmt, ggAmtOfMoney.DecPoint,0)%>"
			End If
		
			.frm1.txtAcctNm.value = "<%=ConvSPChars(lgAcctNm)%>"
			.frm1.txtBizAreaNm.value = "<%=ConvSPChars(lgBizAreaNm)%>"
			.frm1.txtBizAreaNm1.value = "<%=ConvSPChars(lgBizAreaNm1)%>"
		End With
	
	End If
  

</Script>	


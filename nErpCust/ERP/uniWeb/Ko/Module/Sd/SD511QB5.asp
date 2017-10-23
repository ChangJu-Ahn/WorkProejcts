<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                          '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
    Call loadInfTB19029B("Q", "S","NOCOOKIE","QB")
    Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")
    Call LoadBasisGlobalInf()

'    On Error Resume Next

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7, rs8, rs9  '☜ : DBAgent Parameter 선언 
    Dim lgstrData															'☜ : data for spreadsheet data
    Dim lgTailList                                                          '☜ : Orderby절에 사용될 field 리스트 
    Dim lgSelectList,lgSelectList1
    Dim lgSelectListDT        
    Dim lgStrColorFlag
    Dim lgConFromDt
    Dim lgConToDt
    Dim lgBizAreaCd
    Dim lgSalesGrpCd
    Dim lgItemGrpCd
    Dim lgSoldToPartyCd
    Dim lgBillToPartyCd
    Dim lgPayerCd
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
    
    lgConFromDt		= Trim(Request("ConFromDt"))
    lgConToDt		= Trim(Request("ConToDt"))
    lgBizAreaCd		= Trim(Request("BizAreaCd"))
    lgSalesGrpCd	= Trim(Request("SalesGrpCd"))
    lgItemGrpCd		= Trim(Request("ItemGrpCd"))
    lgSoldToPartyCd	= Trim(Request("SoldToPartyCd"))
    lgBillToPartyCd	= Trim(Request("BillToPartyCd"))
    lgPayerCd		= Trim(Request("PayerCd"))
    lgPostFlag		= Trim(Request("PostFlag"))
    lgExceptFlag	= Trim(Request("ExceptFlag"))
    
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd
    
    lgSelectList1  = Request("lgSelectList1")                               '☜ : select 대상목록 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value

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
    Dim	 iBlnDisplayText

    lgstrData      = ""

    iLoopCount = 0
    iBlnDisplayText = False
    lgStrColorFlag = ""
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
        If rs0(0) = 0 Or iBlnDisplayText Then
			iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(0),rs0(0))
		    iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(1),rs0(1))
		Else
			iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(0),rs0(0))
		    iRowStr = iRowStr & Chr(11) & "년계"
		    iBlnDisplayText = True
		End If
		
		For ColCnt = 2 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
		
		If rs0(0) > 0 Then	'집계Row 여부 체크 
			lgStrColorFlag = lgStrColorFlag & CStr(iLoopCount) & gColSep & rs0(0) & gRowSep
		End If
		
        lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        
        rs0.MoveNext
	Loop  
	  	
	rs0.Close
    Set rs0 = Nothing 

End Sub


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	Dim iStrVal    
    
    Redim UNISqlId(7)                                       '☜: SQL ID 저장을 위한 영역확보    

    Redim UNIValue(7,1)                                     '⊙: DB-Agent로 전송될 parameter를 위한 변수 
 
               
    iStrVal	=	"SELECT BT.ITEM_GROUP_CD, YEAR(BH.BILL_DT) YR," & _
				"SUM(CASE WHEN MONTH(BH.BILL_DT) = 1 THEN BL.BILL_AMT_LOC + BL.VAT_AMT_LOC ELSE 0 END)	JAN, " & _
				"SUM(CASE WHEN MONTH(BH.BILL_DT) = 2 THEN BL.BILL_AMT_LOC + BL.VAT_AMT_LOC ELSE 0 END)	FEB, " & _
				"SUM(CASE WHEN MONTH(BH.BILL_DT) = 3 THEN BL.BILL_AMT_LOC + BL.VAT_AMT_LOC ELSE 0 END)	MAR, " & _ 
				"SUM(CASE WHEN MONTH(BH.BILL_DT) = 4 THEN BL.BILL_AMT_LOC + BL.VAT_AMT_LOC ELSE 0 END)	APR, " & _
				"SUM(CASE WHEN MONTH(BH.BILL_DT) = 5 THEN BL.BILL_AMT_LOC + BL.VAT_AMT_LOC ELSE 0 END)	MAY, " & _
				"SUM(CASE WHEN MONTH(BH.BILL_DT) = 6 THEN BL.BILL_AMT_LOC + BL.VAT_AMT_LOC ELSE 0 END)	JUN, " & _ 
				"SUM(CASE WHEN MONTH(BH.BILL_DT) = 7 THEN BL.BILL_AMT_LOC + BL.VAT_AMT_LOC ELSE 0 END)	JUL, " & _ 
				"SUM(CASE WHEN MONTH(BH.BILL_DT) = 8 THEN BL.BILL_AMT_LOC + BL.VAT_AMT_LOC ELSE 0 END)	AUG, " & _ 
				"SUM(CASE WHEN MONTH(BH.BILL_DT) = 9 THEN BL.BILL_AMT_LOC + BL.VAT_AMT_LOC ELSE 0 END)	SEP, " & _ 
				"SUM(CASE WHEN MONTH(BH.BILL_DT) = 10 THEN BL.BILL_AMT_LOC + BL.VAT_AMT_LOC ELSE 0 END)	OCT, " & _ 
				"SUM(CASE WHEN MONTH(BH.BILL_DT) = 11 THEN BL.BILL_AMT_LOC + BL.VAT_AMT_LOC ELSE 0 END)	NOV, " & _ 
				"SUM(CASE WHEN MONTH(BH.BILL_DT) = 12 THEN BL.BILL_AMT_LOC + BL.VAT_AMT_LOC ELSE 0 END)	DEC " & _ 
				"FROM	S_BILL_HDR BH INNER JOIN S_BILL_DTL BL ON (BH.BILL_NO = BL.BILL_NO) " & _ 
				"INNER JOIN B_ITEM BT ON (BL.ITEM_CD =  BT.ITEM_CD) " & _				  
 				"WHERE	BH.BILL_DT >=  " & FilterVar(UNIConvDate(lgConFromDt), "''", "S") & " AND " & _
 				"		BH.BILL_DT <=  " & FilterVar(UNIConvDate(lgConToDt), "''", "S") & ""
 				
  
	'영업그룹명===========================================================================================	
    If Len(lgSalesGrpCd) Then
		UNISqlId(1)		= "s0000qa005"	
		UNIValue(1,0)	= FilterVar(lgSalesGrpCd, "''", "S")

		iStrVal = iStrVal & " AND BH.SALES_GRP =  " & FilterVar(lgSalesGrpCd , "''", "S") & ""
	End If		

	'사업장명=============================================================================================    	
	If Len(lgBizAreaCd) Then
		UNISqlId(2)		= "s0000qa013"	
		UNIValue(2,0)	= FilterVar(lgBizAreaCd, "''", "S")

		iStrVal = iStrVal & " AND BH.BIZ_AREA =  " & FilterVar(lgBizAreaCd , "''", "S") & ""
	End If

    '주문처명=============================================================================================
    If Len(lgSoldToPartyCd) Then		
		UNISqlId(3)		= "s0000qa002"	
		UNIValue(3,0)	= FilterVar(lgSoldToPartyCd, "''", "S")

		iStrVal = iStrVal & " AND BH.SOLD_TO_PARTY =  " & FilterVar(lgSoldToPartyCd , "''", "S") & ""
	End If	
	
	'발행처명=============================================================================================
    If Len(lgBillToPartyCd) Then		
		UNISqlId(4)		= "s0000qa002"	
		UNIValue(4,0)	= FilterVar(lgBillToPartyCd, "''", "S")

		iStrVal = iStrVal & " AND BH.BILL_TO_PARTY =  " & FilterVar(lgBillToPartyCd , "''", "S") & ""
	End If

	'수금처명=============================================================================================	
    If Len(lgPayerCd) Then		
		UNISqlId(5)		= "s0000qa002"	
		UNIValue(5,0)	= FilterVar(lgPayerCd, "''", "S")

		iStrVal = iStrVal & " AND BH.PAYER =  " & FilterVar(lgPayerCd , "''", "S") & ""
	End If
	
	
	'품목그룹명===========================================================================================	
    If Len(lgItemGrpCd) Then		
		UNISqlId(6)		= "s0000qa028"	
		UNIValue(6,0)	= FilterVar(lgItemGrpCd, "''", "S")

		iStrVal = iStrVal & " AND BT.ITEM_GROUP_CD =  " & FilterVar(lgItemGrpCd , "''", "S") & ""
	End If
	
	'확정여부=============================================================================================	
    If Len(lgPostFlag) Then		
		iStrVal = iStrVal & " AND BH.POST_FLAG =  " & FilterVar(lgPostFlag , "''", "S") & ""
	End If
	
	'예외여부=============================================================================================	
    If Len(lgExceptFlag) Then		
		iStrVal = iStrVal & " AND BH.EXCEPT_FLAG =  " & FilterVar(lgExceptFlag , "''", "S") & ""
	End If
    
    iStrVal = iStrVal & " GROUP BY BT.ITEM_GROUP_CD, YEAR(BH.BILL_DT) "  
    
	UNISqlId(0) = "SD511QA501"					
    UNIValue(0,0) = Replace(lgSelectList, "?1", lgConToDt)   
    UNIValue(0,1) = iStrVal   
   
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    'UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                '☜: set ADO read mode
 
End Sub


'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    on error resume next
    Dim lgstrRetMsg                                                     '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF                                                           '☜ : ActiveX Data Factory 지정 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7)
    
    Set lgADF = Nothing													'☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If    
   
	Call BeginScriptTag()												'☜:Write the Script Tag "<Script language=vbscript>"
	
	'영업그룹 존재여부 
	If lgSalesGrpCd <> "" Then
		If rs1.EOF And rs1.BOF Then
			rs1.Close
			Set rs1 = Nothing			
			Call ConNotFound("txtConSalesGrpCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSalesGrpNm", rs1(1))		
		End If
	Else
		Call WriteConDesc("txtConSalesGrpNm", "")
	End If
	
	'사업장 존재여부 
	If lgBizAreaCd <> "" Then
		If rs2.EOF And rs2.BOF Then
			rs2.Close
			Set rs2 = Nothing			
			Call ConNotFound("txtConBizAreaCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConBizAreaNm", rs2(1))		
		End If
	Else
		Call WriteConDesc("txtConBizAreaNm", "")		
	End If

	'주문처 존재여부 
	If lgSoldToPartyCd  <> "" Then
		If rs3.EOF And rs3.BOF Then
			rs3.Close
			Set rs3 = Nothing			
			Call ConNotFound("txtConSoldToPartyCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSoldToPartyNm", rs3(1))		
		End If
	Else
		Call WriteConDesc("txtConSoldToPartyNm", "")
	End If
	
	'발행처 존재여부 
	If lgBillToPartyCd <> "" Then
		If rs4.EOF And rs4.BOF Then
			rs4.Close
			Set rs4 = Nothing			
			Call ConNotFound("txtConBillToPartyCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConBillToPartyNm", rs4(1))		
		End If
	Else
		Call WriteConDesc("txtConBillToPartyNm", "")
	End If
	
	'수금처 존재여부 
	If lgPayerCd <> "" Then
		If rs5.EOF And rs5.BOF Then
			rs5.Close
			Set rs5 = Nothing			
			Call ConNotFound("txtConPayerCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConPayerNm", rs5(1))		
		End If
	Else
		Call WriteConDesc("txtConPayerNm", "")
	End If
	
	'품목그룹 존재여부 
	If lgItemGrpCd <> "" Then
		If rs6.EOF And rs6.BOF Then
			rs6.Close
			Set rs6 = Nothing			
			Call ConNotFound("txtConItemGrpCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConItemGrpNm", rs6(1))		
		End If
	Else
		Call WriteConDesc("txtConItemGrpNm", "")
	End If
			 
    If  rs0.EOF And rs0.BOF Then	
        rs0.Close
        Set rs0 = Nothing
        Call DataNotFound("txtConYMFromDt")	
        Exit Sub
    Else    
        Call MakeSpreadSheetData()
        Call WriteResult()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Write the Result
'----------------------------------------------------------------------------------------------------------
Sub BeginScriptTag()
	Response.Write "<Script language=VBScript> " & VbCr
End Sub

Sub EndScriptTag()
	Response.Write "</Script> " & VbCr
End Sub

' 데이터가 존재하지 않는 경우 처리 Script 작성(조회조건 포함)
Sub ConNotFound(ByVal pvStrField)
	Response.Write " Call Parent.DisplayMsgBox(""970000"", ""X"", parent.frm1." & pvStrField & ".alt, ""X"") " & VbCr
	Response.Write " Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' 조회조건에 해당하는 명을 Display하는 Script 작성 
Sub WriteConDesc(ByVal pvStrField, Byval pvStrFieldDesc)
	Response.Write " Parent.frm1." & pvStrField & ".value = """ & ConvSPChars(pvStrFieldDesc) & """" &VbCr
End Sub

' 데이터가 존재하지 않는 경우 처리 Script 작성 
Sub DataNotFound(ByVal pvStrField)
	Response.Write " Call Parent.DisplayMsgBox(""900014"", ""X"", ""X"", ""X"") " & VbCr
	Response.Write " Call parent.SetFocusToDocument(""M"") " & vbCr	
	Response.Write " Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' 조회 결과를 Display하는 Script 작성 
Sub WriteResult()
	Response.Write " Parent.ggoSpread.Source  = Parent.frm1.vspdData " & vbCr
	Response.Write " Parent.frm1.vspdData.Redraw = False " & vbCr      	
	Response.Write " Parent.ggoSpread.SSShowData  """ & lgstrData & """ ,""F""" & vbCr
	Response.Write " parent.lgStrColorFlag = """ & lgStrColorFlag & """" & vbCr	
	Response.Write " Parent.DbQueryOk " & vbCr		
 	Response.Write " Parent.frm1.vspdData.Redraw = True " & vbCr      
	Call EndScriptTag()
End Sub

%>



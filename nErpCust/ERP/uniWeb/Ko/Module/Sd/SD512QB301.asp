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

    On Error Resume Next

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7, rs8, rs9  '☜ : DBAgent Parameter 선언 
    Dim lgstrData															'☜ : data for spreadsheet data
    Dim lgTailList                                                          '☜ : Orderby절에 사용될 field 리스트 
    Dim lgSelectList
    Dim lgSelectListDT        
    Dim lgStrColorFlag
    
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
	Dim lgConFromDt
	Dim lgConToDt
    Dim lgBizAreaCd
    Dim lgSalesGrpCd
	Dim lgBillToPartyCd
	Dim lgTaxBizCd
	Dim lgBillConfFlag
	Dim lgExceptFlag
	
	lgConFromDt			= uniConvDate(Trim(Request("ConFromDt")))
	lgConToDt			= uniConvDate(Trim(Request("ConToDt")))
    lgBizAreaCd			= Trim(Request("BizAreaCd"))
    lgSalesGrpCd		= Trim(Request("SalesGrpCd"))
    lgBillToPartyCd		= Trim(Request("BillToPartyCd"))
    lgTaxBizCd			= Trim(Request("TaxBizCd"))
    lgBillConfFlag		= Trim(Request("BillConfFlag"))
    lgExceptFlag		= Trim(Request("ExceptFlag"))
    
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd
    
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
    
    Const C_SHEETMAXROWS_D = 20     

    lgstrData      = ""
    
    iLoopCount = 0
    lgStrColorFlag = ""
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
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
    
    Redim UNISqlId(4)                                       '☜: SQL ID 저장을 위한 영역확보    

    Redim UNIValue(4,3)                                     '⊙: DB-Agent로 전송될 parameter를 위한 변수 
               
    iStrVal = ""	
	
	'사업장명=============================================================================================    	
	If Len(lgBizAreaCd) Then
		UNISqlId(1)		= "s0000qa013"	
		UNIValue(1,0)	= FilterVar(lgBizAreaCd, "''", "S")
		
		iStrVal	= iStrVal & " AND BIZ_AREA =  " & FilterVar(lgBizAreaCd , "''", "S") & ""				
	End If

	'영업그룹명===========================================================================================	
    If Len(lgSalesGrpCd) Then
		UNISqlId(2)		= "s0000qa005"	
		UNIValue(2,0)	= FilterVar(lgSalesGrpCd, "''", "S")
		
		iStrVal	= iStrVal & " AND SALES_GRP =  " & FilterVar(lgSalesGrpCd , "''", "S") & ""				
	End If		

	'발행처명=============================================================================================
    If Len(lgBillToPartyCd) Then		
		UNISqlId(3)		= "s0000qa002"	
		UNIValue(3,0)	= FilterVar(lgBillToPartyCd, "''", "S")
		
		iStrVal = iStrVal & " AND BILL_TO_PARTY =  " & FilterVar(lgBillToPartyCd , "''", "S") & ""				
	End If

	'세금신고사업장=============================================================================================
    If Len(lgTaxBizCd) Then		
		UNISqlId(4)		= "S0000QA501"	
		UNIValue(4,0)	= FilterVar(lgTaxBizCd, "''", "S")
		iStrVal = iStrVal & " AND TAX_BIZ_AREA =  " & FilterVar(lgTaxBizCd , "''", "S") & ""				
	End If
	
	'확정여부 =============================================================================================
    If Len(lgBillConfFlag) Then		
		iStrVal = iStrVal & " AND POST_FLAG =  " & FilterVar(lgBillConfFlag , "''", "S") & ""				
	End If

    If Len(lgExceptFlag) Then		
		iStrVal = iStrVal & " AND EXCEPT_FLAG =  " & FilterVar(lgExceptFlag , "''", "S") & ""				
	End If

	UNISqlId(0) = "SD512QA301"					
    UNIValue(0,0) = lgSelectList
	'매출채권일자=============================================================================================    	
	UNIValue(0,1) = " " & FilterVar(lgConFromDt, "''", "S") & ""				
	UNIValue(0,2) = " " & FilterVar(lgConToDt, "''", "S") & ""				
    UNIValue(0,3) = iStrVal
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
'    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)
    
    Set lgADF = Nothing													'☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If    
   
	Call BeginScriptTag()												'☜:Write the Script Tag "<Script language=vbscript>"
	
	'사업장 존재여부 
	If lgBizAreaCd <> "" Then
		If rs1.EOF And rs1.BOF Then
			rs1.Close
			Set rs1 = Nothing			
			Call ConNotFound("txtConBizAreaCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConBizAreaNm", rs1(1))		
		End If
	Else
		Call WriteConDesc("txtConBizAreaNm", "")		
	End If

	'영업그룹 존재여부 
	If lgSalesGrpCd <> "" Then
		If rs2.EOF And rs2.BOF Then
			rs2.Close
			Set rs2 = Nothing			
			Call ConNotFound("txtConSalesGrpCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSalesGrpNm", rs2(1))		
		End If
	Else
		Call WriteConDesc("txtConSalesGrpNm", "")		
	End If

	'발행처 존재여부 
	If lgBillToPartyCd <> "" Then
		If rs3.EOF And rs3.BOF Then
			rs3.Close
			Set rs3 = Nothing			
			Call ConNotFound("txtConBillToPartyCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConBillToPartyNm", rs3(1))		
		End If
	Else
		Call WriteConDesc("txtConBillToPartyNm", "")		
	End If

	'세금신고사업장 존재여부 
	If lgTaxBizCd <> "" Then
		If rs4.EOF And rs4.BOF Then
			rs4.Close
			Set rs4 = Nothing			
			Call ConNotFound("txtConTaxBizCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConTaxBizNm", rs4(1))		
		End If
	Else
		Call WriteConDesc("txtConTaxBizNm", "")		
	End If

    If  rs0.EOF And rs0.BOF Then	
        rs0.Close
        Set rs0 = Nothing
        Call DataNotFound("txtConFromDt")	
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



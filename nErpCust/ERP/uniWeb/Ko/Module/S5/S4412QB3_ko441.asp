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

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3			'☜ : DBAgent Parameter 선언 
    Dim lgstrData															'☜ : data for spreadsheet data
    Dim lgTailList                                                          '☜ : Orderby절에 사용될 field 리스트 
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgPageNo
    
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------

    Dim iPrevEndRow
    
    Dim lgFromDt			'조회기간시작 
    Dim lgToDt				'조회기간끝 
    Dim lgBizAreaCd			'사업장 
    Dim lgDnTypeCd			'출하형태 
    Dim lgRdoFlag			'조회구분 (Y:미매출,N:미확정)
    
    lgFromDt		= Trim(Request("txtHConFromDt"))    
    lgToDt			= Trim(Request("txtHConToDt"))
    lgBizAreaCd		= Trim(Request("txtHConBizArea"))
    lgDnTypeCd		= Trim(Request("txtHConDnType"))
    lgRdoFlag		= Trim(Request("txtHConRdoFlag"))
    txtBizarea = Trim(Request("txtBizarea"))

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd

    lgPageNo       = UNICInt(Trim(Request("txtHlgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("txtHlgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("txtHlgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("txtHlgTailList")                                 '☜ : Orderby value
    iPrevEndRow = 0

    Call FixUNISQLData()
    Call QueryData()
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    Const C_SHEETMAXROWS_D = 100     

    lgstrData      = ""
    iPrevEndRow = 0
    
    If CInt(lgPageNo) > 0 Then
       iPrevEndRow = C_SHEETMAXROWS_D * CInt(lgPageNo)
       rs0.Move  = iPrevEndRow    

    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
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

    If  iLoopCount < C_SHEETMAXROWS_D Then                             '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    
End Sub
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	Dim iStrVal    
    Dim iStrVal2
    
    Redim UNISqlId(3)                                       '☜: SQL ID 저장을 위한 영역확보    

    Redim UNIValue(3,4)                                     '⊙: DB-Agent로 전송될 parameter를 위한 변수 
               
    iStrVal = "WHERE"
	
	'조회기간시작=============/ 'Y' 미매출, 'N' 미확정 /==================================================
	If Len(lgFromDt) Then
		iStrVal = iStrVal & " HDR.ACTUAL_GI_DT >=  " & FilterVar(UNIConvDate(lgFromDt), "''", "S") & ""			
		If lgRdoFlag = "N" Then
			iStrVal = iStrVal & " AND BHDR.BILL_DT >=  " & FilterVar(UNIConvDate(lgFromDt), "''", "S") & ""		
		End If		
	End If		
	
	'조회기간끝===========================================================================================
	If Len(lgToDt) Then
		iStrVal = iStrVal & " AND HDR.ACTUAL_GI_DT <=  " & FilterVar(UNIConvDate(lgToDt), "''", "S") & ""		
		If lgRdoFlag = "N" Then
			iStrVal = iStrVal & " AND BHDR.BILL_DT  <=  " & FilterVar(UNIConvDate(lgToDt), "''", "S") & ""		
		End If
	End If
	
	'call svrmsgbox(txtBizarea,0,1)

	If Len(Trim(Request("txtBizarea"))) Then
		iStrVal = iStrVal & " AND HDR.BIZ_AREA =  " & FilterVar(Request("txtBizarea"), "''", "S") & " "	
		If lgRdoFlag = "N" Then
			iStrVal = iStrVal & " AND BHDR.BIZ_AREA =  " & FilterVar(Request("txtBizarea"), "''", "S") & " "		
		End If		
	End If		
	
	'사업장명=============================================================================================    	
	If Len(lgBizAreaCd) Then
		UNISqlId(1)		= "s0000qa013"	
		UNIValue(1,0)	= FilterVar(lgBizAreaCd, "''", "S") 
		UNIValue(0,2)	=  " " & FilterVar(lgBizAreaCd, "''", "S") & ""
	Else
		UNIValue(0,2)	= "NULL"
	End If
	
	'출하형태명===========================================================================================	
    If Len(lgDnTypeCd) Then		    
		UNISqlId(2)		= "s0000qa000"	
		UNIValue(2,0)	= FilterVar("I0001", "''", "S") 
		UNIValue(2,1)	= FilterVar(lgDnTypeCd, "''", "S") 
		UNIValue(0,3)	= " " & FilterVar(lgDnTypeCd, "''", "S") & ""
	Else
		UNIValue(0,3)	= "NULL"
	End If


	'조회구분 ======/'Y' 미매출, 'N' 미확정/===============================================================	
	If lgRdoFlag = "Y" Then	
		UNISqlId(0) = "S4412QA301"					
		' for summary 
		UNISqlId(3) = "S4412QA303"					
		iStrVal2 =			  " SUM(ISNULL(CASE WHEN DTL.RET_TYPE = '' THEN DTL.GI_QTY ELSE DTL.GI_QTY* (-1) END, 0)) AS TOTAL_GI_QTY, "
		iStrVal2 = iStrVal2 & " (SUM(ISNULL(CASE WHEN DTL.RET_TYPE = '' THEN DTL.GI_AMT_LOC ELSE DTL.GI_AMT_LOC* (-1) END,0) + ISNULL(CASE WHEN DTL.RET_TYPE = '' THEN DTL.VAT_AMT_LOC ELSE DTL.VAT_AMT_LOC* (-1) END,0))) AS GI_TOTAL_AMT, "
		iStrVal2 = iStrVal2 & " 0 AS TOTAL_BILL_QTY, "
		iStrVal2 = iStrVal2 & " 0 AS BILL_TOTAL_AMT "
		UNIValue(3,0) = iStrVal2

	Else
		UNISqlId(0) = "S4412QA302"					
		' for summary 
		UNISqlId(3) = "S4412QA304"					
		iStrVal2 =			  " SUM(ISNULL(CASE WHEN DTL.RET_TYPE = '' THEN DTL.GI_QTY ELSE DTL.GI_QTY* (-1) END, 0)) AS TOTAL_GI_QTY, "
		iStrVal2 = iStrVal2 & " (SUM(ISNULL(CASE WHEN DTL.RET_TYPE = '' THEN DTL.GI_AMT_LOC ELSE DTL.GI_AMT_LOC* (-1) END,0) + ISNULL(CASE WHEN DTL.RET_TYPE = '' THEN DTL.VAT_AMT_LOC ELSE DTL.VAT_AMT_LOC* (-1) END,0))) AS GI_TOTAL_AMT, "
		iStrVal2 = iStrVal2 & " SUM(ISNULL(CASE WHEN DTL.RET_TYPE = '' THEN DTL.BILL_QTY ELSE DTL.BILL_QTY* (-1) END, 0)) AS TOTAL_BILL_QTY, "
		iStrVal2 = iStrVal2 & " (SUM(ISNULL(BDTL.BILL_AMT_LOC,0) + ISNULL(BDTL.VAT_AMT_LOC,0))) AS BILL_TOTAL_AMT "
		UNIValue(3,0) = iStrVal2

	End If


    UNIValue(0,0) = lgSelectList                                      
	UNIValue(0,1) = iStrVal	         
'    Call svrmsgbox(iStrVal,0,1) 
	Dim iLoop
	For iLoop = 1 To 3
		UNIValue(3,iLoop) = UNIValue(0,iLoop)	
	Next
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    on error resume next
    Dim lgstrRetMsg                                                     '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF                                                           '☜ : ActiveX Data Factory 지정 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
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
			Call ConNotFound("txtConBizArea")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConBizAreaNm", rs1(1))		
		End If
	Else
		Call WriteConDesc("txtConBizAreaNm", "")		
	End If
	
	'출하형태 존재여부 
	If lgDnTypeCd <> "" Then
		If rs2.EOF And rs2.BOF Then
			rs2.Close
			Set rs2 = Nothing			
			Call ConNotFound("txtConDnType")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConDnTypeNm", rs2(1))		
		End If
	Else
		Call WriteConDesc("txtConDnTypeNm", "")
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
	Response.Write " Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' 조회 결과를 Display하는 Script 작성 
Sub WriteResult()
	Response.Write " Parent.ggoSpread.Source  = Parent.frm1.vspdData " & vbCr
	Response.Write " Parent.frm1.vspdData.Redraw = False " & vbCr      	
	Response.Write " Parent.ggoSpread.SSShowDataByClip  """ & lgstrData & """ ,""F""" & vbCr
	Response.Write " Parent.lgPageNo = """ & lgPageNo & """" & vbCr

	Response.Write " Parent.frm1.txt_TOTAL_GI_QTY.text = """ & rs3(0) & """" & vbCr
	Response.Write " Parent.frm1.txt_TOTAL_GI_AMT.text = """ & rs3(1) & """" & vbCr
	Response.Write " Parent.frm1.txt_TOTAL_BILL_QTY.text = """ & rs3(2) & """" & vbCr
	Response.Write " Parent.frm1.txt_TOTAL_BILL_AMT.text = """ & rs3(3) & """" & vbCr
		
	Response.Write " Parent.DbQueryOk " & vbCr		
 	Response.Write " Parent.frm1.vspdData.Redraw = True " & vbCr      
	Call EndScriptTag()
End Sub

%>



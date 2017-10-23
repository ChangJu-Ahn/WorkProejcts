<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S3312QB3
'*  4. Program Name         : ���ֽ���(�ŷ�ó)
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/07/01
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kwakeunkyoung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                          '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
    Call loadInfTB19029B("Q", "S","NOCOOKIE","QB")
    Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")
    Call LoadBasisGlobalInf()

    On Error Resume Next

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5	'�� : DBAgent Parameter ���� 
    Dim lgstrData															'�� : data for spreadsheet data
    Dim lgTailList                                                          '�� : Orderby���� ���� field ����Ʈ 
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgPageNo
    
'--------------- ������ coding part(��������,Start)--------------------------------------------------------

    Dim iPrevEndRow
    
    Dim lgFromDt			'��ȸ�⵵ 
    Dim lgToDt				'��ȸ�⵵ - 1
    Dim lgSoldToParty		'�ŷ�ó 
    Dim lgSalesGrp			'�����׷� 
    Dim lgItemCd			'ǰ�� 
    Dim lgSoType			'�������� 
	Dim lgRdoFlag			'��ȸ����    

    
    lgFromDt		= Trim(Request("txtHConFromDt"))    
    lgToDt			= Trim(Request("txtHConToDt"))
    lgSoldToParty	= Trim(Request("txtHSoldToParty"))
    lgSalesGrp		= Trim(Request("txtHSalesGrp"))
    lgItemCd		= Trim(Request("txtHItemCd"))
    lgSoType		= Trim(Request("txtHSoType"))  
    lgRdoFlag		= Trim(Request("txtHConRdoFlag"))

'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd

    lgPageNo       = UNICInt(Trim(Request("txtHlgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("txtHlgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("txtHlgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("txtHlgTailList")                                 '�� : Orderby value
    iPrevEndRow = 0

    Call FixUNISQLData()
    Call QueryData()
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    Const C_SHEETMAXROWS_D = 50     

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

    If  iLoopCount < C_SHEETMAXROWS_D Then                             '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    
End Sub
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Dim iStrVal    
	Dim iStrVal2    
	Dim iStrVal3    
	Dim iStrVal4    
			    
    Redim UNISqlId(5)                                       '��: SQL ID ������ ���� ����Ȯ��    

    Redim UNIValue(5,5)                                     '��: DB-Agent�� ���۵� parameter�� ���� ���� 
               
    iStrVal = "WHERE"
    iStrVal3 = "WHERE"

	'��ȸ�⵵=============================================================================================
	If Len(lgFromDt) Then
		iStrVal = iStrVal & " YEAR(HDR.SO_DT) =  " & FilterVar((lgFromDt), "''", "S") & ""			
		iStrVal3 = iStrVal3 & " YEAR(HDR.SO_DT) =  " & FilterVar((lgToDt), "''", "S") & ""			
	End If		

	'�����׷��===========================================================================================	
    If Len(lgSalesGrp) Then		    
		UNISqlId(1)		= "S0000QA005"	
		UNIValue(1,0)	= FilterVar(lgSalesGrp, "''", "S")
		iStrVal = iStrVal & " AND HDR.SALES_GRP =  " & FilterVar(lgSalesGrp , "''", "S") & ""		
		iStrVal3 = iStrVal3 & " AND HDR.SALES_GRP =  " & FilterVar(lgSalesGrp , "''", "S") & ""		
	End If

	'�ŷ�ó��=============================================================================================    	
	If Len(lgSoldToParty) Then
		UNISqlId(2)		= "s0000qa002"	
		UNIValue(2,0)	= FilterVar(lgSoldToParty, "''", "S")
		iStrVal = iStrVal & " AND HDR.SOLD_TO_PARTY =  " & FilterVar(lgSoldToParty , "''", "S") & ""		
		iStrVal3 = iStrVal3 & " AND HDR.SOLD_TO_PARTY =  " & FilterVar(lgSoldToParty , "''", "S") & ""		
	End If
	
	'ǰ���===============================================================================================	
    If Len(lgItemCd) Then		    
		UNISqlId(3)		= "S0000QA001"	
		UNIValue(3,0)	= FilterVar(lgItemCd, "''", "S")
		iStrVal = iStrVal & " AND DTL.ITEM_CD =  " & FilterVar(lgItemCd , "''", "S") & ""		
		iStrVal3 = iStrVal3 & " AND DTL.ITEM_CD =  " & FilterVar(lgItemCd , "''", "S") & ""		
	End If

	'�������¸�===========================================================================================	
    If Len(lgSoType) Then		    
		UNISqlId(4)		= "S0000QA007"	
		UNIValue(4,0)	= FilterVar(lgSoType, "''", "S")
		iStrVal = iStrVal & " AND HDR.SO_TYPE =  " & FilterVar(lgSoType , "''", "S") & ""		
		iStrVal3 = iStrVal3 & " AND HDR.SO_TYPE =  " & FilterVar(lgSoType , "''", "S") & ""		
	End If

	'��ȸ����==/'Y'����,'N'�ݾ�/=========================================================================	
	If lgRdoFlag = "Y" Then		

		iStrVal2 =            " SUM(CASE WHEN MONTH(HDR.SO_DT) = 1 THEN ISNULL(DTL.SO_BASE_QTY,0) + ISNULL(DTL.BONUS_BASE_QTY,0) ELSE 0 END) JAN, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 2 THEN ISNULL(DTL.SO_BASE_QTY,0) + ISNULL(DTL.BONUS_BASE_QTY,0) ELSE 0 END) FEB, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 3 THEN ISNULL(DTL.SO_BASE_QTY,0) + ISNULL(DTL.BONUS_BASE_QTY,0) ELSE 0 END) MAR, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 4 THEN ISNULL(DTL.SO_BASE_QTY,0) + ISNULL(DTL.BONUS_BASE_QTY,0) ELSE 0 END) APR, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 5 THEN ISNULL(DTL.SO_BASE_QTY,0) + ISNULL(DTL.BONUS_BASE_QTY,0) ELSE 0 END) MAY, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 6 THEN ISNULL(DTL.SO_BASE_QTY,0) + ISNULL(DTL.BONUS_BASE_QTY,0) ELSE 0 END) JUN, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 7 THEN ISNULL(DTL.SO_BASE_QTY,0) + ISNULL(DTL.BONUS_BASE_QTY,0) ELSE 0 END) JUL, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 8 THEN ISNULL(DTL.SO_BASE_QTY,0) + ISNULL(DTL.BONUS_BASE_QTY,0) ELSE 0 END) AUG, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 9 THEN ISNULL(DTL.SO_BASE_QTY,0) + ISNULL(DTL.BONUS_BASE_QTY,0) ELSE 0 END) SEP, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 10 THEN ISNULL(DTL.SO_BASE_QTY,0) + ISNULL(DTL.BONUS_BASE_QTY,0) ELSE 0 END) OCT, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 11 THEN ISNULL(DTL.SO_BASE_QTY,0) + ISNULL(DTL.BONUS_BASE_QTY,0) ELSE 0 END) NOV, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 12 THEN ISNULL(DTL.SO_BASE_QTY,0) + ISNULL(DTL.BONUS_BASE_QTY,0) ELSE 0 END) DEC  "

		iStrVal4 =			  " ISNULL(SUM(ISNULL(DTL.SO_BASE_QTY,0) + ISNULL(DTL.BONUS_BASE_QTY,0)),0) LAST_YAER, "

		' for summary 
		UNIValue(5,0) = " ISNULL(SUM(ISNULL(CASE WHEN HDR.RET_ITEM_FLAG = " & FilterVar("Y", "''", "S") & "  THEN -DTL.SO_BASE_QTY ELSE DTL.SO_BASE_QTY END,0)) + SUM(ISNULL(CASE WHEN HDR.RET_ITEM_FLAG = " & FilterVar("Y", "''", "S") & "  THEN -DTL.BONUS_BASE_QTY ELSE DTL.BONUS_BASE_QTY END,0)),0) AS TOTAL_SO_QTY "
	Else

		iStrVal2 =            " SUM(CASE WHEN MONTH(HDR.SO_DT) = 1 THEN ISNULL(DTL.NET_AMT_LOC,0) + ISNULL(DTL.VAT_AMT_LOC,0) ELSE 0 END) JAN, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 2 THEN ISNULL(DTL.NET_AMT_LOC,0) + ISNULL(DTL.VAT_AMT_LOC,0) ELSE 0 END) FEB, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 3 THEN ISNULL(DTL.NET_AMT_LOC,0) + ISNULL(DTL.VAT_AMT_LOC,0) ELSE 0 END) MAR, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 4 THEN ISNULL(DTL.NET_AMT_LOC,0) + ISNULL(DTL.VAT_AMT_LOC,0) ELSE 0 END) APR, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 5 THEN ISNULL(DTL.NET_AMT_LOC,0) + ISNULL(DTL.VAT_AMT_LOC,0) ELSE 0 END) MAY, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 6 THEN ISNULL(DTL.NET_AMT_LOC,0) + ISNULL(DTL.VAT_AMT_LOC,0) ELSE 0 END) JUN, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 7 THEN ISNULL(DTL.NET_AMT_LOC,0) + ISNULL(DTL.VAT_AMT_LOC,0) ELSE 0 END) JUL, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 8 THEN ISNULL(DTL.NET_AMT_LOC,0) + ISNULL(DTL.VAT_AMT_LOC,0) ELSE 0 END) AUG, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 9 THEN ISNULL(DTL.NET_AMT_LOC,0) + ISNULL(DTL.VAT_AMT_LOC,0) ELSE 0 END) SEP, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 10 THEN ISNULL(DTL.NET_AMT_LOC,0) + ISNULL(DTL.VAT_AMT_LOC,0) ELSE 0 END) OCT, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 11 THEN ISNULL(DTL.NET_AMT_LOC,0) + ISNULL(DTL.VAT_AMT_LOC,0) ELSE 0 END) NOV, "
		iStrVal2 = iStrVal2 & " SUM(CASE WHEN MONTH(HDR.SO_DT) = 12 THEN ISNULL(DTL.NET_AMT_LOC,0) + ISNULL(DTL.VAT_AMT_LOC,0) ELSE 0 END) DEC  "

		iStrVal4 =			  " ISNULL(SUM(ISNULL(DTL.NET_AMT_LOC,0) + ISNULL(DTL.VAT_AMT_LOC,0)),0) LAST_YAER, "

		' for summary 
		UNIValue(5,0) = " ISNULL(SUM(ISNULL(CASE WHEN HDR.RET_ITEM_FLAG = " & FilterVar("Y", "''", "S") & "  THEN -DTL.NET_AMT_LOC ELSE DTL.NET_AMT_LOC END,0)) + SUM(ISNULL(CASE WHEN HDR.RET_ITEM_FLAG = " & FilterVar("Y", "''", "S") & "  THEN -DTL.VAT_AMT_LOC ELSE DTL.VAT_AMT_LOC END,0)),0) AS TOTAL_SO_AMT "
	End If

	UNISqlId(0) = "S3312QA301"					
	UNIValue(0,0) = lgSelectList                                      
	UNIValue(0,2) = iStrVal	         
	UNIValue(0,1) = iStrVal2	             
	UNIValue(0,4) = iStrVal3	             
	UNIValue(0,3) = iStrVal4	             

	' for summary 
	UNISqlId(5) = "S3312QA302"					
	UNIValue(5,1) = iStrVal	         

  
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    on error resume next
    Dim lgstrRetMsg                                                     '�� : Record Set Return Message �������� 
    Dim iStr
    Dim lgADF                                                           '�� : ActiveX Data Factory ���� �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)
    
    Set lgADF = Nothing													'��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If    
   
	Call BeginScriptTag()												'��:Write the Script Tag "<Script language=vbscript>"
	
	'�����׷� ���翩�� 
	If lgSalesGrp <> "" Then
		If rs1.EOF And rs1.BOF Then
			rs1.Close
			Set rs1 = Nothing			
			Call ConNotFound("txtSalesGrp")			
			Exit Sub
		Else	
			Call WriteConDesc("txtSalesGrpNm", rs1(1))		
		End If
	Else
		Call WriteConDesc("txtSalesGrpNm", "")
	End If


	'�ŷ�ó ���翩�� 
	If lgSoldToParty <> "" Then
		If rs2.EOF And rs2.BOF Then
			rs2.Close
			Set rs2 = Nothing			
			Call ConNotFound("txtSoldToParty")			
			Exit Sub
		Else	
			Call WriteConDesc("txtSoldToPartyNm", rs2(1))		
		End If
	Else
		Call WriteConDesc("txtSoldToPartyNm", "")		
	End If
	
 
	'ǰ�� ���翩�� 
	If lgItemCd <> "" Then
		If rs3.EOF And rs3.BOF Then
			rs3.Close
			Set rs3 = Nothing			
			Call ConNotFound("txtItemCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtItemNm", rs3(1))		
		End If
	Else
		Call WriteConDesc("txtItemNm", "")
	End If

	'�������� ���翩�� 
	If lgSoType <> "" Then
		If rs4.EOF And rs4.BOF Then
			rs4.Close
			Set rs4 = Nothing			
			Call ConNotFound("txtSoType")			
			Exit Sub
		Else	
			Call WriteConDesc("txtSoTypeNm", rs4(1))		
		End If
	Else
		Call WriteConDesc("txtSoTypeNm", "")
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

' �����Ͱ� �������� �ʴ� ��� ó�� Script �ۼ�(��ȸ���� ����)
Sub ConNotFound(ByVal pvStrField)
	Response.Write " Call Parent.DisplayMsgBox(""970000"", ""X"", parent.frm1." & pvStrField & ".alt, ""X"") " & VbCr
	Response.Write " Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' ��ȸ���ǿ� �ش��ϴ� ���� Display�ϴ� Script �ۼ� 
Sub WriteConDesc(ByVal pvStrField, Byval pvStrFieldDesc)
	Response.Write " Parent.frm1." & pvStrField & ".value = """ & ConvSPChars(pvStrFieldDesc) & """" &VbCr
End Sub

' �����Ͱ� �������� �ʴ� ��� ó�� Script �ۼ� 
Sub DataNotFound(ByVal pvStrField)
	Response.Write " Call Parent.DisplayMsgBox(""900014"", ""X"", ""X"", ""X"") " & VbCr
	Response.Write " Call parent.SetFocusToDocument(""M"") " & vbCr	
	Response.Write " Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' ��ȸ ����� Display�ϴ� Script �ۼ� 
Sub WriteResult()
	If lgRdoFlag = "Y" Then		
		Response.Write " Parent.ggoSpread.Source  = Parent.frm1.vspdData " & vbCr
		Response.Write " Parent.frm1.vspdData.Redraw = False " & vbCr      	
	Else
		Response.Write " Parent.ggoSpread.Source  = Parent.frm1.vspdData1 " & vbCr
		Response.Write " Parent.frm1.vspdData1.Redraw = False " & vbCr      	
	End If
	Response.Write " Parent.ggoSpread.SSShowDataByClip  """ & lgstrData & """ ,""F""" & vbCr
	Response.Write " Parent.lgPageNo = """ & lgPageNo & """" & vbCr
	Response.Write " Parent.frm1.txt_TOTAL.text = """ & rs5(0) & """" & vbCr
	Response.Write " Parent.DbQueryOk " & vbCr		
	If lgRdoFlag = "Y" Then		
		Response.Write " Parent.frm1.vspdData.Redraw = True " & vbCr      
	Else
		Response.Write " Parent.frm1.vspdData1.Redraw = True " & vbCr      
	End If
	Call EndScriptTag()
End Sub

%>



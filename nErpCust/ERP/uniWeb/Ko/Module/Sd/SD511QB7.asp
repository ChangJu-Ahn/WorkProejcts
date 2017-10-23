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

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7, rs8, rs9  '�� : DBAgent Parameter ���� 
    Dim lgstrData															'�� : data for spreadsheet data
    Dim lgTailList                                                          '�� : Orderby���� ���� field ����Ʈ 
    Dim lgSelectList,lgSelectList1
    Dim lgSelectListDT        
    Dim lgStrColorFlag
    Dim lgConDt
    Dim lgBizAreaCd
    Dim lgSalesGrpCd
    Dim lgItemGrpCd
    Dim lgSoldToPartyCd
    Dim lgBillToPartyCd
    Dim lgPayerCd
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
    
    lgConDt			= Trim(Request("ConDt"))
    
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd
    
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectList1  = Request("lgSelectList1")                                
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value

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
    Dim  iTmpCnt
    Const C_SHEETMAXROWS_D = 20     

    lgstrData      = ""

    iLoopCount = 0
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
		
        lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        
        rs0.MoveNext
	Loop
	  	
	    
	iTmpCnt=0
    Do while Not (rs1.EOF Or rs1.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
	        iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs1(ColCnt))
		Next
	
        lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
       
		lgStrColorFlag = lgStrColorFlag & CStr(iLoopCount) & gColSep & "1" & gRowSep
        rs1.MoveNext
	Loop
	  	
	rs0.Close
    Set rs0 = Nothing 

	rs1.Close
    Set rs1 = Nothing 

End Sub


'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Dim iStrVal    
    Dim iStrSql
    Redim UNISqlId(1)                                       '��: SQL ID ������ ���� ����Ȯ��    

    Redim UNIValue(1,2)                                     '��: DB-Agent�� ���۵� parameter�� ���� ���� 
 
	iStrSql=""
	iStrSql = iStrSql & " SELECT  AR.BIZ_AREA_CD BIZ_AREA_CD, SA.BIZ_AREA_NM BIZ_AREA_NM,"
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 1 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END)	JAN, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 2 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END)	FEB, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 3 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END)	MAR, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 4 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END)	APR, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 5 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END)	MAY, " 
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 6 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END)	JUN, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 7 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END)	JUL, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 8 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END)	AUG, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 9 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END)	SEP, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 10 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END) OCT, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 11 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END) NOV, " 
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 12 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END) DEC"
	iStrSql = iStrSql & " FROM	A_OPEN_AR AR, "
	iStrSql = iStrSql & " A_CLS_AR CLS, "
	iStrSql = iStrSql & " A_AR_ADJUST ADJ,"
	iStrSql = iStrSql & " B_BIZ_AREA SA"
	iStrSql = iStrSql & " WHERE  YEAR(AR.AR_DT) =  " & FilterVar(lgConDt , "''", "S") & ""
	iStrSql = iStrSql & " AND    AR.AR_NO *= CLS_AR_NO"
	iStrSql = iStrSql & " AND    MONTH(AR.AR_DT) *= MONTH(CLS.CLS_DT)"
	iStrSql = iStrSql & " AND    AR.AR_NO *= ADJ.AR_NO"
	iStrSql = iStrSql & " AND    MONTH(AR.AR_DT) *= MONTH(ADJ.ADJUST_DT)"
	iStrSql = iStrSql & " AND    AR.BIZ_AREA_CD = SA.BIZ_AREA_CD"
	iStrSql = iStrSql & " GROUP BY AR.BIZ_AREA_CD, SA.BIZ_AREA_NM" 

	UNISqlId(0) = "SD511QA701"					
    UNIValue(0,0) = lgSelectList   
    UNIValue(0,1) = iStrSql
    UNIValue(0,2) = " " & FilterVar(cLng(lgConDt)-1, "''", "S") & ""

	iStrSql = ""
	iStrSql = iStrSql & " SELECT  AR.BIZ_AREA_CD BIZ_AREA_CD,"
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 1 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END) JAN, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 2 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END) FEB, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 3 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END) MAR, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 4 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END) APR, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 5 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END) MAY, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 6 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END) JUN, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 7 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END) JUL, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 8 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END) AUG, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 9 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END) SEP, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 10 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END) OCT, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 11 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END) NOV, "
	iStrSql = iStrSql & " SUM(CASE WHEN MONTH(AR.AR_DT) = 12 THEN ISNULL(CLS.CLS_LOC_AMT,0) + ISNULL(ADJ.ADJUST_LOC_AMT,0) ELSE 0 END) DEC"
	iStrSql = iStrSql & " FROM	A_OPEN_AR AR, "
	iStrSql = iStrSql & " A_CLS_AR CLS, "
	iStrSql = iStrSql & " A_AR_ADJUST ADJ"
	iStrSql = iStrSql & " WHERE	YEAR(AR.AR_DT) =  " & FilterVar(lgConDt , "''", "S") & ""
	iStrSql = iStrSql & " AND	AR.AR_NO *= CLS_AR_NO"
	iStrSql = iStrSql & " AND	MONTH(AR.AR_DT) *= MONTH(CLS.CLS_DT)"
	iStrSql = iStrSql & " AND	AR.AR_NO *= ADJ.AR_NO"
	iStrSql = iStrSql & " AND	MONTH(AR.AR_DT) *= MONTH(ADJ.ADJUST_DT)"
	iStrSql = iStrSql & " GROUP BY AR.BIZ_AREA_CD"
	
	UNISqlId(1) = "SD511QA702"					
    UNIValue(1,0) = lgSelectList1   
    UNIValue(1,1) = iStrSql   
    UNIValue(1,2) = " " & FilterVar(cLng(lgConDt)-1, "''", "S") & ""
	
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNILock = DISCONNREAD :	UNIFlag = "1"                                '��: set ADO read mode
 
End Sub


'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    on error resume next
    Dim lgstrRetMsg                                                     '�� : Record Set Return Message �������� 
    Dim iStr
    Dim lgADF                                                           '�� : ActiveX Data Factory ���� �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7)
    
    Set lgADF = Nothing													'��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If    
   
	Call BeginScriptTag()												'��:Write the Script Tag "<Script language=vbscript>"
	
    If  rs0.EOF And rs0.BOF Then	
        rs0.Close
        Set rs0 = Nothing
        Call DataNotFound("txtConYYYYDt")	
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
	Response.Write " Parent.frm1." & pvStrField & ".focus " & VbCr
	Call EndScriptTag()
End Sub

' ��ȸ ����� Display�ϴ� Script �ۼ� 
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



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
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgStrColorFlag
    Dim lgPageNo
    
'--------------- ������ coding part(��������,Start)--------------------------------------------------------

    Dim iPrevEndRow
    
	Dim lgFromDt
	Dim lgToDt
	Dim lgSalesGrpCd1
	Dim lgSalesGrpCd2
	Dim lgBpCd1
	Dim lgBpCd2
	Dim lgBillTypeCd
	Dim lgBillConfFlag
	Dim lgExceptFlag
    lgFromDt		= Trim(Request("txtHdnConFrDt"))    
    lgToDt			= Trim(Request("txtHdnConToDt"))
    lgSalesGrpCd1	= Trim(Request("txtHdnConSalesGrpCd"))
    lgBpCd1			= Trim(Request("txtHdnConBpCd"))
    lgBillTypeCd	= Trim(Request("txtHdnConBillTypeCd"))
    lgBillConfFlag	= Trim(Request("txtHdnConBillConfFlag"))
    lgExceptFlag	= Trim(Request("txtHdnExceptFlag"))
    
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd

    lgPageNo       = UNICInt(Trim(Request("txtHdnPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("txtHdnlgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("txtHdnlgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("txtHdnlgTailList")                                 '�� : Orderby value
    iPrevEndRow = 0

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
    Dim  iMaxRsCnt
    
    Const C_SHEETMAXROWS_D = 20     

    lgstrData      = ""

    iLoopCount = 0
    lgStrColorFlag = ""    
    
    iMaxRsCnt = rs0.recordcount
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
			If iLoopCount = iMaxRsCnt and isnumeric(rs0(0)) and ColCnt < 1 Then
				iRowStr = iRowStr & Chr(11) & "�հ�"		
			Else
				iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
			End If
		Next
		'COLOR
 		If isnumeric(rs0(0)) Then
 			If  rs0(0) > 0 Then
				lgStrColorFlag = lgStrColorFlag & CStr(iLoopCount) & gColSep & rs0(0) & gRowSep
			End If
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

	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Dim iStrVal    
    
    Redim UNISqlId(5)                                       '��: SQL ID ������ ���� ����Ȯ��    

    Redim UNIValue(5,6)                                     '��: DB-Agent�� ���۵� parameter�� ���� ���� 
               
	'�����׷����=========================================================================================
	If Len(lgSalesGrpCd1) Then
		UNISqlId(1)		= "s0000qa005"	
		UNIValue(1,0)	= FilterVar(lgSalesGrpCd1, "''", "S")

		iStrVal = iStrVal & " AND SALES_GRP =  " & FilterVar(lgSalesGrpCd1 , "''", "S") & ""		
	End If		
	
	'�ŷ�ó����=========================================================================================
	If Len(lgBpCd1) Then
		UNISqlId(3)		= "s0000qa002"	
		UNIValue(3,0)	= FilterVar(lgBpCd1, "''", "S")
		
		iStrVal = iStrVal & " AND SOLD_TO_PARTY =  " & FilterVar(lgBpCd1 , "''", "S") & ""		
	End If		
	
	'����ä�����¸�=============================================================================================    	
	If Len(lgBillTypeCd) Then
		UNISqlId(5)		= "s0000qa011"	
		UNIValue(5,0)	= FilterVar(lgBillTypeCd, "''", "S")

		iStrVal = iStrVal & " AND BILL_TYPE =  " & FilterVar(lgBillTypeCd , "''", "S") & ""				
	End If
	
	'Ȯ������=============================================================================================    	
	If Len(lgBillConfFlag) Then
		iStrVal = iStrVal & " AND POST_FLAG =  " & FilterVar(lgBillConfFlag , "''", "S") & ""				
	End If
	
	'���ܿ���=============================================================================================    	
	If Len(lgExceptFlag) Then
		iStrVal = iStrVal & " AND EXCEPT_FLAG =  " & FilterVar(lgExceptFlag , "''", "S") & ""				
	End If

	UNISqlId(0) = "SD513QA501"					
    UNIValue(0,0) = lgSelectList                                      
	UNIValue(0,1) = " " & FilterVar(UNIConvDate(lgFromDt), "''", "S") & ""		
	UNIValue(0,2) = " " & FilterVar(UNIConvDate(lgToDt), "''", "S") & ""		
	UNIValue(0,3) = iStrVal	         
	UNIValue(0,4) = " " & FilterVar(UNIConvDate(lgFromDt), "''", "S") & ""		
	UNIValue(0,5) = " " & FilterVar(UNIConvDate(lgToDt), "''", "S") & ""		
	UNIValue(0,6) = iStrVal	         
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
'    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7, rs8, rs9)
    
    Set lgADF = Nothing													'��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If    
   
	Call BeginScriptTag()												'��:Write the Script Tag "<Script language=vbscript>"

	Response.Write "parent.frm1.txtHConFromDt.value			=""" & lgFromDt & """" & vbcr
	Response.Write "parent.frm1.txtHConFromDt.value			=""" & lgFromDt	& """" & vbcr
	Response.Write "parent.frm1.txtHConToDt.value			=""" & lgToDt & """" & vbcr
	Response.Write "parent.frm1.txtHConSalesGrpCd1.value	=""" & lgSalesGrpCd1 & """" & vbcr
	Response.Write "parent.frm1.txtHConBpCd1.value			=""" & lgBpCd1 & """" & vbcr
	Response.Write "parent.frm1.txtHConBillTypeCd.value		=""" & lgBillTypeCd & """" & vbcr
	Response.Write "parent.frm1.rdoHBillConfFlag.value		=""" & lgBillConfFlag & """" & vbcr
	Response.Write "parent.frm1.rdoHExceptFlag.value		=""" & lgExceptFlag & """" & vbcr
	
	'�����׷�#1 ���翩�� 
	If lgSalesGrpCd1 <> "" Then
		If rs1.EOF And rs1.BOF Then
			rs1.Close
			Set rs1 = Nothing			
			Call ConNotFound("txtConSalesGrpCd1")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSalesGrpNm1", rs1(1))		
		End If
	Else
		Call WriteConDesc("txtConSalesGrpNm1", "")
	End If
	
	'�ŷ�ó#1 ���翩�� 
	If lgBpCd1 <> "" Then
		If rs3.EOF And rs3.BOF Then
			rs3.Close
			Set rs3 = Nothing			
			Call ConNotFound("txtConBpCd1")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConBpNm1", rs3(1))		
		End If
	Else
		Call WriteConDesc("txtConBpNm1", "")		
	End If
	
	'����ä������ ���翩�� 
	If lgBillTypeCd <> "" Then
		If rs5.EOF And rs5.BOF Then
			rs5.Close
			Set rs5 = Nothing			
			Call ConNotFound("txtConBillTypeCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConBillTypeNm", rs5(1))		
		End If
	Else
		Call WriteConDesc("txtConBillTypeNm", "")
	End If
	
	'
    If  rs0.RecordCount<=1 Then	
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
	Response.Write " Parent.ggoSpread.Source  = Parent.frm1.vspdData " & vbCr
	Response.Write " Parent.frm1.vspdData.Redraw = False " & vbCr      	
	Response.Write " Parent.ggoSpread.SSShowData  """ & lgstrData & """ ,""F""" & vbCr
	Response.Write " Parent.lgStrColorFlag = """ & lgStrColorFlag & """" & VbCr
	Response.Write " Parent.DbQueryOk " & vbCr		
 	Response.Write " Parent.frm1.vspdData.Redraw = True " & vbCr      
	Call EndScriptTag()
End Sub

%>



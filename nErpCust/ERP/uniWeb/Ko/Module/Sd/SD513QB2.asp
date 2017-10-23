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

    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3	'�� : DBAgent Parameter ���� 
    Dim lgstrData																'�� : data for spreadsheet data
    Dim lgTailList																'�� : Orderby���� ���� field ����Ʈ 
    Dim lgSelectList
    Dim lgSelectListDT
    Dim lgStrColorFlag
    
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
   
    Dim lgFromDt			'��ȸ�Ⱓ(��,��)���� 
    Dim lgToDt				'��ȸ�Ⱓ(��,��)�� 
    Dim lgBizAreaCd			'����� 
    Dim lgSalesOrgCd		'�������� 
    Dim lgSalesGrpCd		'�����׷� 
    Dim lgExceptFlag		'���ܿ��� 
    Dim lgPostFlag			'Ȯ������ 
        
    lgFromDt		= Trim(Request("txtHConFromDt"))
    lgToDt			= Trim(Request("txtHConToDt"))
    lgBizAreaCd		= Trim(Request("txtHConBizAreaCd"))
    lgSalesOrgCd	= Trim(Request("txtHConSalesOrgCd"))
    lgSalesGrpCd	= Trim(Request("txtHConSalesGrpCd"))
    lgExceptFlag	= Trim(Request("rdoHConExceptFlag"))
    lgPostFlag		= Trim(Request("rdoHConPostFlag"))
            
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd

    lgSelectList   = Request("txtHlgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("txtHlgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("txtHlgTailList")                                 '�� : Orderby value

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

    lgstrData      = ""

    iLoopCount = 0
    lgStrColorFlag = ""    
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 'COLOR
 		If rs0(0) > 0 Then
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

	'--------------- ������ coding part(�������,Start)----------------------------------------------------
   
    Redim UNISqlId(3)                                       '��: SQL ID ������ ���� ����Ȯ��    

    Redim UNIValue(3,8)										'��: DB-Agent�� ���۵� parameter�� ���� ���� 

	'��ȸ�Ⱓ(��,��)����==================================================================================
	If Len(lgFromDt) Then
		UNIValue(0,1) = " " & FilterVar(UNIConvDate(lgFromDt), "''", "S") & ""                           
	Else
		UNIValue(0,1) = "Null"
	End If		
	
	'��ȸ�Ⱓ(��,��)��====================================================================================
	If Len(lgToDt) Then
		UNIValue(0,2) = " " & FilterVar(UNIConvDate(lgToDt), "''", "S") & ""                           	
	Else
		UNIValue(0,2) = "Null"
	End If               	

	'����������===========================================================================================
    If Len(lgSalesOrgCd) Then		
		UNISqlId(2)		= "s0000qa006"	
		UNIValue(2,0)	= FilterVar(lgSalesOrgCd, "''", "S")
		UNIValue(0,3)	= " " & FilterVar(lgSalesOrgCd, "''", "S") & ""
	Else
		UNIValue(0,3)	= "NULL"
	End If	

	'�����׷��===========================================================================================	
    If Len(lgSalesGrpCd) Then
		UNISqlId(3)		= "s0000qa005"	
		UNIValue(3,0)	= FilterVar(lgSalesGrpCd, "''", "S")
		UNIValue(0,4)	= " " & FilterVar(lgSalesGrpCd, "''", "S") & ""
	Else
		UNIValue(0,4)	= "NULL"
	End If
	
	'������=============================================================================================    	
	If Len(lgBizAreaCd) Then
		UNISqlId(1)		= "s0000qa013"	
		UNIValue(1,0)	= FilterVar(lgBizAreaCd, "''", "S")
		UNIValue(0,5)	=  " " & FilterVar(lgBizAreaCd, "''", "S") & ""
	Else
		UNIValue(0,5)	= "Null"
	End If
	
	'���ܿ���=============================================================================================	
    If Len(lgExceptFlag) Then
		UNIValue(0,6)	= " " & FilterVar(lgExceptFlag, "''", "S") & ""
	Else
		UNIValue(0,6)	= "NULL"
	End If
	
	'Ȯ������=============================================================================================	
    If Len(lgPostFlag) Then
		UNIValue(0,7)	= " " & FilterVar(lgPostFlag, "''", "S") & ""
	Else
		UNIValue(0,7)	= "NULL"
	End If

	UNISqlId(0) = "SD513QA201"					
    UNIValue(0,0) = lgSelectList        
   
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    Set lgADF = Nothing													'��: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Exit Sub
    End If    
   
	Call BeginScriptTag()												'��:Write the Script Tag "<Script language=vbscript>"

	'����� ���翩�� 
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

	'�������� ���翩�� 
	If lgSalesOrgCd <> "" Then
		If rs2.EOF And rs2.BOF Then
			rs2.Close
			Set rs2 = Nothing			
			Call ConNotFound("txtConSalesOrgCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSalesOrgNm", rs2(1))		
		End If
	Else
		Call WriteConDesc("txtConSalesOrgNm", "")
	End If
	
	'�����׷� ���翩�� 
	If lgSalesGrpCd <> "" Then
		If rs3.EOF And rs3.BOF Then
			rs3.Close
			Set rs3 = Nothing			
			Call ConNotFound("txtConSalesGrpCd")			
			Exit Sub
		Else	
			Call WriteConDesc("txtConSalesGrpNm", rs3(1))		
		End If
	Else
		Call WriteConDesc("txtConSalesGrpNm", "")
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



<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

On Error Resume Next

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� ��������
Dim lgstrRetMsg                                                            '�� : Record Set Return Message ��������
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1                         '�� : DBAgent Parameter ����
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� ��
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ
Dim lgSelectList
Dim lgSelectListDT
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strFromDt																'�� : ������
Dim strToDt																	'�� : ������
Dim strBizCd																'�� : �����
Dim strFromAmt																'�� : ���۱ݾ�
Dim strToAmt																'�� : ���ݾ�
Dim striOpt																	'�� : ��� Grid����..
Dim strRdo																	'�� : '1' �̰���ȸ, '2' �ϰ���ȸ
Dim strTemphq

Dim strCond

Dim strMsgCd
Dim strMsg1

'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "A", "NOCOOKIE", "MB")	

    lgStrPrevKey   = Request("lgStrPrevKey")								'�� : Next key flag
    lgSelectList   = Request("lgSelectList")								'�� : select �����
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)				'�� : �� �ʵ��� ����Ÿ Ÿ��
    lgTailList     = Request("lgTailList")									'�� : Orderby value
	striOpt		   = Request("txtIOpt")										'��� Grid����..
	strRdo		   = Request("txtRdoFg")									''1'�̰���ȸ, '2' �ϰ���ȸ                             	

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
 
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub  MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr

	Const C_SHEETMAXROWS_D = 100 

    iCnt = 0
    lgstrData = ""

    If Len(Trim(lgStrPrevKey)) Then                                        '�� : Chnage Nextkey str into int value
		If Isnumeric(lgStrPrevKey) Then
			iCnt = CInt(lgStrPrevKey)
		End If   
    End If   

    For iRCnt = 1 to iCnt  *  C_SHEETMAXROWS_D                                   '�� : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
    Do While Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
        
        For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iRCnt < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If iRCnt < C_SHEETMAXROWS_D Then                                            '��: Check if next data exists
        lgStrPrevKey = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
	rs0.Close
	Set rs0 = Nothing 
	Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub  FixUNISQLData()
    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ��
    Redim UNIValue(1,2)

    UNISqlId(0) = "A8103MA101"
	UNISqlId(1) = "ABIZNM"

    UNIValue(0,0) = " DISTINCT " & lgSelectList                                          '��: Select list
    UNIValue(0,1) = strCond
    
    UNIValue(1,0) = UCase(FilterVar(strBizCd, "''", "S") )
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))

    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub  QueryData()
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
	If iStr(0) <> "0" Then
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If    
	
	If strRdo = "1" Then	    
		If striOpt = "B" Then       
		    If (rs1.EOF And rs1.BOF) Then
				If strMsgCd = "" And strBizCd <> "" Then
					strMsgCd = "970000"		'Not Found
					strMsg1 = Request("txtBIzArea_Alt")
				End If
		    Else
		%>
				<Script Language=vbScript>
				With parent
					.frm1.txtBIzArea.value = "<%=Trim(rs1(0))%>"
					.frm1.txtBIzAreaNm.value = "<%=Trim(rs1(1))%>"
				End With
				</Script>
		<%
		    End If
		    
			rs1.Close
			Set rs1 = Nothing 	
		End If   
	Else
		If striOpt = "A" Then       
		    If (rs1.EOF And rs1.BOF) Then
				If strMsgCd = "" And strBizCd <> "" Then
					strMsgCd = "970000"		'Not Found
					strMsg1 = Request("txtBIzArea_Alt")
				End If
		    Else
		%>
				<Script Language=vbScript>
				With parent
					.frm1.txtBIzArea.value = "<%=Trim(rs1(0))%>"
					.frm1.txtBIzAreaNm.value = "<%=Trim(rs1(1))%>"
				End With
				</Script>
		<%
		    End If

			rs1.Close
			Set rs1 = Nothing
		End If
	End If	

	If striOpt = "A" Then	
		If  rs0.EOF And rs0.BOF Then
		    Call DisplayMsgBox("990007", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
			Set rs0 = Nothing
		    Set lgADF = Nothing		
		Else
			Call  MakeSpreadSheetData()		
		End If
	Else	
		If  rs0.EOF And rs0.BOF Then
			
		Else
			Call  MakeSpreadSheetData()			
		End If			
	End If			
	
'    If  rs0.EOF And rs0.BOF Then
'		If strRdo = "1" Then	    
'			If striOpt = "B" Then       
'			    Call DisplayMsgBox("990007", vbOKOnly, "", "", I_MKSCRIPT)
'				rs0.Close
'			    Set rs0 = Nothing
'			    Set lgADF = Nothing	
'			End If	
'		Else
'			If striOpt = "A" Then       
'			    Call DisplayMsgBox("990007", vbOKOnly, "", "", I_MKSCRIPT)
'			    rs0.Close
'				Set rs0 = Nothing
'			    Set lgADF = Nothing	
'			End If	
'		End If	
'	Else    
'       Call  MakeSpreadSheetData()
'   End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub  TrimData()
    strFromDt  = UNIConvDate(Request("txtFromDt"))			'��������
    strToDt	   = UNIConvDate(Request("txtToDt"))			'��������
	strBizCd   = UCase(Trim(Request("txtBizArea")))         '�����
    strFromAmt = Request("txtFromAmt")						'���۱ݾ�
    strToAmt   = Request("txtToAmt")						'����ݾ�	 
    strTemphq  = Request("hTemphq")  
     
	If strFromDt <> "" Then
		strCond = strCond & " and A.TEMP_GL_DT >= " & FilterVar(strFromDt, "''", "S") 
    End If
     
    If strToDt <> "" Then
		strCond = strCond & " and A.TEMP_GL_DT <= " & FilterVar(strToDt, "''", "S")
    End If
     
    If strBizCd <> "" Then
		strCond = strCond & " and  D.BIZ_AREA_CD = " & FilterVar(strBizCd, "''", "S")
    End If
     
    If strFromAmt <> "" Then
		strCond = strCond & " and A.DR_LOC_AMT >= " & UNIConvNum(strFromAmt,0)
    End If

    If strToAmt <> "" Then
		strCond = strCond & " and A.DR_LOC_AMT <= " & UNIConvNum(strToAmt,0)
    End If
     	
	If strRdo = "1" Then							'�̰�
		If striOpt = "A" Then						'����
			strCond = strCond & " AND B.DR_CR_FG = " & FilterVar("DR", "''", "S") & "  "
		ElseIf striOpt = "B" Then					'�뺯
			strCond = strCond & " AND B.DR_CR_FG = " & FilterVar("CR", "''", "S") & "  "
		End If
		
	 	strCond = strCond & " AND A.GL_INPUT_TYPE = " & FilterVar("TG", "''", "S") & "  AND isnull(A.HQ_BRCH_NO,'')  = '' "
	 	
	ElseIf strRdo = "2" Then						'�ϰ�
		If striOpt = "A" Then						'����

			strCond = strCond & " AND B.DR_CR_FG = " & FilterVar("DR", "''", "S") & "  AND isnull(A.HQ_BRCH_NO,'')  <> '' "
		ElseIf striOpt = "B" Then					'�뺯
			strCond = strCond & " AND B.DR_CR_FG = " & FilterVar("CR", "''", "S") & "  AND  A.HQ_BRCH_NO = " & FilterVar(strTemphq, "''", "S")
		End If	
	End If
End Sub
%>

<Script Language=vbscript>
    With parent
<% 
		If striOpt = "A" Then 
%>		
			.ggoSpread.Source    = .frm1.vspdData
			.lgStrPrevKey_1      = "<%=lgStrPrevKey%>"                       '��: set next data tag
<%			
        ElseIf striOpt = "B" Then   
%>        
			.ggoSpread.Source    = .frm1.vspdData2
			.lgStrPrevKey_2      = "<%=lgStrPrevKey%>"                       '��: set next data tag
<%			
        End If
%>        
        If Trim(.frm1.txtBizArea.value) = "" Then	
			.frm1.txtBizAreaNm.Value = ""
		End If     
        .ggoSpread.SSShowData  "<%=lgstrData%>"                          '��: Display data 
        '.lgStrPrevKey_A      = "<%=lgStrPrevKey%>"                       '��: set next data tag       
	    .DbQueryOk("<%=striOpt%>")
	End with
</Script>

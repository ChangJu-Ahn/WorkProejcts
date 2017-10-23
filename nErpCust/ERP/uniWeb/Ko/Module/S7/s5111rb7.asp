<%'======================================================
'*  1. Module Name          : ���� 
'*  2. Function Name        : ������� 
'*  3. Program ID           : s5111rb7
'*  4. Program Name         : ������ ��Ȳ Popup
'*  5. Program Desc         : ������ ��Ȳ�� Display�Ѵ�.
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/05/07
'*  8. Modified date(Last)  : 2002/05/07
'*  9. Modifier (First)     : Hwangseongbae
'* 10. Modifier (Last)      : Hwangseongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'=======================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%
On Error Resume Next                                                                         
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3			   '�� : DBAgent Parameter ���� 
Dim lgStrData                                                 '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
Dim lgMaxCount                                                '�� : Spread sheet �� visible row �� 
Dim lgTailList                                                '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

	Call LoadBasisGlobalInf()
    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = 30							             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
	lgTailList     = Request("lgTailList")                                 '�� : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgDataExist      = "No"
	    
    Call FixUNISQLData()									 '�� : DB-Agent�� ���� parameter ����Ÿ set
    Call QueryData()										 '�� : DB-Agent�� ���� ADO query
    
'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < lgMaxCount Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop

    If iLoopCount < lgMaxCount Then                                 '��: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Sub SetConditionData()
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(0,7)

    UNISqlId(0) = "S5111RA701"
'--------------- ������ coding part(�������,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '��: Select list
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    UNIValue(0,1) = " " & FilterVar(Request("txtSoldToParty"), "''", "S") & ""	'�ֹ�ó 
    UNIValue(0,2) = " " & FilterVar(Request("txtDocCur"), "''", "S") & ""                            'ȭ�� 
    
    If Len(Request("txtPrrcptType")) Then
		UNIValue(0,3) = " " & FilterVar(Request("txtPrrcptType"), "''", "S") & ""                      '���������� 
	Else
		UNIValue(0,3) = "Null"                              '���������� 
	End If
    
    If Len(Request("txtFrPrRctpDt")) Then
		UNIValue(0,4) = " " & FilterVar(UNIConvDate(Request("txtFrPrRctpDt")), "''", "S") & ""                             '���������� 
	Else
		UNIValue(0,4) = "Null"                              '�Ⱓ 
	End If

    If Len(Request("txtToPrRctpDt")) Then
		UNIValue(0,5) = " " & FilterVar(UNIConvDate(Request("txtToPrRctpDt")), "''", "S") & ""                             '���������� 
	Else
		UNIValue(0,5) = "Null"                              '�Ⱓ 
	End If

    If Len(Request("txtPrrcptNo")) Then
		UNIValue(0,6) = " " & FilterVar(Trim(Request("txtPrrcptNo")) & "%", "''", "S") & ""                 '�����ݹ�ȣ 
	Else
		UNIValue(0,6) = "" & FilterVar("%", "''", "S") & ""									'�����ݹ�ȣ 
	End If

'--------------- ������ coding part(�������,End)------------------------------------------------------

    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
    Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        rs0.Close
        Set rs0 = Nothing
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        %>
		<Script Language=vbscript>
		Call parent.DbQueryOk
		</Script>	
        <%
    Else    
        Call  MakeSpreadSheetData()
        Call  SetConditionData()
        Call WriteResult()
    End If  
End Sub

' ��ȸ ����� Display�ϴ� Script �ۼ� 
Sub WriteResult()
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent.frm1" & vbCr
	Response.Write "Parent.ggoSpread.Source	= .vspdData" & vbCr
 	Response.Write ".vspdData.Redraw = False " & vbCr      
	Response.Write "parent.ggoSpread.SSShowDataByClip """ & lgstrData  & """ ,""F""" & vbCr
	Response.Write "parent.lgPageNo	= """ & lgPageNo & """" & vbCr
	Response.Write "parent.DbQueryOk" & vbCr
 	Response.Write ".vspdData.Redraw = True " & vbCr      
	Response.Write "End with" & vbCr
	Response.Write "</Script>" & vbCr
End Sub
%>

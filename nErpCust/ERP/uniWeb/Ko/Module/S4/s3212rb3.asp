<%
'********************************************************************************************************
'*  1. Module Name          : ��������																	*
'*  2. Function Name        : 																			*
'*  3. Program ID           : s3212rb3.asp																*
'*  4. Program Name         : Local L/C ��������(Local L/C��Ȳ��ȸ����)									*
'*  5. Program Desc         : Local L/C ��������(Local L/C��Ȳ��ȸ����)									*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2002/04/11																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Kwak Eun Kyoung															*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : ȭ�� design												*
'*                            2. 2001/12/18 : Son bum Yeol												*
'*                            3. 2002/04/11 : ADO��ȯ													*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
																				
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "RB")   
Call LoadBNumericFormatB("Q","S","NOCOOKIE","RB")

On Error Resume Next

   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0	           '�� : DBAgent Parameter ���� 
   Dim lgStrData                                               '�� : Spread sheet�� ������ ����Ÿ�� ���� ����   
   Dim lgTailList
   Dim lgSelectList
   Dim lgSelectListDT
   Dim lgDataExist
   Dim lgPageNo

'--------------- ������ coding part(��������,Start)----------------------------------------------------
   Dim iFrPoint
   iFrPoint=0
   Const C_SHEETMAXROWS_D  = 30                                          '��: Fetch max count at once
'--------------- ������ coding part(��������,End)------------------------------------------------------

    Call HideStatusWnd 
   
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)                  
    lgSelectList     = Request("lgSelectList")
    lgTailList     = " ORDER BY A.LC_SEQ ASC "
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '�� : �� �ʵ��� ����Ÿ Ÿ�� 

    lgDataExist      = "No"

    Call  FixUNISQLData()                                                '�� : DB-Agent�� ���� parameter ����Ÿ set
    call  QueryData()                                                    '�� : DB-Agent�� ���� ADO query

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
       rs0.Move = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       iFrPoint	= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
    End If
    
    iLoopCount = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop

    If iLoopCount < C_SHEETMAXROWS_D Then                                      '��: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '��: Close recordset object
    Set rs0 = Nothing	                                            '��: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Dim arrVal(0)
    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    Redim UNIValue(0,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
    UNISqlId(0) = "s3212ra301" 
'--------------- ������ coding part(�������,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)		                              '��: Select ������ Summary    �ʵ� 

'--------------- ������ coding part(�������,Start)----------------------------------------------------
	strVal = " "

	If Len(Request("txtLCNo")) Then
		strVal = FilterVar(Trim(Request("txtLCNo")),"","S")
	End If

'--------------- ������ coding part(�������,End)----------------------------------------------------
    UNIValue(0,1) = strVal  

 
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

    Set lgADF   = Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
 
    If rs0.EOF And rs0.BOF Then
       Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
       rs0.Close
       Set rs0 = Nothing
       Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If
    
End Sub

%>
<Script Language=vbscript>

    If "<%=lgDataExist%>" = "Yes" Then
       'Show multi spreadsheet data from this line
		With parent       
			.ggoSpread.Source  = .frm1.vspdData
			.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowDataByClip  "<%=lgstrData%>", "F"
		
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,Parent.frm1.txtCurrency.value,Parent.GetKeyPos("A",7),"C", "Q" ,"X","X")		
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.frm1.vspdData,"<%=iFrPoint+1%>",parent.frm1.vspddata.maxrows,Parent.frm1.txtCurrency.value,Parent.GetKeyPos("A",8),"A", "Q" ,"X","X")
			
			.lgPageNo	  	   =  "<%=lgPageNo%>"  				  '��: Next next data tag
			.DbQueryOk
			.frm1.vspdData.Redraw = True		
			
		End with
    End If   
</Script>	

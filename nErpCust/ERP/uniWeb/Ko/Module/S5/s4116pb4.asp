<%
'********************************************************************************************************
'*  1. Module Name          : ���� 
'*  2. Function Name        : 
'*  3. Program ID           : S4116PA4
'*  4. Program Name         : ������Ȳ 
'*  5. Program Desc         : ������Ȳ 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/29
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Cho in kuk
'* 10. Modifier (Last)      : 
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
Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "RB")
On Error Resume Next

Call HideStatusWnd

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data

Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList													       '�� : select ����� 
Dim lgSelectListDT														   '�� : �� �ʵ��� ����Ÿ Ÿ��	
Dim lgPageNo

Const C_SHEETMAXROWS_D  = 30   

Dim iStrDNNo
Dim iStrDNType
Dim iStrShipToParty
Dim iStrFromDt
Dim iStrToDt
Dim iStrConfFlag

iStrDNNo = Request("txtConDNNo")
iStrDNType = Request("txtConDnType")
iStrShipToParty = Request("txtConShipToParty")
iStrFromDt = Request("txtConFromDt")
iStrToDt = Request("txtConToDt")
iStrConfFlag = Request("txtConConfFlag")

lgPageNo       = UNICInt(Trim(Request("txtHlgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)    
lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
lgTailList     = Request("lgTailList")                                 '�� : Order by value

Call TrimData()
Call FixUNISQLData()
Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    lgstrData      = ""
    
    If CInt(lgPageNo) > 0 Then
       rs0.Move  =  C_SHEETMAXROWS_D * CInt(lgPageNo)
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
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim iStrVal					
								
    Redim UNISqlId(0)           
								
    Redim UNIValue(0,2)			

    UNISqlId(0) = "S4116PA401" 
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

	iStrVal = " "
   
	
	'��ȸ�Ⱓ����=========================================================================================
	If Len(iStrFromDt) Then
		If iStrConfFlag = "Y" Then
			iStrVal = iStrVal & " DH.ACTUAL_GI_DT >=  " & FilterVar(UNIConvDate(iStrFromDt), "''", "S") & ""		
		Else
			iStrVal = iStrVal & " DH.PROMISE_DT >=  " & FilterVar(UNIConvDate(iStrFromDt), "''", "S") & ""		
		End If
	End If		
	
	'��ȸ�Ⱓ��===========================================================================================
	If Len(iStrToDt) Then
		If iStrConfFlag = "Y" Then
			iStrVal = iStrVal & " AND DH.ACTUAL_GI_DT <=  " & FilterVar(UNIConvDate(iStrToDt), "''", "S") & ""		
		Else 
			iStrVal = iStrVal & " AND DH.PROMISE_DT <=  " & FilterVar(UNIConvDate(iStrToDt), "''", "S") & ""		
		End If
	End If
		
    '���Ϲ�ȣ=============================================================================================    	
	If Len(iStrDNNo) Then
		iStrVal = iStrVal & " AND DH.DN_NO =  " & FilterVar(iStrDNNo, "''", "S") & ""				
	End If
	
	'��������=============================================================================================    	
	If Len(iStrDNType) Then
		iStrVal = iStrVal & " AND DH.MOV_TYPE =  " & FilterVar(iStrDNType, "''", "S") & ""				
	End If
	
	'��ǰó=============================================================================================    	
	If Len(iStrShipToParty) Then
		iStrVal = iStrVal & " AND DH.SHIP_TO_PARTY =  " & FilterVar(iStrShipToParty, "''", "S") & ""				
	End If	 
	
	'Ȯ������===========================================================================================
	If Len(iStrConfFlag) Then
		iStrVal = iStrVal & " AND DH.POST_FLAG =  " & FilterVar(iStrConfFlag , "''", "S") & ""		
	End If
		
    UNIValue(0,1) = iStrVal   
        
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)�� ������ ?�� �Էµ�	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	Dim FalsechkFlg
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")									'��:ADO ��ü�� ���� 
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

    FalsechkFlg = False
    
    iStr = Split(lgstrRetMsg,gColSep)
	
	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If  rs0.EOF And rs0.BOF And FalsechkFlg = False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End     
    Else   
        Call  MakeSpreadSheetData()
    End If
   
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
End Sub
%>
<Script Language=vbscript>
    With parent  
        .ggoSpread.Source = .frm1.vspdData
        .frm1.vspdData.Redraw = False
        .ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"
        .lgPageNo = "<%=lgPageNo%>"		
		.frm1.vspdData.Redraw = True
		Call .DbQueryOk
   	End with
</Script>	 	
<%
	Response.End													'��: �����Ͻ� ���� ó���� ������ 
%>

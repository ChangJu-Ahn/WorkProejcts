<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : A5124mb1
'*  4. Program Name         : ������ ��������ȸ 
'*  5. Program Desc         : 
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.11.15
'*  8. Modified date(Last)  : 2001.11.15
'*  9. Modifier (First)     : Oh, Soo Min
'* 10. Modifier (Last)      : Oh, Soo Min
'* 11. Comment              :
'*   - 2001.03.21  Song,Mun Gil  �����ڵ�, ���¹�ȣ ���� Check ���� �߰� 
'=======================================================================================================
Response.Expires = -1                                                       '�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True                                                     '�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

%>

<!-- #Include file="../../inc/IncServer.asp"  -->

<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
	                                                                   '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
'On Error Resume Next

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ����Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgstrData 
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strFromGlDt, strToGLDt, strDeptCd, strAcctCd, strInternalCd
Dim TDrLocAmt,TCrLocAmt,NDrLocAmt,NCrLocAmt, Bal_Fg
Dim TDrSumAmt,NDrSumAmt,SDrSumAmt,TCrSumAmt,NCrSumAmt,SCrSumAmt,SDrAmt, SCrAmt
Dim strMsgCd, strMsg1, strMsg2 
Dim strWhere0, strWhere1
Dim strGetInternalNm
Dim strCompYr,strCompMnth,strCompDt, strGlDtYr, strGlDtMnth, strGlDtDt
Dim strCompFiscStartDt
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 

    lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr
    Dim  ctrlval1
    Dim  ctrlcd1
    Dim  ctrlval2
    Dim  ctrlcd2
    iCnt = 0

    If Len(Trim(lgStrPrevKey)) Then                                        '�� : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  lgMaxCount                                   '�� : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            Select Case  lgSelectListDT(ColCnt)
               Case "DD"   '��¥ 
                           iStr = iStr & Chr(11) & UNIDateClientFormat(rs0(ColCnt))
               Case "F2"  ' �ݾ� 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggAmtOfMoney.DecPoint, 0)
               Case "F3"  '����7
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggQty.DecPoint       , 0)
               Case "F4"  '�ܰ� 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggUnitCost.DecPoint  , 0)
               Case "F5"   'ȯ�� 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggExchRate.DecPoint  , 0)
               Case Else
                    iStr = iStr & Chr(11) & ConvSPChars(Trim("" & rs0(ColCnt)))
            End Select
        
           
            Select Case  ColCnt
               Case 1
                 ctrlval1 = ConvSPChars(Trim("" & rs0(ColCnt)))                                     
               Case 3
                 ctrlval2 = ConvSPChars(Trim("" & rs0(ColCnt)))                                        
               Case 7
                 ctrlcd1  = ConvSPChars(Trim("" & rs0(ColCnt)))                       
               Case 8                   
                 ctrlcd2  = ConvSPChars(Trim("" & rs0(ColCnt)))                                                       
               Case Else
                    
            End Select      
            IF ctrlval1 <>  "" THEN
            
            END IF
		Next
		
		
 
        If  iRCnt < lgMaxCount Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < lgMaxCount Then                                            '��: Check if next data exists
        lgStrPrevKey = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
                                               '��: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNISqlId(0) = "A5132MA101"	'������������ȸ 
    UNISqlId(1) = "A5132MA102"	'������������ȸ 
	
	Redim UNIValue(1,2)
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    UNIValue(0,1) = UCase(Trim(strWhere0))
	UNIValue(0,2) = UCase(Trim(lgTailList))
	UNIValue(1,0) = UCase(Trim(strWhere0))
	
	'--------------- ������ coding part(�������,End)------------------------------------------------------
	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    'lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1)
    
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    If Trim(strGetInternalNm) = "" Then
       If strMsgCd = "" And strDeptCd <> "" Then 
		  strMsgCd = "970000"												'Not Found
          strMsg1 = Request("txtDeptCd_Alt")
       End If
    %>	

   <%	

    Else
    %>	
    <Script Language=vbScript>
	  With parent
		.frm1.txtDeptCd.value = "<%=ConvSPChars(strDeptCd)%>"
		.frm1.txtDeptNm.value = "<%=ConvSPChars(Trim(strGetInternalNm))%>"
				
	  End With 
    </Script>
   
   <%	
    End If

	
%>
	
<%	
		
	If rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
'		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
'		rs0.Close
'		Set rs0 = Nothing
'		Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
        Call  MakeSpreadSheetData()
    End If
    
    TDrSumAmt = 0 
    TCrSumAmt = 0
    
    
    
	If Not(rs1.EOF And rs1.BOF) Then
		If IsNull(rs1(0)) = False Then TDrSumAmt    = rs1(0)
		If IsNull(rs1(1)) = False Then TCrSumAmt    = rs1(1)
	End If
	
	
	rs1.Close
	Set rs1 = Nothing
	
        	
 %>
    
    <Script Language=vbscript>
    With parent
       	.frm1.txtSDrAmt.value		= "<%=UNINumClientFormat(TDrSumAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtSCrAmt.value		= "<%=UNINumClientFormat(TCrSumAmt, ggAmtOfMoney.DecPoint, 0)%>"			
	End With
	</script>
	<%
	
	rs0.Close
	Set rs0 = Nothing 
	
	Set lgADF = Nothing  	
	                                                  '��: ActiveX Data Factory Object Nothing
	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		'Response.End 
	
	End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()  
	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	strFromGlDt = uniconvdate(Request("txtFromGlDt"))
	strToGLDt = uniconvdate(Request("txtToGlDt"))
	strAcctCd = UCase(FilterVar(Request("txtAcctCd"),"","SNM"))
	strWhere0 = ""
	strWhere0 = strWhere0 & " A.Acct_cd = '" & strAcctCd & "' "
	strWhere0 = strWhere0 & " and A.FISC_YR +A.FISC_MNTH + A.FISC_DT  between '" & strFromGlDt & "' and '" & strToGLDt & "' "
	
End Sub




%>

<Script Language=vbscript>
	With parent
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData "<%=lgstrData%>"                            '��: Display data 
		.lgStrPrevKey =  "<%=ConvSPChars(lgStrPrevKey)%>"                       '��: set next data tag
		.DbQueryOk
	End with
	
</Script>	

<%
	Response.End 
%>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  


<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
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
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strfrgldt	                                                           
Dim strtogldt
Dim strfrglno	                                                           
Dim strtoglno
Dim strdeptcd
Dim strrefno
Dim strUsrId
Dim strdesc
Dim strInputType
Dim strDrLocAmtFr
Dim strDrLocAmtTo
	                                                           '�� : ������ 
Dim strCond
Dim strDeptNm
Dim strInputTypeNm
Dim strBizArea

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
Call HideStatusWnd 
Call LoadBasisGlobalInf()
Call loadInfTB19029B("Q", "A","NOCOOKIE","QB")
'Call LoadBNumericFormatB("Q", "A","NOCOOKIE","QB")

lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
lgTailList     = Request("lgTailList")                                 '�� : Orderby value

	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))
	
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
Const C_SHEETMAXROWS_D  = 30                                   '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 

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
    
    strDeptNm = rs0(1)
    strInputTypeNm = rs0(7)
    
    Do while Not (rs0.EOF Or rs0.BOF)
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

    If  iRCnt < C_SHEETMAXROWS_D Then                                            '��: Check if next data exists
        lgStrPrevKey = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(0,2)

    UNISqlId(0) = "A5104RA101"

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNIValue(0,1) = strCond
    'UNIValue(0,2) = UCase(Trim(strToglDt))
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
     strfrgldt     = UNIConvDate(Request("txtfrgldt"))
     strtogldt     = UNIConvDate(Request("txttogldt"))
     strfrglno	   = Request("txtfrglno")                                                          
     strtoglno     = Request("txttoglno")
     strdeptcd     = Request("txtdeptcd")
     strrefno      = Request("txtrefno")
     strdesc       = Request("txtdesc")
     strUsrId	  = Request("txtUsrId")
     strInputType  = Trim(Request("txtInputType"))'
     strDrLocAmtFr = UNIConvNum(Request("txtDrLocAmtFr"),0)'	("txtDrLocAmtFr")     
     strDrLocAmtTo = UNIConvNum(Request("txtDrLocAmtTo"),0)'Request("txtDrLocAmtTo")     
     
     
     

    
     If strfrgldt <> "" Then
		strCond = strCond & " and a.gl_dt >=  " & FilterVar(strfrgldt , "''", "S") & ""
     End If
     
     If strtogldt <> "" Then
		strCond = strCond & " and a.gl_dt <=  " & FilterVar(strtogldt , "''", "S") & "" 
     End If
     
     If strfrglno <> "" Then
		strCond = strCond & " and a.gl_no >= " & FilterVar(strfrglno , "''", "S") 
     End If
     
     If strtoglno <> "" Then
		strCond = strCond & " and a.gl_no <= " & FilterVar(strtoglno , "''", "S") 
     End If
     
     If strdeptcd <> "" Then
		strCond = strCond & " and a.dept_cd = " & FilterVar(strdeptcd , "''", "S") 
     End If
 
	If strBizArea <> "" Then
		strCond = strCond & " and a.biz_area_cd = " & FilterVar(strBizArea, "''", "S") 
	End If
    
    If strUsrId <> "" Then
		strCond = strCond & " and a.INSRT_USER_ID = " & FilterVar(strUsrId, "''", "S") 
    End If
    
'�Է°�� 
     If strInputType <> "" Then
		strCond = strCond & " and a.gl_input_type = " & FilterVar(strInputType, "''", "S")
     End If
'�ݾ�UNIConvNum(Request("txtPrPaymLocAmt"),0)	
     If strDrLocAmtFr <> 0 or strDrLocAmtTo <> 0 Then
		if strDrLocAmtFr > 0 and strDrLocAmtTo <= 0 Then
			strCond = strCond & " and a.dr_loc_amt >= " & strDrLocAmtFr 
		Elseif strDrLocAmtFr <= 0 and strDrLocAmtTo > 0 Then
			strCond = strCond & " and a.dr_loc_amt <= " & strDrLocAmtTo 
		Else
			strCond = strCond & " and a.dr_loc_amt between " & strDrLocAmtFr & " and " & strDrLocAmtTo
		end If
     End If

'-   

     If strrefno <> "" Then
			strrefno = strrefno & "%"
     		strCond = strCond & " and a.ref_no LIKE " & FilterVar(strrefno, "''", "S") 
     End If

     If strdesc <> "" Then
		strdesc = "%" & strdesc & "%" 
		strCond = strCond & " and a.gl_desc LIKE " & FilterVar(strdesc, "''", "S") 
     End If

	

     strCond = strCond & " and A.GL_INPUT_TYPE <> " & FilterVar("GD", "''", "S") & "  "
     
     iF Request("lgAuthorityFlag") = "Y" then      '���Ѱ��� �߰� 
		strCond = strCond & " and EXISTS ( SELECT 1 FROM z_usr_authority_value S WHERE a.dept_cd = S.code_value and S.usr_id =  " & FilterVar(gUsrId , "''", "S") & " AND S.module_cd = " & FilterVar("A", "''", "S") & "  )  "   '���Ѱ��� �߰� 
	 end if      '���Ѱ��� �߰� 


	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then
		strCond  = strCond & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  
	End If
	
	If lgInternalCd <> "" Then
		strCond  = strCond & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  
	End If
	
	If lgSubInternalCd <> "" Then
		strCond  = strCond & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If
	
	If lgAuthUsrID <> "" Then
		strCond  = strCond & " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If
	     
    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub






%>

<Script Language=vbscript>
    With parent
		 IF Trim(.frm1.txtDeptCd.value) <> "" Then
			.frm1.txtDeptNm.Value = "<%=ConvSPChars(strDeptNm)%>"
		 ElseIF Trim(.frm1.txtDeptcd.value) = "" Then	
			.frm1.txtDeptNm.Value = ""
		 END IF	         
		 IF Trim(.frm1.txtInputType.value) <> "" Then
			.frm1.txtInputTypeNM.Value = "<%=ConvSPChars(strInputTypeNm)%>"
		 ElseIF Trim(.frm1.txtInputType.value) = "" Then	
			.frm1.txtInputTypeNM.Value = ""
		 END IF	         
		 
         .ggoSpread.Source    = .frm1.vspdData 
         .ggoSpread.SSShowData "<%=ConvSPChars(lgstrData)%>"                            '��: Display data 
         .lgStrPrevKey        = "<%=ConvSPChars(lgStrPrevKey)%>"                       '��: set next data tag
         .DbQueryOk
	End with
</Script>

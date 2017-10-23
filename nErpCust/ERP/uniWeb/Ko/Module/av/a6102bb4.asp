<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>


<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<% 
	Call LoadBasisGlobalInf() 

    On Error Resume Next																'��: 
    Err.Clear																			'��: Clear Error status

	Dim lgOpModeCRUD


    Call HideStatusWnd																	'��: Hide Processing message
    lgOpModeCRUD      = Request("txtMode")												'��: Read Operation Mode (CRUD)
 
  
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)															'��: Query
            Call SubBizQuery()
            Case CStr(UID_M0002)                                                        '��: Save,Update
          '  Call SubBizSave()
          '  Call SubBizSaveMulti()
			Call SubFileDownLoad()          
                 
        Case CStr(UID_M0003)															'��: Delete
            Call SubFileDownLoad2()          
    End Select
    
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'-----------------------------------------------------------------------------------------------
'            File DownLoad(With B.A)
'-----------------------------------------------------------------------------------------------
Sub SubFileDownLoad()
    On Error Resume Next    
    Err.Clear																			'��: Clear Error status

    Dim strFilePath																		'��: Protect system from crashing
    
    Call HideStatusWnd

    strFilePath = "http://" & Request.ServerVariables("SERVER_NAME") & ":" _
        & Request.ServerVariables("SERVER_PORT")

    If Instr(1, Request.ServerVariables("URL"), "Module") <> 0 Then
        strFilePath = strFilePath & Mid(Request.ServerVariables("URL"), 1, InStr(1, Request.ServerVariables("URL"), "Module") - 1)     
    End If

    strFilePath = strFilePath  & "files/u2000/"
    strFilePath = strFilePath & Request("txtFileName")
End Sub

'--------------------------------------------------------------------------------------------------------
'                 File DownLoad(With xinSoft solution)
'--------------------------------------------------------------------------------------------------------
Sub SubFileDownLoad2()
    On Error Resume Next    
    Err.Clear                                                               '��: Protect system from crashing
    Dim xdn

    Call HideStatusWnd

    set xdn = Server.CreateObject("Xionsoft.XionFileDownLoad")

    strFilePath = "/" & Mid(Request.ServerVariables("URL"), 2, instr(2, Request.ServerVariables("URL"), "/") - 2) & "/template/files/" & gCompany & "/"

    'Call ServerMesgBox(strFilePath, vbInformation, I_MKSCRIPT)

    '���� �� �� �ӽ� COMMENT 2001/1/27
    xdn.DownFromFile strFilePath & Request("txtFileName")
    xdn.OnEndPage

    set xdn = nothing

    Response.End
End Sub

'--------------------------------------------------------------------------------------------------------
'              
'--------------------------------------------------------------------------------------------------------

Sub SubBizQuery()
 
    Dim ex1_file_name
    Dim ex2_return_code
    
    Const I1_ief_supplied           = 0
    Const I2_biz_area_cd            = 1
    Const I3_start_issued_dt        = 2
    Const I4_end_issued_dt          = 3
    Const I5_report_issued_dt       = 4
    Const I6_file_path_lef_supplied = 5
    Const I6_B_st_dt                = 6
    Const I6_B_end_dt               = 7     
    Const I7_daeRi                  = 8
    Const I8_gigubun                = 9
    Const I9_singoGubun             = 10
    Const I10_year                  = 11
    Const I10_chkYN                 = 12

    Dim strGubun
    Dim arrValue
    ReDim arrValue(I10_chkYN) 
    Dim strFilePath
    Dim iPAVG013

    On Error Resume Next
	Err.Clear  
    
    arrValue(I1_ief_supplied)           = UCase(Trim(Request("cbofileGubun")))
    arrValue(I2_biz_area_cd)            = UCase(Trim(Request("txtBizAreaCD")))
    arrValue(I3_start_issued_dt)        = replace(UNIConvDate(Request("txtIssueDt1")),"-","")
    arrValue(I4_end_issued_dt)          = replace(UNIConvDate(Request("txtIssueDt2")),"-","")
    arrValue(I5_report_issued_dt)       = replace(UNIConvDate(Request("txtReportDt")),"-","")
    strFilePath                         = Server.MapPath("../../files/u2000") & "\"
    arrValue(I6_file_path_lef_supplied) = strFilePath
	arrValue(I6_B_st_dt)	            = UNIConvDate(Request("txtIssueDt5"))
	arrValue(I6_B_end_dt)	            = UNIConvDate(Request("txtIssueDt6"))
    arrValue(I7_daeRi)                  = Trim(Request("chkDaeri")) 
    arrValue(I8_gigubun)                = Trim(Request("cboGiGubun")) 
    arrValue(I9_singoGubun)             = Trim(Request("cboSingoGubun"))
    arrValue(I10_year)                  = Trim(Request("txtYear"))
    arrValue(I10_chkYN)                  = Trim(Request("chkYN"))

    Set iPAVG013 = Server.CreateObject("PAVG013.cACardDiskSvrEab")

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = True Then
        Exit Sub
    End If	

    If Trim(gStrGlobalCollection) = "" Then
        Call DisplayMsgBox("127310", vbInformation, "", "", I_MKSCRIPT)     'Can not create(Demo code)
        Exit Sub   
    End If

    '//Response.Write "gStrGlobalCollection = " & gStrGlobalCollection
    '//gStrGlobalCollection = "2::0::::::::::::1900-01-01::YYYY-MM-DD::-::Provider=SQLOLEDB.1;Persist Security Info=False;User ID              = sa;password             = dba0203;Initial Catalog      = uni7test;Data Source          = 70.7.103.151::unierp::KO::U2000::KFC2::70.7.103.151::uni7test::70.7.31.157::KRW"

    call iPAVG013.EAB_A_CARD_DISK_SVR(gStrGlobalCollection,arrValue,ex1_file_name,ex2_return_code) 

    If CheckSYSTEMError(Err, True) = True Then     
        Set iPAVG013 = Nothing
        Exit Sub
    End If    

    Select Case ex2_return_code
   Case "A" ' �Ű����������� ã�� ��  �����ϴ�.
%>
 <Script Language=vbscript>
    Call DisplayMsgBox("700106","X","X","X")
 </Script>
<%
    'Call ServerMesgBox("�Ű����������� ã�� �� �����ϴ�." , vbInformation, I_MKSCRIPT)
   Case "B" ' �����ڷ������� ã�� ��  �����ϴ�.
%>
 <Script Language=vbscript>
    Call DisplayMsgBox("700107","X","X","X")
 </Script>
<%
    'Call ServerMesgBox("�����ڷ������� ã�� ��  �����ϴ�." , vbInformation, I_MKSCRIPT)
   Case "C" ' �����ڷ������� ã�� ��  �����ϴ�.
%>
 <Script Language=vbscript>
    Call DisplayMsgBox("700108","X","X","X")
 </Script>
<%
    'Call ServerMesgBox("�����ڷ������� ã�� ��  �����ϴ�." , vbInformation, I_MKSCRIPT)
   Case "D" ' �ΰ��������� ó���� �Ϸ���� �ʾҽ��ϴ�. �ٽ� ���� �Ͻʽÿ�.
%>
 <Script Language=vbscript>
    Call DisplayMsgBox("700109","X","X","X")
 </Script>
<%
    'Call ServerMesgBox("�ΰ��������� ó���� �Ϸ���� �ʾҽ��ϴ�. �ٽ� ���� �Ͻʽÿ�." , vbInformation, I_MKSCRIPT)
   Case "Z" ' ���ȭ�� �ٿ�ε� �۾�
    
   
    Call HideStatusWnd
    Set Ag0102 = Nothing

%>
    <SCRIPT LANGUAGE=VBSCRIPT>
        On Error Resume Next
        parent.frm1.txtFileName.value = "<%=ex1_file_name%>"
        parent.subVatDiskOK("<%=ex1_file_name%>")

        'Dim SF

        'Set SF = CreateObject("uni2kCM.SaveFile")
        'Call SF.SaveTextFile("<%= strFilePath %>")

        'Set SF = Nothing
		 'parent.subVatDiskOK2()
    </SCRIPT>
<%
  End Select

  Call HideStatusWnd
  Set iPAVG013 = Nothing               '��: Unload Comproxy

  Response.End   
       
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
%>




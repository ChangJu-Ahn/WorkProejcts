
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : �λ�/�޿� 
'*  2. Function Name        : ������ ��ȸ 
'*  3. Program ID           : b2903mb1
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'                             
'*  7. Modified date(First) : 2001//
'*  8. Modified date(Last)  : 2002/12/17
'*  9. Modifier (First)     : �̼��� 
'* 10. Modifier (Last)      : Sim Hae Young
'* 11. Comment              : Ʈ������ �̺�Ʈ�� ó���Ѵ� 
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -2000/03/22 : ..........
'**********************************************************************************************

Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.


'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
-->
<HTML>
<HEAD> <TITLE>�������</TITLE>

<!--
'=======================================================================================================
'												1. �� �� �� 
'=======================================================================================================
'=======================================================================================================
'                                               1.1 Inc ����   
'	���: Inc. Include
'=======================================================================================================-->
<!-- #Include file="../inc/IncServer.asp" -->
<!-- #Include file="../inc/adovbs.inc" -->
<!-- #Include file="../inc/incServeradodb.asp" -->

<!--
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================-->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">
<!--
'============================================  1.1.2 ���� Include  ======================================
'========================================================================================================-->

<SCRIPT LANGUAGE="VBScript"   SRC="../inc/IncCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/IncCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/IncCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/IncCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/incCliRdsQuery.vbs"></SCRIPT>
<Script Language="JavaScript" SRC="../inc/incImage.js"> </SCRIPT>

<SCRIPT LANGUAGE="VBScript">
Option Explicit                                                        '��: indicates that All variables must be declared in advance
Dim arrParent

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)


<%

													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

    Dim ADOConn
    Dim ADORs
    Dim StrSql
    Dim EmpDetail
      
    Dim NAME
    Dim EMP_NO
    Dim PAY_GRADE1
    Dim DEPT_NM
    Dim TEL_NO
    Dim EM_TEL_NO
    Dim HAND_TEL_NO
    Dim ADDR
    Dim EMAIL_ADDR
    Dim FUNC
    Dim ROLE
    Dim ImageSrc
    
    
	EMP_NO = Trim(Request("EMP_NO"))
    Call SubOpenDB(ADOConn)                                                        '��: Make  a DB Connection
	' �⺻���� ��ȸ 
    strSql = "Select NAME, DEPT_NM, MINOR_NM, TEL_NO, EM_TEL_NO, HAND_TEL_NO, ADDR, EMAIL_ADDR"
    strSql = strSql & " from HAA010T, B_MINOR"
    strSql = strSql & " where EMP_NO = " & EMP_NO & " AND (PAY_GRD1 = MINOR_CD AND MAJOR_CD='H0001')"
    
    If 	FncOpenRs("R",ADOConn,ADORs,strSql,"X","X") = False Then                    'If data not exists
        EmpDetail =  ""
    Else
		NAME         =ADORs("NAME")
		DEPT_NM      =ADORs("DEPT_NM")
		PAY_GRADE1   =ADORS("MINOR_NM")
		TEL_NO       =ADORs("TEL_NO")
		EM_TEL_NO    =ADORs("EM_TEL_NO")
		HAND_TEL_NO  =ADORs("HAND_TEL_NO")
		ADDR         =ADORs("ADDR")
		EMAIL_ADDR   =ADORS("EMAIL_ADDR")
    End If
    
    
    strSql = "Select MINOR_NM from HAA010T, B_MINOR"                 ' ������ ��ȸ 
    strSql = strSql & " where EMP_NO = " & EMP_NO & " AND (FUNC_CD = MINOR_CD AND MAJOR_CD='H0004')"
    
    If 	FncOpenRs("R",ADOConn,ADORs,strSql,"X","X") = False Then                    'If data not exists
        FUNC =  ""
    Else
		FUNC   =ADORS("MINOR_NM")
    End If
    
    strSql = "Select MINOR_NM from HAA010T, B_MINOR"                 ' ��å ��ȸ 
    strSql = strSql & " where EMP_NO = " & EMP_NO & " AND (ROLE_CD = MINOR_CD AND MAJOR_CD='H0026')"
    
    If 	FncOpenRs("R",ADOConn,ADORs,strSql,"X","X") = False Then                    'If data not exists
        ROLE =  ""
    Else
		ROLE = ADORS("MINOR_NM")
    End If
    
    
    Call SubCloseRs(ADORs)                                                          '��: Release RecordSSet
    Call SubCloseDB(ADOConn)                                                       '��: Colse a DB Connection													
    
    imageSrc = "../ComASP/CPictRead.asp" & "?txtKeyValue=" & EMP_NO '��: query key
    imageSrc = imageSrc     & "&txtDKeyValue=" & "default"                            '��: default value
    imageSrc = imageSrc     & "&txtTable="     & "HAA070T"                            '��: Table Name
    imageSrc = imageSrc     & "&txtField="     & "Photo"	                          '��: Field
    imageSrc = imageSrc     & "&txtKey="       & "Emp_no"	                          '��: Key
	
											
%>


'========================================================================================================
' Function Name : exitClick()
' Function Desc : ok image Ŭ�������� ó�� 
'========================================================================================================


function ExitClick()
	Self.Returnvalue = false
 	self.close()
end function


Sub Form_load()
	
	call ggoOper.LockField(Document, "Q")		
	Frm1.imgPhoto.src = "<%=imageSrc%>"
End Sub

Function FncExit()
    FncExit = True
End Function


function pgmjump1()
	Self.Returnvalue = true
	self.close()
end function

</script>


<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">

<!-- #Include file="../inc/uni2kcmcom.inc" -->	
</HEAD>

<BODY SCROLL=no TABINDEX="-1">
<form name=frm1>
<TABLE Class="BASICTB">
	<TR>
		<TD>
			<TABLE CLASS="TB3" CellSPACING=5 border="TAB11">
				<TR HEIGHT=40%>
					<TD WIDTH=30%>
						<TABLE "<%=LR_SPACE_TYPE_60%>" CELLSPACING=0 HEIGHT=40% >
							<TR><TD ALIGN=CENTER><img src="" name="imgPhoto" WIDTH=100 HEIGHT=150></TD></TR>
						</TABLE>
					</TD>
					<TD WIDTH=70%>
						<TABLE  "<%=LR_SPACE_TYPE_60%>" CELLSPACING=0 HEIGHT=40%>
							<TR>
								<TD class="TD5">����</TD>
								<TD Class="TD6" ><INPUT Type=Text size=22 t value="<%=NAME%>"  name=Text1 tag="24"></TD>
							</TR>
							<TR>
								<TD class="TD5">���</TD>
								<TD Class="TD6"><INPUT Type=Text size=22  value="<%=EMP_NO%>"  name=Text2 tag="24"></TD>
							</TR>
							<TR>
								<TD  class="TD5">��ȣ</TD>
								<TD  Class="TD6"><INPUT Type=Text size=22  value="<%=PAY_GRADE1%>" name=Text3 tag="24"></TD>
							</TR>
							
							<TR>
								<TD  class="TD5">�μ�</TD>
								<TD  Class="TD6"><INPUT Type=Text size=22  value="<%=DEPT_NM%>"  name=Text5 tag="24"></TD>
							</TR>
							<TR>
								<TD class="TD5">��å</TD>
								<TD Class="TD6"><INPUT  Type=Text size=22  value="<%=ROLE%>" name=Text4 tag="24"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR HEIGHT=60%>
					<TD COLSPAN=2>
						<Table "<%=LR_SPACE_TYPE_60%>" CELLSPACING=0>
							<TR>
								<TD  class="TD5">������</TD>
								<TD  Class="TD6"><INPUT Type=Text size=35  value="<%=FUNC%>"  name=Text6 tag="24"></TD>
										
							</TR>
							<TR>
								<TD  class="TD5">��ȭ��ȣ</TD>
								<TD  Class="TD6"><INPUT Type=Text size=35  value="<%=TEL_NO%>" name=Text7 tag="24">	</TD>
							</TR>
							<TR>
								<TD  class="TD5">��󿬶�</TD>
								<TD  Class="TD6" ><INPUT Type=Text size=35  value="<%=EMP_TEL_NO%>" name=Text8 tag="24"></TD>
							</TR>
							<TR>
								<TD  class="TD5">�ڵ���</TD>
								<TD Class="TD6" ><INPUT Type=Text size=35  value="<%=HAND_TEL_NO%>" name=Text9 tag="24"></TD>
							</TR>
							<TR>
								<TD class="TD5">���ּ�</TD>
								<TD  Class="TD6"><INPUT Type=Text size=35  value="<%=ADDR%>" name=Text10 tag="24"></TD>
							</TR>
							<TR>
								<TD class="TD5">E-Mail</TD>
								<TD Class="TD6" ><INPUT Type=Text size=35  value="<%=EMAIL_ADDR%>"  name=Text11 tag="24"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD>
			<TABLE Class="BASICTB">
				<TR>
					<TD WIDTH="*" ALIGN=left><a onclick= "VBSCRIPT:PgmJump1()"  >�λ縶��Ÿ</a></TD>
					<TD ALIGN=RIGHT>
						<IMG SRC="../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="VBSCRIPT:exitclick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../CShared/image/OK.gif',1)"></IMG>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
</TABLE>
</form>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

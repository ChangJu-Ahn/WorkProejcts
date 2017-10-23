<%@ page language="java" contentType="text/html; charset=utf-8"
    pageEncoding="utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1">
   <%@ include file="/WEB-INF/include/include-header.jspf" %>
</head>
<body>
	<div class="jbMenu">
		<div id='cssmenu'>
		<ul class="t_menu">
		   <li><a href='../sample/openBoardList.do'><span>LIST</span></a></li>
		   <!-- <li class = 'active has-sub'><a href='#'><span>계약서</span></a> -->
		   <li class = 'has-sub'><a href='#'><span>계약서</span></a>
		      <ul>
<!-- 	      		 <li><a href='../sample/openBoardWrite.do'><span>계약서 작성</span></a> -->
	      		 <li><a href='../sample/openBoardWrite.do'><span>체결계약서 등록</span></a>
	         	 </li>
<!-- 		         <li class='last'><a href='../sample/openBoardModify.do'><span>변경합의서 작성</span></a> -->
		         <li class='last'><a href='../sample/openBoardModify.do'><span>변경합의서 등록</span></a>
		         </li>
		      </ul>
		   </li>
		   <sec:authorize access="hasRole('ROLE_ADMIN')" var="ROLE_ADMIN"/>		   
		   <li><a href="../report/reportView.do"><span>보고서</span></a></li>
		   <li class ='has-sub'><a href='#'><span>환경설정</span></a>
		      <ul>
		      		<c:if test="${ROLE_ADMIN == true}"> <!--계약서 기준정보, 사용자 관리는 관리자만 볼 수 있도록-->
			      		<li><a href="../admin/adminView.do"><span>계약서 기준정보</span></a></li>
	<!-- 		      		<li><a href="../admin/admin_SysInfo.do"><span>시스템 기준정보</span></a></li> -->
			      		<li><a href="../admin/openAdmin_account.do"><span>사용자 관리</span></a></li>
		      		</c:if>
		      		<li class='last'><a href="#this" id="InfoChange"><span>패스워드 변경</span></a></li>	
		      </ul>
		   </li>
		   <li class='last'><a href='#' id ="logout"><span>로그아웃</span></a></li>
		</ul>
		</div>
	</div>	
	<br/>
	<br/>
	현재 로그인 ID : <sec:authentication property="principal.username"/>
	<div id ="category_div"></div>
	<script type="text/javascript">
        $(document).ready(function(){
        	$('#logout').on('click', function() {
        		Lobibox.confirm({
        		    msg: "로그아웃 하시겠습니까?",
        		    callback: function ($this, type, ev) {
        		        if(type == "yes"){
        		    		location.href = "../login/logout.do";
        		        }
        		    }
        		});   
        	});
        	
        	$('#InfoChange').on('click', function(e) {
        		e.preventDefault();	
        		fn_popOpen();
        	});
        	
        });
                
        function fn_active(){
        	$.ajax({  
                type: "POST",  
                url: "../admin/adminView.do",  
                dataType: "html",    //받는 방식  
                success: function (data, txtStatus) {  
                	$("#category_div").html(data); //str로 받은 data를 넘긴다.  
                },  
                error: function (xhr, txtStatus, errorThrown) {  
                    alert("error" + errorThrown);  
                }  
            });  
        	//$('ul').children('li').attr('class','')
    		//location.href = url;
    		//$('ul').children('li').eq(seq).attr('class','active')
    	}
        
        function fn_popOpen(){
        	var popOption = "width=370, height=700, resizable=no, scrollbars=yes, status=no;";
    		var popUrl = "../sample/openUserinfoChange.do";
    		var newWindow = window.open(popUrl, "", popOption);
    		
    		newWindow.focus();
        }
        
        
		
     </script>
</body>



</html>
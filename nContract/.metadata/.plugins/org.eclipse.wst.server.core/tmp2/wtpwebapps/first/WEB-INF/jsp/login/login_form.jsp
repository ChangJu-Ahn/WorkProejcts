<%@ page language="java" contentType="text/html; charset=utf-8"
    pageEncoding="utf-8"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%@ include file="/WEB-INF/include/include-header.jspf"%>

<!-- 로그인만 특별하게 CSS, JS 관리 -->
<link rel="stylesheet" type="text/css" href="../css/login/style.css">
<link rel="stylesheet" type="text/css" href="../css/login/reset.css" >
<link rel="stylesheet" href="../css/lobibox.min.css"/>

<script src="//code.jquery.com/jquery-1.11.3.min.js"></script>
<script src="../js/common.js" charset="utf-8"></script>
<script src="../js/lobibox.min.js"></script>

<title>로그인 페이지</title>
</head>
<body>	
<c:choose>
	<c:when test = "${error == 'true'}">
		<script type="text/javascript">
		 Lobibox.alert("error", //AVAILABLE TYPES: "error", "info", "success", "warning"
	    		 {
			 		 msg: "로그인 실패 (Id, Password를 확인 해주세요.)"
	    		 });
		 </script>
	</c:when>
	 <c:otherwise>
	<!-- 정상로직일 경우 하단부의 form생성 후 비밀번호 변경창을 생성 함 -->
	 </c:otherwise>
</c:choose>	
<form id="frm" name="frm" >
	<div class="pen-title">
		<img alt="메인로고" src="../images/main.png">	
	</div>	
	<div class="container">
		<div class="card"></div>
		<div class="card">
			<h1 class="title">Login</h1>
			<!-- <form> -->
				<div class="input-container">
					<input type="text" name ="j_username"/> 
					<label for="Username">사번</label>
					<div class="bar"></div>
				</div>
				<div class="input-container">
					<input type="password" name ="j_password" onkeyPress="checkCapsLock(event)"/>
					<label for="Password">비밀번호</label>
					<div class="bar"></div>
					<img alt="imgCapsLock" id="imgCapsLock" src="../images/imgCapsLock.png" style="visibility: hidden;">
				</div>
				<div class="button-container">
					<button id = "btn">
						<span>로그인</span>
					</button>
				</div>
				<!-- <div class="footer">
					<a href="#">Forgot your password?</a>
				</div> -->
		</div>
<!-- 		<div class="card alt"> -->
<!-- 			<div class="toggle"></div> -->
<!-- 			<h1 class="title"> -->
<!-- 				패스워드 변경 -->
<!-- 				<div class="close"></div> -->
<!-- 			</h1> -->
<!-- 			<!-- <form> --> 
<!-- 				<div class="input-container"> -->
<!-- 					<input type="text" id="USER_NO" />  -->
<!-- 					<label for="Username">ID(사번)</label> -->
<!-- 					<div class="bar"></div> -->
<!-- 				</div> -->
<!-- 				<div class="input-container"> -->
<!-- 					<input type="password" id="USER_PWD" />  -->
<!-- 					<label for="Password">비밀번호</label> -->
<!-- 					<div class="bar"></div> -->
<!-- 				</div> -->
<!-- 				<div class="input-container"> -->
<!-- 					<input type="password" id="REPEAT_PASSWORD" /> -->
<!-- 					<label for="Repeat Password">비밀번호 확인</label> -->
<!-- 					<div class="bar"></div> -->
<!-- 				</div> -->
<!-- 				<div class="input-container"> -->
<!-- 					<input type="password" id="CHANGE_PASSWORD" /> -->
<!-- 					<label for="Repeat Password">변경 비밀번호</label> -->
<!-- 					<div class="bar"></div> -->
<!-- 				</div> -->
<!-- 				<div class="button-container"> -->
<!-- 					<button id="change"> -->
<!-- 						<span>변경</span> -->
<!-- 					</button> -->
<!-- 				</div> -->
<!-- 			<!-- </form> --> 
<!-- 		</div> -->
	</div>
</form>
	
	<%@ include file="/WEB-INF/include/include-body.jspf" %>
	<script type="text/javascript">
        $(document).ready(function(){
        	$('.toggle').on('click', function() {
        		  $('.container').stop().addClass('active');
        		});

       		$('.close').on('click', function() {
       		  $('.container').stop().removeClass('active');
       		});
       		
       		$("#btn").on('click', function(e){
       			e.preventDefault();
       			fn_openChkLogin();
       		});
       		
       		$("#change").on('click', function(e){
       			e.preventDefault();
       			fn_openChkChangePasswrd();
       		});	
        });
        
        function fn_openChkChangePasswrd(){
        	if($("input[id = 'USER_NO']").val().length == 0){
        		Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
       	    		 {
       	    		     msg: "[ID(사번)]란을 입력해주세요."
       	    		 });
        	}else if($("input[id = 'USER_PWD']").val().length == 0){
        		Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
       	    		 {
       	    		     msg: "[비밀번호]란을 입력해주세요."
       	    		 });
        	}else if($("input[id = 'REPEAT_PASSWORD']").val().length == 0){ 
        		Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
          	    		 {
          	    		     msg: "[비밀번호 확인]란을 입력해주세요."
          	    		 });
        	}else if($("input[id = 'CHANGE_PASSWORD']").val().length < 5){
        		Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
          	    		 {
          	    		     msg: "변경될 비밀번호는 최소 6자리 이상입니다."
          	    		 });
        	}
        	else{
         		var comSubmit = new ComSubmit("frm");
                 comSubmit.setUrl("<c:url value='/admin/initUserInfo.do'/>");
                 comSubmit.addParam("type", "U");			//진행 타입 -> I: 초기화(initial), U: 수정(update)
                 
                 comSubmit.submit();   
        	}                     
        }
        
        function fn_openChkLogin(){
        	if($("input[name = 'j_username']").val() == ""){
        		Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
       	    		 {
       	    		     msg: "사번을 입력해주세요."
       	    		 });
        	}else if($("input[name = 'j_password']").val() == ""){
        		Lobibox.alert("warning", //AVAILABLE TYPES: "error", "info", "success", "warning"
       	    		 {
       	    		     msg: "비밀번호를 입력해주세요."
       	    		 });
        	}else{
        		var comSubmit = new ComSubmit("frm");
                comSubmit.setUrl("j_spring_security_check");
                
                comSubmit.submit();   
        	}                     
        }
        
        function checkCapsLock(e){
        	var myKeyCode=0;
        	var myShiftKey=false;
        	
            if ( document.all ) {	// Internet Explorer 4+
                  myKeyCode=e.keyCode;
                  myShiftKey=e.shiftKey;

            } else if ( document.layers ) {	// Netscape 4
                  myKeyCode=e.which;
                  myShiftKey=( myKeyCode == 16 ) ? true : false;

            } else if ( document.getElementById ) {	// Netscape 6
                  myKeyCode=e.which;
                  myShiftKey=( myKeyCode == 16 ) ? true : false;
            }

            if ((myKeyCode >= 65 && myKeyCode <= 90 ) && !myShiftKey) $("#imgCapsLock").css("visibility","visible");
            else if ((myKeyCode >= 97 && myKeyCode <= 122 ) && myShiftKey) $("#imgCapsLock").css("visibility","visible");
           	else $("#imgCapsLock").css("visibility", "hidden");
        }
    </script> 
</body>

 	<!-- <form action="j_spring_security_check" method='POST'>
        <label>아이디</label>
        <input type="text" name="j_username">
        <label>패스워드</label>
        <input type="password" name="j_password">
        <button id ="btn">로그인</button>  
    </form>
    </body> -->

</html>
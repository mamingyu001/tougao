<!-- #include file="conn.asp" -->


<head>
<title>管理员登录</title>
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body background="images/123456.jpeg">  
<% if request("login")<>1 and session("loginok")<>"yes"then%>  
<form method=post action="admin.asp">  
<table width="100%" align="center" border="0" cellspacing="0" cellpadding="3">  
<tr align=center><td></td></tr>  
<tr>   
<td>
  <p align="center">管理员名称：<input name=user_account  class=ourfont size=20>  
<br>管理员密码：<input name=user_passwd type=password class=ourfont size=20>  
           <br>  
  </p>
            </td>  
          </tr>  
          <tr align="center">   
            <td>   
               <input type=submit value='登 录'></td></tr>  
        </table>  
  
 <INPUT TYPE="hidden" name="login" value=1>  
 </form>        
<% else  
    if session("loginok")<>"yes" then  
         set rs=Server.CreateObject("Adodb.recordset")  
         username=request.form("user_account")  
         password=request.form("user_passwd")  
         sql="select * from admin where user ='"&username &"'"  
         set rs=conn.execute(sql)  
         if not(rs.eof) then  
            if password=rs("password") then  
	           session("loginok")="yes"	     
	        else  
%>	  
          <script>  
		  alert("密码错误！\n返回");  
		  history.back();  
		  </script>  
<%	        end if  
          else  
%>	  
          <script>  
		  alert("帐号不存在！\n返回");  
		  history.back();  
		  </script>  
<%  
          end if  
		  rs.close         
     end if  
	  if session("loginok")="yes" then  
         set rs=server.createobject("adodb.recordset")  
		 set rs2=server.createobject("adodb.recordset")  
         sql="select * from mail order by id desc"  
		 sql2="select * from admin"  
         rs.open sql,conn,3  
         rs2.open sql2,conn,3,2  
		 if request("types")="" then  
		 types=rs2("writetype")  
		 else  
		 rs2("writetype")=request("types")  
         rs2.update  
		 types=rs2("writetype")  
         end if  
		   
         if not(rs.eof) then  
            rs.pagesize=rs2("pagesize")  
            if isnumeric(request("page")) then  
               page=cint(request("page"))  
               if page<1 or page>rs.pagecount then page=1  
            else  
            page=0  
         end if  
         rs.absolutepage=page  
         maxpage=rs.pagecount  
%>  
  
 <p align="center"><a href="modpassword.asp">更改管理密码</a> | 管理稿件 |              
 <a href="setpagesize.asp">设定每页显示记录数</a> | <a href="logout.asp">退出管理</a></p>                                         
                                         
<center>                                         
<TABLE >                                         
<tr>                                         
                                         
<TD align=right >共<% =maxpage%>页<% =rs.RecordCount %>篇                                              
           <% if page>1 then %><a href="admin.asp?login=1&page=<% =page-1 %>"><font color="#FF0000">前一页</font></a>                                           
      <% end if %>                                              
	  <% for k=1 to rs.pagecount %>                                              
	  <% if k=page then %> &nbsp;第<% =k %>页                                               
                   <%end if%>                                           
	  <% next %>                                              
	  <% if page<rs.pagecount then %>                                              
	  <a href="admin.asp?login=1&page=<% =page+1 %>"><font color="#FF0000">后一页</font></a>                                               
      <% end if %></TD>                                              
                                          
</tr>                                          
</TABLE>                                          
        </center>                                          
<% for j=1 to rs.pagesize %>                                           
<center>                                          
<table width="496" cellspacing="0" cellpadding="0">                                          
<tr>                                          
 <td width="100%"  height="23">                                          
<p align="center"><span>&nbsp;<img border="0" src="images/goto.gif"> 投稿人姓名：<% =rs("name")%>&nbsp;</span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="delmessage.asp?id=<% =rs("id")%>"><img border="0" src="images/delete.gif" width="17" height="17" alt="删除">删除</a></p>                                        
 </td>                                        
</tr>                                        
<tr>                                        
 <td width="50%"  height="20">&nbsp; 联系方式：<% =rs("address")%>　
  </td>
</tr>
<tr>                                    
 <td width="50%"  height="20">&nbsp; 电子信箱：<% =rs("email")%>　
  </td>
</tr>
<tr>                                    
 <td width="50%"  height="20">&nbsp; 文章标题：<% =rs("title")%>　
  </td>
</tr>
<tr>
 <td width="100%" height="100" valign="top" style="padding-left: 15; padding-right: 15; padding-top: 10; padding-bottom: 10"><TEXTAREA NAME="" ROWS="10" COLS="65"><% =rs("comment")%></TEXTAREA></td>
</tr>
<tr align=center>

 <td width="100%"  height="25"  align=center><FONT  COLOR="red">投稿日期：<% =rs("datenow")%>
 &nbsp;&nbsp;&nbsp;&nbsp;&nbsp; FROM：<% =rs("ip")%></FONT>     
 </td>       
</tr>       
</table>     
</center><BR>     
<%     
   rs.movenext     
     if rs.eof then     
     exit for     
     end if     
   next     
%>     
<BR><center>     
        共<% =maxpage%>页<% =rs.RecordCount %>篇        
           <% if page>1 then %><a href="admin.asp?page=<% =page-1 %>"><font color="#FF0000">前一页</font></a>      
      <% end if %>        
	  <% for k=1 to rs.pagecount %>        
	  <% if k=page then %> &nbsp;第<% =k %>页         
                   <%end if%>      
	  <% next %>        
	  <% if page<rs.pagecount then %>        
	  <a href="admin.asp?page=<% =page+1 %>"><font color="#FF0000">后一页</font></a>         
      <% end if %> <BR>        
        </center>     
<%     
rs.close     
conn.close     
else      
  rs.close     
 response.write "目前没有稿件!"     
end if     
%>     
<%  end if                                             
                                          
end if                                 		                                          
%>                                           
                                          
<BR>                                          
                                        
</body>

<head>
<title></title>
</head>
<body BGCOLOR="bbddbb">
<%
  response.write("<center>")
  response.write("<h2>���A���w�ЪŶ��d��</h>")
  response.write("<table border= 1 width = 100% >")
      response.write("<tr>")
      response.write("<td align=center>���A��")
      response.write("</td>")
      response.write("<td align=center>�Ϻо�")
      response.write("</td>")
      response.write("<td align=center>�����Ŷ�(GB)")
      response.write("</td>")
      response.write("<td align=center>�Ѿl�Ŷ�(GB)")
      response.write("</td>")
      response.write("<td align=center>�Ѿl�Ŷ����%")
      response.write("</td>")
      response.write("<td align=center>�^���ɶ�")
      response.write("</td>")
      response.write("</tr>")

  Set DB = Server.CreateObject("ADODB.Connection")
  DB.OPEN "DRIVER={SQL SERVER};SERVER=CXLSVR55;DATABASE=svrmanage;UID=isvrmanage;PWD=5678"
  SQL = "Select * from CurrDisk order by computername,diskname"
  SET RS = db.EXECUTE(SQL)
   do while not rs.eof
      computername = rs("computername")
      diskname = rs("diskname")
      TotalSpace = rs("TotalSpace")
	  FreeSpace = rs("FreeSpace")
      FreePercent = rs("FreePercent")
      DateTime = rs("DateTime")
	  if TotalSpace > 0 then
         response.write("<tr>")
         response.write("<td align=center>" & computername)
         response.write("</td>")
         response.write("<td align=center>" & diskname)
         response.write("</td>")
         response.write("<td align=center>" & TotalSpace)
         response.write("</td>")
   	     if FreeSpace >=0.5 then
            response.write("<td align=center >" & FreeSpace)
	     else if FreeSpace >=0.2 then
                 response.write("<td align=center bgcolor=yellow>" & FreeSpace)
	          else
                 response.write("<td align=center bgcolor=#ff0000#>" & FreeSpace)
		      end if
	     end if
         response.write("</td>")
	     if FreePercent >= 5 then
            response.write("<td align=center>" & FreePercent)
         else if FreePercent >= 2 then
                response.write("<td align=center bgcolor=yellow>" & FreePercent)
	        else
                response.write("<td align=center bgcolor=#ff0000#>" & FreePercent)
		    end if
	     end if
         response.write("</td>")
	     if day(date) = day(DateTime) then 
            response.write("<td align=center>" & DateTime)
         else
            response.write("<td align=center bgcolor=#ff0000#>" & DateTime)
   	     end if
         response.write("</td>")
         response.write("</tr>")
	   end if
       RS.movenext
   loop
   rs.close
   db.close
   set rs = nothing
   set db = nothing

%>
</body>
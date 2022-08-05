<html>
    <head>
        <meta content="charset=utf-8" />
    </head>
    <body>
        <%
        '设置文件路径'
        dim fn_ans
            fn_ans="C:\Users\panha\Desktop\KwapWebSurvey\answers.txt"
        '创建FileSystemObject'
        dim fso
            set fso=Server.CreateObject("Scripting.FileSystemObject")
        '从末尾打开文件'
        dim tso_end
            set tso_end=fso.OpenTextFile(fn_ans,8,true,-1)
        '写入'
        if Request.Form("way_else") <> "" then
            tso_end.WriteLine("03.方向: " & Request.Form("way") & "其他(" & Request.Form("way_else") & ")")
        else
            tso_end.WriteLine("03.方向: " & Request.Form("way"))
        end if
        tso_end.Close
        response.redirect("4.html")
        %>
    <p style="text-align:center; font-size:50px; font-weight:bold">
        Loading...
    </p>
    
    </body>
</html>
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
        '从头打开文件'
        dim tso_head
            set tso_head=fso.OpenTextFile(fn_ans,1)
        '从末尾打开文件'
        dim tso_end
            set tso_end=fso.OpenTextFile(fn_ans,8,true,-1)
        '获取用户次序'
            dim usernumber
            '跳过前面部分'
            do while tso_head.AtEndOfStream = false
                tso_head.SkipLine
            loop
            '计算用户次序'
            usernumber=(tso_head.Line+8)\9
        '写入'
        tso_end.WriteLine("User " & usernumber & ":")
        if Request.Form("location") = "其他" then
            tso_end.WriteLine("01.获取到本调查问卷链接的位置: 其他(" & Request.Form("location_else") & ")")
        else
            tso_end.WriteLine("01.获取到本调查问卷链接的位置: " & Request.Form("location"))
        end if
        tso_end.Close
        if Request.Form("location") = "7B" then
            response.redirect("2-7B.html")
        else 
            response.redirect("2-else.html")
        end if
        %>
    <p style="text-align:center; font-size:50px; font-weight:bold">
        Loading...
    </p>
    </body>
</html>
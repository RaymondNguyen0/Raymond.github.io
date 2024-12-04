
    <%
    ' Retrieve the background color from the query string
    Dim bgColor
    bgColor = Request.QueryString("color")

    ' Check if the user has a cookie for the last visit
    Dim lastVisit, firstVisit
    lastVisit = Request.Cookies("LastVisit")

    If lastVisit = "" Then
        firstVisit = True
        lastVisit = "This is your first visit!"
    Else
        firstVisit = False
        lastVisit = FormatDateTime(lastVisit, vbLongDate) & " at " & FormatDateTime(lastVisit, vbLongTime)
    End If

    ' Set a cookie for the current visit
    Response.Cookies("LastVisit") = Now()  ' Note: currentDateTime should be Now()
    Response.Cookies("LastVisit").Expires = DateAdd("d", 30, Now()) ' Expires in 30 days
    %>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dynamic Background Color</title>
    <style>
        body {
            color: #333;
            font-family: Arial, sans-serif;
            text-align: center;
            padding: 20px;
        }
        .message {
            margin-top: 50px;
            font-size: 1.2em;
        }
    </style>
</head>
<body style="background-color: <%= bgColor %>">
    <h1>Welcome to the Dynamic Background Page!</h1>
    <div class="message">
        <% If firstVisit Then %>
            <p><strong>Welcome! This is your first visit.</strong></p>
        <% Else %>
            <p>Your last visit was on: <strong><%= lastVisit %></strong></p>
        <% End If %>
    </div>
    <p>To change the background color, append <code>?color=COLOR_NAME_OR_HEX</code> to the URL.</p>
    <p>For example: <code>?color=lightblue</code> or <code>?color=#FF5733</code></p>
</body>
</html>
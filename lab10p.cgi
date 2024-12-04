#!/usr/bin/env python3
import cgi

print("Content-Type: text/html\n")

form = cgi.FieldStorage()
city = form.getvalue('city', '').upper()
province = form.getvalue('province', '').upper()
country = form.getvalue('country', '').upper()
image_url = form.getvalue('image_url', '')

print(f"""
<html>
<head><title>City Information</title></head>
<body style="text-align: center; background-color: lightblue;">
    <h1 style="font-size: 3em; color: red;">{city}, {province}, {country}</h1>
    <div style="border: 10px solid black; display: inline-block; margin-top: 20px;">
        <img src="{image_url}" style="width: 80%; height: auto;">
    </div>
</body>
</html>
""")

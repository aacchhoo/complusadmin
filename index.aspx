<%@ Page Language="C#" AutoEventWireup="true" CodeFile="index.aspx.cs" Inherits="index" %>  
  
<!DOCTYPE html>  
<html lang="en">  
<head>  
    <meta charset="UTF-8">  
    <meta name="viewport" content="width=device-width, initial-scale=1.0">  
    <title>COM+ Object Control</title>  
    <script>  
        async function manageComPlusObject(action) {  
            try {  
                const response = await fetch(`index.aspx?action=${action}`, {  
                    method: 'POST'  
                });  
                const result = await response.text();  
                document.getElementById('result').innerText = result;  
            } catch (error) {  
                document.getElementById('result').innerText = 'Error: ' + error.message;  
            }  
        }  
    </script>  
</head>  
<body>  
    <h1>COM+ Object Control</h1>  
    <button onclick="manageComPlusObject('stop')">Stop LebbermanREDCom</button>  
    <button onclick="manageComPlusObject('restart')">Restart LebbermanREDCom</button>  
    <div id="result"></div>  
</body>  
</html>  

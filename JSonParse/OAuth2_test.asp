
<%@Language="VBScript" %>

<script language="JScript" runat="server" src='json2.js'></script>

<%

Function getAccessToken(sGrant_type, sUser, sPassword, sPostUrl)

    dim postData, ServerXmlHttp, encodedAuthHeader, jsonResponse, myJson, access_token
    
    'Parámetros a enviar al servicio 
    'Ej.: grand_type=password&username=user&password=222
    postData = "grant_type=" & sGrant_type & "&username=" & sUser & "&password=" & sPassword    'Datos de acceso

    Set ServerXmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
    ServerXmlHttp.open "POST", sPostUrl
    ServerXmlHttp.setRequestHeader "Content-Type", "application/json;charset=utf-8"
    ServerXmlHttp.setRequestHeader "Authorization", "Basic " & encodedAuthHeader
    ServerXmlHttp.setRequestHeader "Content-Length", Len(PostData)
    ServerXmlHttp.send PostData

    jsonResponse = ServerXmlHttp.responseText

    'Implementamos la utilería json2.js para poder parsear el JSon de respuesta
    set myJson = JSON.parse(jsonResponse)

    'Ej.: "bearer 5F8iyiIxmd6jETQUlFRxK1id3loz2FLqTc4COlh_f_IIe3..."
    access_token = myJson.token_type &" "& myJson.access_token

    getAccessToken = access_token

    Set ServerXmlHttp = Nothing

End Function  

'En esta función vamos enviar el Token al WebAPI para que valide la petición
'Como respuesta tendremos un mensaje que indicará si el token fue válido o no
Function getAccessAPI(sAccesToken, sGetUrl)

    
    dim ServerXmlHttp, strResponse    

    sGetUrl = sGetUrl & "?Authorization=" & sAccesToken    

    Set ServerXmlHttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
    ServerXmlHttp.open "GET", sGetUrl, false
    ServerXmlHttp.setRequestHeader "Content-Type", "application/json;charset=utf-8"
    
    ServerXmlHttp.send 

    strResponse = ServerXmlHttp.responseText    

    getAccessAPI = strResponse

    Set ServerXmlHttp = Nothing    

End Function

%>

<%
    dim sGrant_type, sUser, sPassword, sToken, sPostUrl

    'Parámetros de usuario    
    sGrant_type = "password"
    sUser = "admin"
    sPassword = "123"
    sPostUrl = "http://localhost:58186/recuperartoken"

    'Hacemos la petición del Token al WebAPI 
    sToken = getAccessToken(sGrant_type, sUser, sPassword, sPostUrl)
    
    Response.Write("ACCESS TOKEN: <br>" & sToken)

    Response.Write("<br/><br/>")  

    dim validationMessage, sGetUrl
    sGetUrl = "http://localhost:58186/api/usuario"

    'Enviamos al WebAPI el token para validar al usuario
    validationMessage = getAccessAPI(sToken, sGetUrl)

    'Imprimimos el mensaje de validación del WebAPI
    Response.Write("Validation: " & validationMessage)
%>



<%

    <!-- Paso de parámetros entre páginas POST-->    
    dim name, idRol, rolDescription
    name=Request.form("user") 
    idRol=Request.form("rol")

    IF name <> "" THEN        

        SELECT CASE idRol
            CASE 1
            rolDescription = "Usuario"
            CASE 2
            rolDescription = "Desarrollador"
        END SELECT

        ELSE    <!--No se introdujo el campo usuario -->
            Response.Redirect("Login.asp?Error=1")  <!-- Redireccionamos a Login pero con 1 parámetro GET -->
    END IF    

%>

<html>

<head>
    <script  language="javascript" src="jsFunctions.js" runat="server"></script>
    <!-- <script type="text/javascript" src="jsFunctions.js"></script> -->
    
</head>

<body>  

    <table align="center" border="0" width="300px">
        <tr>
            <td>
                Nombre: 
            </td>
            <td>
                <%= Name %> <!-- BVScript - NO es Case Sensitive ejemplo con la variable name - Name -->    
            </td>
        </tr>
        <tr>
            <td>
                Rol: 
            </td>
            <td>
                <%= rolDescription %>
            </td>
        </tr>
        <tr>
            <td>
                Identificador: 
            </td>
            <td>
                <%
                    sResult = createID()
                    Response.Write sResult
                %>
            </td>
        </tr>
    </table>     

    <%             
        response.write("<script> var idSesion = createID();")        
        response.write("alert(idSesion); ")        
        response.write("</script>")        
    %>              

</body>

</html>
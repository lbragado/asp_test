<html>
    <head>
    </head>
    <body>        
        <form action=Main.asp method=post>
        <table align="center" border="0" width="300px">
            <tr>
                <td colspan="2" align="center">
                    <b> LOGIN </b>
                </td>
            </tr>
            <tr><td colspan="2">&nbsp;</td>></tr>
            <tr>
                <td>
                    User:
                </td>
                <td align="right">
                    <input type=text name="user" />
                </td>
            </tr>
            <tr>
                <td>
                    Rol:
                </td>
                <td align="right">
                    <select name="rol">
                        <option selected value="1">Usuario</option>
                        <option value="2">Desarrollador</option>
                    </select>
                </td>
            </tr>
            <tr><td colspan="2">&nbsp;</td>></tr>
            <tr>
                <td colspan="2" align="right">
                    <input type=submit value=LOGIN name=submit />
                </td>
            </tr>
            <tr><td colspan="2">&nbsp;</td>></tr>
            <tr>
                <td colspan="2">
                    <%
                        <!-- Recuperando parámetro GET -->
                        IF Request.QueryString("Error") = "1" THEN
                            Response.Write("*El campo Usuario es obligatorio.")                                                
                        END IF
                    %>
                </td>
            </tr>
        </table>
    </form>
        </form>
    </body>
</html>
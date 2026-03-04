<%@ Page Language="C#" CodeFile="CargaMasivaFacturas.aspx.cs" Inherits="CargaMasivaFacturas_aspx" MasterPageFile="~/MasterPage.master" MaintainScrollPositionOnPostback="true" %>

<%@ MasterType VirtualPath="~/MasterPage.master" %>
<%@ Register Assembly="DevExpress.Web.v12.2, Version=12.2.8.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxCallback" TagPrefix="dxcb" %>
<%@ Register Assembly="DevExpress.Web.v12.2, Version=12.2.8.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxRoundPanel" TagPrefix="dxrp" %>
<%@ Register Assembly="DevExpress.Web.v12.2, Version=12.2.8.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxPanel" TagPrefix="dxp" %>
<%@ Register Assembly="DevExpress.Web.v12.2, Version=12.2.8.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxGridView" TagPrefix="dxwgv" %>
<%@ Register Assembly="DevExpress.Web.v12.2, Version=12.2.8.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dxe" %>
<%@ Register Assembly="DevExpress.Web.v12.2, Version=12.2.8.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxEditors" TagPrefix="dx" %>
<%@ Register Assembly="DevExpress.Web.v12.2, Version=12.2.8.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxUploadControl" TagPrefix="dxuc" %>

<%@ Register Src="../UserControls/StaticsBlanksCV.ascx" TagName="BlanksCV" TagPrefix="uc1" %>

<%@ Register Assembly="DevExpress.Web.ASPxHtmlEditor.v12.2, Version=12.2.8.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxHtmlEditor" TagPrefix="dxhe" %>
<%@ Register Assembly="DevExpress.Web.ASPxSpellChecker.v12.2, Version=12.2.8.0, Culture=neutral, PublicKeyToken=b88d1754d700e49a" Namespace="DevExpress.Web.ASPxSpellChecker" TagPrefix="dxwsc" %>




<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="Server">
    <link rel="Stylesheet" href="../JQuery/css/Main.css" />
    <link rel="Stylesheet" href="../JQuery/css/Impromptu.css" />
    <script language="javascript" src="../JQuery/js/jquery-1.4.2.min.js"></script>
    <script language="javascript" src="../JQuery/js/jquery-impromptu.3.1.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>



    <style type="text/css">
        .style1 {
            height: 9px;
        }

        .auto-style1 {
            width: 141px;
        }

        .auto-style2 {
            width: 91px;
        }

        .auto-style3 {
            width: 107px;
        }

        .auto-style4 {
            width: 247px;
        }

        .auto-style6 {
            width: 4px;
        }

        .auto-style9 {
            width: 120px;
        }

        .auto-style10 {
            width: 110px;
        }

        .auto-style11 {
            width: 96px;
        }

        .auto-style12 {
            width: 85px;
        }

        .vertical-align-content {
            display: flex;
            vertical-align: middle;
            text-align: left;
        }
    </style>

   <script type="text/javascript">
       var downloadTimer;
      
       function iniciarDescarga() {
           // A. Mostrar el Spinner de Carga
           Swal.fire({
               title: 'Generando Excel...',
               html: 'Por favor espere, esto puede tardar unos segundos.',
               allowOutsideClick: false, 
               didOpen: () => {
                   Swal.showLoading(); 
               }
           });

           
           document.cookie = "downloadStatus=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/;";

           
           downloadTimer = setInterval(function () {
               var cookieValue = getCookie("downloadStatus");

               
               if (cookieValue == "completed") {
                   finalizarDescarga();
               }
           }, 1000); 
       }

      
       function finalizarDescarga() {
           
           clearInterval(downloadTimer);

          
           Swal.fire({
               icon: 'success',
               title: '¡Descarga Lista!',
               text: 'El archivo se ha descargado correctamente.',
               confirmButtonText: 'Aceptar',
               confirmButtonColor: '#28a745' 
           });

        
           document.cookie = "downloadStatus=; expires=Thu, 01 Jan 1970 00:00:00 UTC; path=/;";
       }
              
       function getCookie(name) {
           var value = "; " + document.cookie;
           var parts = value.split("; " + name + "=");
           if (parts.length == 2) return parts.pop().split(";").shift();
       }
    </script>

</asp:Content>

<asp:Content ID="Content" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">


    <table width="100%">
        <tr>
            <td class="titulo" style="width: 100%">
                <table style="width: 100%">

                    <tr style="display: flex">
                        <td class="auto-style8" style="flex-grow: 1">
                            <dxe:ASPxLabel ID="ASPxLabelTitulo" runat="server" Text="Carga Masiva de facturas" Font-Bold="true"></dxe:ASPxLabel>
                        </td>
                        <td style="width: 25px"></td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td>
                <dxe:ASPxLabel ID="ASPxLabelMsg" runat="server" CssClass="MensajeAlert" Width="100%" Visible="false">
                </dxe:ASPxLabel>
            </td>

        </tr>
    </table>
    <table>
        <tr>
            <td colspan="2">1.	Descargue la siguiente plantilla: recuerde que los campos con “*” son obligatorios para rellenar.</td>
        </tr>
        <tr>
            <td style="height: 20px"></td>
        </tr>
        <tr>
            <td align="right">&nbsp &nbsp &nbsp &nbsp &nbsp &nbsp
                <dxe:ASPxButton ID="ASPxButtonNew" runat="server" Cursor="hand" ToolTip="Descargar Plantilla" HorizontalAlign="Left" OnClick="GeneraExcel" BackColor="White" Height="20px" Width="20px" Style="display: inline-block; margin: 0 auto">
                    <ClientSideEvents Click="function(s, e) { iniciarDescarga(); }" />
                    <Image Height="20px" Width="20px"></Image>
                    <BorderRight BorderColor="White" />
                    <BorderLeft BorderColor="White" />
                    <Border BorderColor="White" />
                    <BorderTop BorderColor="White" />
                    <BorderBottom BorderColor="White" />
                    <BackgroundImage ImageUrl="../images/drive-download.png" Repeat="NoRepeat" />
                </dxe:ASPxButton>
                &nbsp &nbsp &nbsp 
                <dxe:ASPxLabel ID="ASPxLabel1" runat="server" Text="Descargar Plantilla." Style="display: inline-block; margin: 0 auto"></dxe:ASPxLabel>
            </td>
            <td align="left"></td>
        </tr>
        <tr>
            <td style="height: 20px"></td>
        </tr>
        <tr>
            <td colspan="2" style="height: 10px"></td>
        </tr>
        <tr>
            <td colspan="2">2. Suba la plantilla rellena con la información de las facturas y los documentos asociados a las facturas</td>
        </tr>
        <tr>
            <td style="color: red"><strong>Recuerde que el nombre que se introduzca en la columna de documento debe tener el mismo nombre que el archivo que se asocie a la factura</strong></td>
        </tr>

        <tr>
            <td><b>Plantilla</b></td>
            <td><b>Ficheros</b></td>
        </tr>
        <tr>
            <td>

                <asp:FileUpload ID="FileUpload" runat="server" />

            </td>
            <td>
                <input type="file" id="myfile" multiple="multiple" name="myFile[]" runat="server" />


            </td>
        </tr>
        <tr>
            <td colspan="2" style="height: 40px"></td>
        </tr>
        <tr>
            <td colspan="2">El tamaño máximo de todos los archivos adjuntos no debe superar el tamaño de  143MB</td>
        </tr>
    </table>
    <table width="100%">
        <tr>
            <td colspan="2" style="height: 30px"></td>
        </tr>
        <tr>
            <td align="center">
                <table width="100%">
                    <tr>
                        <td width="40%">&nbsp;</td>
                        <td width="40px">
                            <asp:Button runat="server" ID="validarButton" Text="Validar" OnClick="ASPxButtonValidar_Click" CssClass="btn btn-mfast-3" Height="41px" Width="139px" />
                        </td>
                    </tr>
                    <tr>
                        <td>&nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <table width="100%">
        <tr align="left">
            <td class="titulo" style="width: 1180px">
                <table width="100%">
                    <tr>
                        <td>
                            <dx:ASPxLabel ID="ASPxLabel5" runat="server" Font-Bold="true" Text="LOG DE ERRORES">
                            </dx:ASPxLabel>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <table>
        <table width="100%">
            <tr>
                <td style="width: 1180px">
                    <dx:ASPxListBox ID="ASPxListBoxWarnings" runat="server" Height="200px"
                        Width="100%">
                    </dx:ASPxListBox>
                </td>
            </tr>
        </table>
        <table>
            <tr>
                <td>
                    <div style="overflow: auto; width: 1200px; height: 100%">
                        <dxwgv:ASPxGridView runat="server" ID="ASPxGridViewFacturas" KeyFieldName="Nº Factura" EnableRowsCache="false" Width="100%" Visible="false">
                        </dxwgv:ASPxGridView>
                    </div>

                </td>
            </tr>
            <tr>
                <td align="center">
                    <table width="100%">
                        <tr>
                            <td width="40%">&nbsp;</td>
                            <td width="40px">
                                <asp:Button runat="server" ID="ASPxButtonProcesar" Text="Procesar" OnClick="ASPxButtonProcesar_Click" CssClass="btn btn-mfast-3" Height="41px" Width="139px" />
                            </td>
                        </tr>
                        <tr>
                            <td>&nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp &nbsp
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>       
</asp:Content>


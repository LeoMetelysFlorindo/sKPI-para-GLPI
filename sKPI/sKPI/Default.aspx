<%@ Page Title="Home Page" Language="vb" MasterPageFile="~/Site.Master" AutoEventWireup="false"
    CodeBehind="Default.aspx.vb" Inherits="sKPI._Default" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
    <script type="text/javascript">

     window.onload = maxWindow;
     window.status = false;


     function maxWindow() {
         window.moveTo(0, 0);


         if (document.all) {
             top.window.resizeTo(screen.availWidth, screen.availHeight);
         }

         else if (document.layers || document.getElementById) {
             if (top.window.outerHeight < screen.availHeight || top.window.outerWidth < screen.availWidth) {
                 top.window.outerHeight = screen.availHeight;
                 top.window.outerWidth = screen.availWidth;
             }
         }
     }

      
    
  </script>
   
</asp:Content>

<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    
   <table>
   <tr>
   <td><asp:Label ID="Label1" runat="server" style="font-size: medium; font-weight: 700; font-family: 'Trebuchet MS';" Text="Data Inicial"></asp:Label> 
   </td>
      <td><asp:Label ID="Label3" runat="server" Text= "         "></asp:Label> 
   </td>
   <td class="style1"><asp:Label ID="Label2" runat="server" style="font-size: medium; font-weight: 700; font-family: 'Trebuchet MS';" Text="Data Final"></asp:Label> 
   </td>
   <td>
   </td>
   </tr>
   <tr>
   <td><asp:TextBox ID="TxtDataini" runat="server"></asp:TextBox> </td>
    <td> </td>
    <td class="style1"><asp:TextBox ID="TxtDatafini" runat="server" AutoPostBack="True"></asp:TextBox> </td>
    <td><asp:Button ID="Button1" runat="server" Text="Ok" /> </td>
    <td><asp:Label ID="Label4" runat="server" Text= "         "></asp:Label>  </td>
    <td><asp:Label ID="Label5" runat="server" Text= "         "></asp:Label>  </td>
    <td><asp:Label ID="Label6" runat="server" Text= "         "></asp:Label>  </td>
    <td><asp:Button ID="BtnSalvar" runat="server" Text="Gerar Excel" Visible="False" /> </td>
   </tr>
   </table>
   <p>
       </table>
          <div class="blocoGrupoCampos" style="margin-left: 25px;">
                        <div class="blocoeditor">
                            <asp:Label ID="lblAviso" style="font-family: Arial Black;margin-left: 425px;font-size: 9pt;" runat="server" Text="Erro ou Invalido." ForeColor="Red"
                                Visible="False" Width="400px"></asp:Label>
                        </div>
                </div>
    <table>
                               <asp:GridView ID="gdItens" runat="server" CaptionAlign="Top"   DataKeyNames="ID"                               
                               Height="41px" style="margin-left: 3px; margin-top: 19px; color: #000000; font-size: medium; font-family: 'Trebuchet MS'" 
                               Width="1200px" Caption="KPI´s" SelectedIndex="1" 
                               PageSize="20"  AutoGenerateColumns="true" 
                               CellPadding="2" ForeColor="#333333" 
                                   EnableTheming="True" EnableSortingAndPagingCallbacks="True" 
                                   EnablePersistedSelection="True">
                               <HeaderStyle CssClass="th" />
                               <AlternatingRowStyle BackColor="White" />
                               <EditRowStyle BackColor="#2461BF" />
                               <FooterStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                               <HeaderStyle BackColor="#507CD1" Font-Bold="True" ForeColor="White" />
                               <PagerStyle Height="30px" BackColor="#2461BF" ForeColor="White" 
                                   HorizontalAlign="Center" />
                              
                               <RowStyle BackColor="#EFF3FB" />
                               <SelectedRowStyle BackColor="#D1DDF1" Font-Bold="false" ForeColor="#333333" />
                               <SortedAscendingCellStyle BackColor="#F5F7FB" />
                               <SortedAscendingHeaderStyle BackColor="#6D95E1" />
                               <SortedDescendingCellStyle BackColor="#E9EBEF" />
                               <SortedDescendingHeaderStyle BackColor="#4870BE" />
                              
                           </asp:GridView>
   </p>
    
   
</asp:Content>


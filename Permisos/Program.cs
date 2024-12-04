using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = SAPbouiCOM.Framework.Application;

namespace Permisos
{
    internal static class Program
    {


        public static SAPbobsCOM.Company oCompany;
        public static SAPbouiCOM.EditText _Bton;
        public static SAPbouiCOM.Button oButtonAdd;
        public static SAPbouiCOM.Item oItem;
        public static SAPbouiCOM.Form oForm;
        public static SAPbouiCOM.EditText CardCode;
        public static SAPbouiCOM.ComboBox U_Sucursal;
        public static SAPbouiCOM.Matrix oMatrix;
        public static SAPbouiCOM.Application sbo_application;
        public static string ItemCode = string.Empty;
        public static bool Band_Pressed = false;
        public static bool band = false;

        public static List<string> lista_Botones = new List<string>();
        public static string caracteres = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        public static int longitud = 8;
        public static string textoRandom = string.Empty;



        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                if (args.Length < 1)
                {
                    oApp = new Application();
                }
                else
                {
                    //If you want to use an add-on identifier for the development license, you can specify an add-on identifier string as the second parameter.
                    //oApp = new Application(args[0], "XXXXX");
                    oApp = new Application(args[0]);
                }
                sbo_application = Application.SBO_Application;
                oCompany = (SAPbobsCOM.Company)sbo_application.Company.GetDICompany();

                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                sbo_application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                sbo_application.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);
                sbo_application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
                sbo_application.StatusBarEvent += new SAPbouiCOM._IApplicationEvents_StatusBarEventEventHandler(SBO_Application_StatusBarEvent); oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }



        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }



        static void SBO_Application_StatusBarEvent(string Text, SAPbouiCOM.BoStatusBarMessageType MessageType)
        {

            //sbo_application.MessageBox(@"Status bar event with message: """ + Text + @""" has been sent", 1, "Ok", "", "");
            return;
        }



        static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {

                //Antes
                #region pValBeforeAction item event 

                if (pVal.BeforeAction)
                {

                    #region OFERTA DE VENTAS

                    if (pVal.FormTypeEx == "149") //Forma Oferta de ventas
                    {

                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                        {

                            foreach (var item in lista_Botones)
                            {

                                if (pVal.ItemUID == item)// Verificamos si el ítem presionado es nuestro botón
                                {

                                    try
                                    {

                                        //Forma
                                        oForm = Application.SBO_Application.Forms.ActiveForm;

                                        //CardCode
                                        CardCode = (SAPbouiCOM.EditText)oForm.Items.Item("4").Specific;

                                        //Matrix
                                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;

                                        if (CardCode.Value == "")
                                        {
                                            sbo_application.StatusBar.SetText("Favor de agregar un socio de negocio", SAPbouiCOM.BoMessageTime.bmt_Short,
                                               SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                            return;
                                        }

                                        //U_Sucursal
                                        U_Sucursal = (SAPbouiCOM.ComboBox)oForm.Items.Item("U_Sucursal").Specific;

                                        string Almacen_Cliente = GetAlmacenCliente(CardCode.Value);

                                        //Validacion almacen cliente
                                        if (U_Sucursal.Value == Almacen_Cliente)
                                        {

                                            if (band == false)
                                            {

                                                if (oMatrix.RowCount == 1)
                                                {
                                                    Application.SBO_Application.SetStatusBarMessage("Favor de ingresar al menos una partida", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                                    return;
                                                }

                                                //Band_Pressed  AHORA ES TRUE
                                                Band_Pressed = true;
                                                band = true;

                                                SAPbouiCOM.Column Col_ItemCode = oMatrix.Columns.Item("1");
                                                //string Col_CodeBarstitle = Col_CodeBars.Title;

                                                SAPbouiCOM.Column Col_U_Disponible = oMatrix.Columns.Item("U_Disponible");

                                                switch (U_Sucursal.Value)
                                                {
                                                    case "01":
                                                        Col_U_Disponible = oMatrix.Columns.Item("U_DisponibleGDL");
                                                        //string Col_U_Disponible_title = Col_U_Disponible.Title;
                                                        break;

                                                    case "02":
                                                        Col_U_Disponible = oMatrix.Columns.Item("U_DisponibleCDMX");
                                                        //string Col_U_Disponible_title = Col_U_Disponible.Title;
                                                        break;

                                                    case "03":
                                                        Col_U_Disponible = oMatrix.Columns.Item("U_DisponibleMTY");
                                                        //string Col_U_Disponible_title = Col_U_Disponible.Title;
                                                        break;

                                                    case "05":
                                                        Col_U_Disponible = oMatrix.Columns.Item("U_DisponibleSLT");
                                                        //string Col_U_Disponible_title = Col_U_Disponible.Title;
                                                        break;
                                                }

                                                SAPbouiCOM.Column Col_Tipo = oMatrix.Columns.Item("257");
                                                //string Tipo_title = Col_Tipo.Title;

                                                SAPbouiCOM.Column Col_Cantidad = oMatrix.Columns.Item("11");
                                                //string Col_Cantidad_title = Col_Cantidad.Title;

                                                SAPbouiCOM.Column Col_U_BackOrder = oMatrix.Columns.Item("U_BackOrder");
                                                //string Col_U_BackOrder_title = Col_U_BackOrder.Title;

                                                int matrixInsert = oMatrix.RowCount - 1;

                                                int matrixOrigin = oMatrix.RowCount - 1;

                                                for (int i = 1; i <= matrixOrigin; i++)
                                                {


                                                    ItemCode = ((dynamic)((SAPbouiCOM.ColumnClass)Col_ItemCode).Cells.Item(i).Specific).value;


                                                    string U_Disponible = ((dynamic)((SAPbouiCOM.ColumnClass)Col_U_Disponible).Cells.Item(i).Specific).value;
                                                    if (U_Disponible.Contains(","))
                                                    {
                                                        U_Disponible = U_Disponible.Replace(",", "");
                                                    }
                                                    int disponible = Convert.ToInt32(U_Disponible.Replace(".00", ""));


                                                    string tipo = ((dynamic)((SAPbouiCOM.ColumnClass)Col_Tipo).Cells.Item(i).Specific).value;


                                                    string CantidadString = ((dynamic)((SAPbouiCOM.ColumnClass)Col_Cantidad).Cells.Item(i).Specific).value;
                                                    int Cantidad = Convert.ToInt32(CantidadString.Replace(".000000", ""));


                                                    if (Cantidad > disponible)
                                                    {
                                                        if (disponible < 1)
                                                        {
                                                            //Cambiamos nueva linea insertada a alternativa
                                                            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)Col_Tipo.Cells.Item(i).Specific;
                                                            oCombo.Select("A", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                                        }
                                                        else
                                                        {
                                                            int diferencia = Cantidad - disponible;

                                                            //cambiamos la cantidad a disponible
                                                            SAPbouiCOM.EditText CantidadEnLinea = (SAPbouiCOM.EditText)Col_Cantidad.Cells.Item(i).Specific;
                                                            CantidadEnLinea.Value = disponible.ToString();


                                                            //Insertamos una linea nueva con la diferencia
                                                            matrixInsert++;
                                                            SAPbouiCOM.EditText NewItemCode = (SAPbouiCOM.EditText)Col_ItemCode.Cells.Item(matrixInsert).Specific;
                                                            NewItemCode.Value = ItemCode;

                                                            //Reutilizamos variable CantidadEnLinea para modificar cantidad de nueva linea
                                                            CantidadEnLinea = (SAPbouiCOM.EditText)Col_Cantidad.Cells.Item(matrixInsert).Specific;
                                                            CantidadEnLinea.Value = diferencia.ToString();

                                                            //Cambiamos nueva linea insertada a backorder
                                                            SAPbouiCOM.ComboBox oComboRef = (SAPbouiCOM.ComboBox)Col_U_BackOrder.Cells.Item(matrixInsert).Specific;
                                                            oComboRef.Select("01", SAPbouiCOM.BoSearchKey.psk_ByValue);

                                                            //Cambiamos nueva linea insertada a alternativa
                                                            oComboRef = (SAPbouiCOM.ComboBox)Col_Tipo.Cells.Item(matrixInsert).Specific;
                                                            oComboRef.Select("A", SAPbouiCOM.BoSearchKey.psk_ByValue);

                                                        }
                                                    }

                                                    #region COdigo anterior comentado
                                                    //string BackOrder = ((dynamic)((SAPbouiCOM.ColumnClass)Col_U_BackOrder).Cells.Item(i).Specific).value;


                                                    //if (Cantidad <= disponible)
                                                    //{
                                                    //    //Si backorder no es 03 cambiamos a 03 y alternativo a regular
                                                    //    if (BackOrder != "03")
                                                    //    {
                                                    //        //cambiamos cambiamos backorder a 03
                                                    //        SAPbouiCOM.ComboBox oComboRef = (SAPbouiCOM.ComboBox)Col_U_BackOrder.Cells.Item(i).Specific;
                                                    //        oComboRef.Select("03", SAPbouiCOM.BoSearchKey.psk_ByValue);

                                                    //        if (tipo != "")
                                                    //        {
                                                    //            //cambiamos tipo a regular
                                                    //            oComboRef = (SAPbouiCOM.ComboBox)Col_Tipo.Cells.Item(i).Specific;
                                                    //            oComboRef.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                                    //        }

                                                    //    }
                                                    //    else
                                                    //    {
                                                    //        if (tipo != "")
                                                    //        {
                                                    //            //cambiamos tipo a regular
                                                    //            SAPbouiCOM.ComboBox oComboRef = (SAPbouiCOM.ComboBox)Col_Tipo.Cells.Item(i).Specific;
                                                    //            oComboRef.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                                    //        }
                                                    //    }
                                                    //}
                                                    //else
                                                    //{

                                                    //    if (BackOrder == "01")
                                                    //    {
                                                    //        if (tipo != "A")
                                                    //        {
                                                    //            //cambiamos tipo a alternativo
                                                    //            SAPbouiCOM.ComboBox oComboRef = (SAPbouiCOM.ComboBox)Col_Tipo.Cells.Item(i).Specific;
                                                    //            oComboRef.Select("A", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                                    //        }
                                                    //    }
                                                    //    else
                                                    //    {

                                                    //        //cambiamos cambiamos backorder a 01
                                                    //        SAPbouiCOM.ComboBox oComboRef = (SAPbouiCOM.ComboBox)Col_U_BackOrder.Cells.Item(i).Specific;
                                                    //        oComboRef.Select("01", SAPbouiCOM.BoSearchKey.psk_ByValue);

                                                    //        if (tipo != "A")
                                                    //        {
                                                    //            //cambiamos tipo a alternativo
                                                    //            oComboRef = (SAPbouiCOM.ComboBox)Col_Tipo.Cells.Item(i).Specific;
                                                    //            oComboRef.Select("A", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                                    //        }

                                                    //    }
                                                    //}
                                                    #endregion

                                                }

                                                #region codigo util comentado
                                                //Quitar lo desabilitado de boton agregar
                                                //oForm.Items.Item("1").Enabled = true;
                                                //oForm.Items.Item("2349990001").Enabled = true;
                                                #endregion

                                                sbo_application.StatusBar.SetText("Análisis  Back Order completado", SAPbouiCOM.BoMessageTime.bmt_Short,
                                                SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                            }
                                            else
                                            {
                                                sbo_application.StatusBar.SetText("Análisis  Back Order previamente completado", SAPbouiCOM.BoMessageTime.bmt_Short,
                                                SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                            }
                                        }

                                        return;

                                    }
                                    catch (Exception ex)
                                    {
                                        Application.SBO_Application.SetStatusBarMessage("Error " + ex.Message + ". " + ItemCode, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                        return;
                                    }
                                    //break;
                                    //return;
                                }

                            }

                            #region comentado 
                            //if (pVal.ItemUID == "4")
                            //{
                            //    //Desabilitado de boton agregar
                            //    //oForm.Items.Item("1").Enabled = false;
                            //    //oForm.Items.Item("2349990001").Enabled = false;
                            //}
                            #endregion

                            if (pVal.ItemUID == "1")
                            {
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                {
                                    if (!Band_Pressed)
                                    {
                                        sbo_application.StatusBar.SetText("Es necesario hacer clic en el botón 'Procesar BackOrder'.", SAPbouiCOM.BoMessageTime.bmt_Short,
                                        SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        BubbleEvent = false;
                                        //return;
                                    }
                                    else
                                    {
                                        Band_Pressed = false;
                                    }
                                }
                            }



                        }

                        #region leer  y modificar cuando hay cambion en cantidad
                        //if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE) // Evento cuando se valida el valor
                        //{
                        //    if (pVal.ItemUID == "38") // Verifica el UID del ítem (en este caso, el grid o el campo de interés)
                        //    {
                        //        if (pVal.ColUID == "11") // Verifica la columna en la que se realizó el cambio
                        //        {

                        //            try
                        //            {

                        //                //Forma
                        //                oForm = Application.SBO_Application.Forms.ActiveForm;

                        //                //Matrix
                        //                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;

                        //                SAPbouiCOM.Column Col_U_Disponible = oMatrix.Columns.Item("U_Disponible");
                        //                SAPbouiCOM.Column Col_Cantidad = oMatrix.Columns.Item("11");
                        //                SAPbouiCOM.Column Col_U_BackOrder = oMatrix.Columns.Item("U_BackOrder");

                        //                for (int i = 1; i <= oMatrix.RowCount - 1; i++)
                        //                {

                        //                    if(pVal.Row == i)
                        //                    {

                        //                        //Campo U_Disponible
                        //                        string U_Disponible = ((dynamic)((SAPbouiCOM.ColumnClass)Col_U_Disponible).Cells.Item(i).Specific).value;
                        //                        if (U_Disponible.Contains(","))
                        //                        {
                        //                            U_Disponible = U_Disponible.Replace(",", "");
                        //                        }
                        //                        int disponible = Convert.ToInt32(U_Disponible.Replace(".00", ""));

                        //                        //Campo Cantidad
                        //                        string CantidadString = ((dynamic)((SAPbouiCOM.ColumnClass)Col_Cantidad).Cells.Item(i).Specific).value;
                        //                        int Cantidad = Convert.ToInt32(CantidadString.Replace(".000000", ""));

                        //                        //Campo BackOrder
                        //                        string BackOrder = ((dynamic)((SAPbouiCOM.ColumnClass)Col_U_BackOrder).Cells.Item(i).Specific).value;

                        //                    }
                        //                }
                        //                return;
                        //            }
                        //            catch (Exception ex)
                        //            {
                        //                Application.SBO_Application.SetStatusBarMessage("Error " + ex.Message + ". " + CodeBar, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                        //                return;
                        //            }

                        //        }
                        //    }
                        //}
                        #endregion

                    }

                    #endregion

                }

                #endregion

                //Despues
                #region  !pVal.BeforeAction item event

                else
                {
                    if (pVal.FormTypeEx == "149") //Forma Oferta de ventas
                    {

                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DRAW) //Evento cargando forma
                        {

                            Band_Pressed = false; //Band_Pressed  SE REINICIA A FALSE
                            band = false;

                            oForm = oForm = Application.SBO_Application.Forms.ActiveForm;

                            textoRandom = string.Empty;
                            textoRandom = GenerarTextoRandom(caracteres, longitud);

                            lista_Botones.Add(textoRandom);

                            ////oItem = oForm.Items.Add("btnPrueba", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            oItem = oForm.Items.Add(textoRandom, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            _Bton = (SAPbouiCOM.EditText)(oForm.Items.Item("4").Specific);

                            oItem.Top = 120;
                            oItem.Left = _Bton.Item.Left;
                            oItem.Width = 150;

                            oButtonAdd = (SAPbouiCOM.Button)oItem.Specific;
                            oButtonAdd.Caption = "Procesar BackOrder";
                            oButtonAdd = null;
                        }

                    }
                }

                #region comentado
                //else if (!pVal.BeforeAction)
                //{
                //    if (pVal.FormTypeEx == "149") //Forma Oferta de ventas
                //    {
                //        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                //        {
                //            if (pVal.ItemUID == "btnPrueba")
                //            {
                //                if (Band_Pressed)
                //                {
                //                    //Desabilitado de boton agregar
                //                    oForm = sbo_application.Forms.GetForm("-149", 1);
                //                    //oForm.Items.Item("2349990001").Enabled = false;
                //                    //cambiamos cambiamos backorder de cabecera
                //                    SAPbouiCOM.ComboBox oComboRef = (SAPbouiCOM.ComboBox)oForm.Items.Item("U_BackOrder").Specific;
                //                    oComboRef.Select("02", SAPbouiCOM.BoSearchKey.psk_ByValue);
                //                }

                //            }
                //        }
                //    }
                //}
                #endregion

                #endregion


            }
            catch (Exception ex)
            {
                sbo_application.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }

        }



        static void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {

            BubbleEvent = true;

        }



        static void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {

            BubbleEvent = true;

        }



        static string GenerarTextoRandom(string caracteres, int longitud)
        {
            Random random = new Random();
            char[] resultado = new char[longitud];

            for (int i = 0; i < longitud; i++)
            {
                int indice = random.Next(caracteres.Length);
                resultado[i] = caracteres[indice];
            }

            return new string(resultado);
        }



        public static string GetAlmacenCliente(string CardCode )
        {
            string Almacen = "";
            try
            {
                SAPbobsCOM.Recordset oRS;
                StringBuilder query = new StringBuilder();
                string valor = "";

                oRS = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                query.Append(
                    $@"SELECT
	                    T0.""U_Almacen""

                   FROM
	                    OCRD T0

                   WHERE
	                    T0.""CardCode"" = '{CardCode}'"
                           );

                oRS.DoQuery(query.ToString());

                if (oRS.RecordCount > 0)
                {
                    while (oRS.BoF)
                    {
                        valor = oRS.Fields.Item("U_Almacen").Value.ToString();
                        if (valor != "")
                        {
                            Almacen = valor;
                        }
                        else
                        {
                            //Nada
                        }
                        oRS.MoveNext();
                    }
                }
                else
                {
                    //Nada
                }
            }
            catch (Exception ex)
            {
                //Nada
                string err = string.Empty;
                err = ex.Message;

            }
            return Almacen;
        }



    }
}

using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = SAPbouiCOM.Framework.Application;

namespace Permisos
{
    internal static class Program
    {



        public static SAPbouiCOM.EditText _Bton;
        public static SAPbouiCOM.Button oButtonAdd;
        //public static SAPbouiCOM.EditText _Bton_;
        //public static SAPbouiCOM.EditText _Bton_Agregar;
        public static SAPbouiCOM.Item oItem;
        public static SAPbouiCOM.Form oForm;
        public static SAPbouiCOM.EditText oEditTextDocNum;
        public static SAPbouiCOM.Matrix oMatrix;
        public static SAPbouiCOM.Application sbo_application;
        public static string CodeBar = string.Empty;
        public static bool Band_Pressed = false;


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
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DRAW) //Evento cargando forma
                        {

                            Band_Pressed = false; //Band_Pressed  SE REINICIA A FALSE

                            oForm = sbo_application.Forms.GetForm("149", 1);

                            oItem = oForm.Items.Add("btnPrueba", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            _Bton = (SAPbouiCOM.EditText)(oForm.Items.Item("4").Specific);

                            oItem.Top = 120;
                            oItem.Left = _Bton.Item.Left;
                            oItem.Width = 150;

                            oButtonAdd = (SAPbouiCOM.Button)oItem.Specific;
                            oButtonAdd.Caption = "Procesar BackOrder";
                            oButtonAdd = null;

                        }

                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                        {

                            if (pVal.ItemUID == "btnPrueba")// Verificamos si el ítem presionado es nuestro botón
                            {
                                try
                                {

                                    //Forma
                                    oForm = Application.SBO_Application.Forms.ActiveForm;

                                    //Docnum
                                    oEditTextDocNum = (SAPbouiCOM.EditText)oForm.Items.Item("4").Specific;

                                    //Matrix
                                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;

                                    if (oEditTextDocNum.Value == "")
                                    {
                                        sbo_application.StatusBar.SetText("Favor de agregar un socio de negocio", SAPbouiCOM.BoMessageTime.bmt_Short,
                                           SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                        return;
                                    }

                                    if (oMatrix.RowCount == 1)
                                    {
                                        Application.SBO_Application.SetStatusBarMessage("Favor de ingresar al menos una partida", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                        return;
                                    }

                                    //Band_Pressed  AHORA ES TRUE
                                    Band_Pressed = true;

                                    SAPbouiCOM.Column Col_CodeBars = oMatrix.Columns.Item("4");
                                    string Col_CodeBarstitle = Col_CodeBars.Title;

                                    SAPbouiCOM.Column Col_U_Disponible = oMatrix.Columns.Item("U_Disponible");
                                    string Col_U_Disponible_title = Col_U_Disponible.Title;

                                    SAPbouiCOM.Column Col_Tipo = oMatrix.Columns.Item("257");
                                    string Tipo_title = Col_Tipo.Title;

                                    SAPbouiCOM.Column Col_Cantidad = oMatrix.Columns.Item("11");
                                    string Col_Cantidad_title = Col_Cantidad.Title;

                                    SAPbouiCOM.Column Col_U_BackOrder = oMatrix.Columns.Item("U_BackOrder");
                                    string Col_U_BackOrder_title = Col_U_BackOrder.Title;

                                    for (int i = 1; i <= oMatrix.RowCount - 1; i++)
                                    {
                                        CodeBar = ((dynamic)((SAPbouiCOM.ColumnClass)Col_CodeBars).Cells.Item(i).Specific).value;

                                        string U_Disponible = ((dynamic)((SAPbouiCOM.ColumnClass)Col_U_Disponible).Cells.Item(i).Specific).value;
                                        if (U_Disponible.Contains(","))
                                        {
                                            U_Disponible = U_Disponible.Replace(",", "");
                                        }
                                        int disponible = Convert.ToInt32(U_Disponible.Replace(".00", ""));

                                        string tipo = ((dynamic)((SAPbouiCOM.ColumnClass)Col_Tipo).Cells.Item(i).Specific).value;

                                        string CantidadString = ((dynamic)((SAPbouiCOM.ColumnClass)Col_Cantidad).Cells.Item(i).Specific).value;
                                        int Cantidad = Convert.ToInt32(CantidadString.Replace(".000000", ""));

                                        string BackOrder = ((dynamic)((SAPbouiCOM.ColumnClass)Col_U_BackOrder).Cells.Item(i).Specific).value;

                                        if (Cantidad <= disponible)
                                        {
                                            //Si backorder no es 03 cambiamos a 03 y alternativo a regular
                                            if (BackOrder != "03")
                                            {
                                                //cambiamos cambiamos backorder a 03
                                                SAPbouiCOM.ComboBox oComboRef = (SAPbouiCOM.ComboBox)Col_U_BackOrder.Cells.Item(i).Specific;
                                                oComboRef.Select("03", SAPbouiCOM.BoSearchKey.psk_ByValue);

                                                if (tipo != "")
                                                {
                                                    //cambiamos tipo a regular
                                                    oComboRef = (SAPbouiCOM.ComboBox)Col_Tipo.Cells.Item(i).Specific;
                                                    oComboRef.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                                }

                                            }
                                            else
                                            {
                                                if (tipo != "")
                                                {
                                                    //cambiamos tipo a regular
                                                    SAPbouiCOM.ComboBox oComboRef = (SAPbouiCOM.ComboBox)Col_Tipo.Cells.Item(i).Specific;
                                                    oComboRef.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                                }
                                            }
                                        }
                                        else
                                        {

                                            if (BackOrder == "01")
                                            {
                                                if (tipo != "A")
                                                {
                                                    //cambiamos tipo a alternativo
                                                    SAPbouiCOM.ComboBox oComboRef = (SAPbouiCOM.ComboBox)Col_Tipo.Cells.Item(i).Specific;
                                                    oComboRef.Select("A", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                                }
                                            }
                                            else
                                            {

                                                //cambiamos cambiamos backorder a 01
                                                SAPbouiCOM.ComboBox oComboRef = (SAPbouiCOM.ComboBox)Col_U_BackOrder.Cells.Item(i).Specific;
                                                oComboRef.Select("01", SAPbouiCOM.BoSearchKey.psk_ByValue);

                                                if (tipo != "A")
                                                {
                                                    //cambiamos tipo a alternativo
                                                    oComboRef = (SAPbouiCOM.ComboBox)Col_Tipo.Cells.Item(i).Specific;
                                                    oComboRef.Select("A", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                                }

                                            }
                                        }


                                    }
                                    //Quitar lo desabilitado de boton agregar
                                    //oForm.Items.Item("1").Enabled = true;
                                    //oForm.Items.Item("2349990001").Enabled = true;
                                    sbo_application.StatusBar.SetText("Análisis  Back Order completado", SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                    return;
                                }
                                catch (Exception ex)
                                {
                                    Application.SBO_Application.SetStatusBarMessage("Error " + ex.Message + ". " + CodeBar, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                    return;
                                }
                            }

                            //if (pVal.ItemUID == "4")
                            //{
                            //    //Desabilitado de boton agregar
                            //    //oForm.Items.Item("1").Enabled = false;
                            //    //oForm.Items.Item("2349990001").Enabled = false;
                            //}

                            if(pVal.ItemUID == "1")
                            {
                                if(oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
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

                    }

                    #endregion

                }

                #endregion

                //Despues
                #region  !pVal.BeforeAction item event

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



    }
}

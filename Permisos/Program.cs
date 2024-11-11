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
        public static SAPbouiCOM.Item oItem;
        public static SAPbouiCOM.Button oButtonAdd;
        public static SAPbouiCOM.Form oForm;
        public static SAPbouiCOM.EditText oEditTextDocNum;
        public static SAPbouiCOM.Matrix oMatrix;
        public static SAPbouiCOM.Application sbo_application;



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
            
        }



        static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {

                //Antes
                # region pValBeforeAction item event

                if (pVal.BeforeAction)
                {

                    #region OFERTA DE VENTAS

                    if (pVal.FormTypeEx == "149")
                    {
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DRAW)
                        {
                            oForm = sbo_application.Forms.GetForm("149", 1);

                            oItem = oForm.Items.Add("btnPrueba", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            _Bton = (SAPbouiCOM.EditText)(oForm.Items.Item("4").Specific);

                            oItem.Top = 120;
                            oItem.Left = _Bton.Item.Left;
                            oItem.Width = 70;

                            oButtonAdd = (SAPbouiCOM.Button)oItem.Specific;
                            oButtonAdd.Caption = "BackOrder";
                            oButtonAdd = null;
                        }

                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                        {
                            // Verificamos si el ítem presionado es nuestro botón
                            if (pVal.ItemUID == "btnPrueba")
                            {
                                sbo_application.StatusBar.SetText("boton presionado", SAPbouiCOM.BoMessageTime.bmt_Short,
                           SAPbouiCOM.BoStatusBarMessageType.smt_Success);

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

                                        string U_Disponible = ((dynamic)((SAPbouiCOM.ColumnClass)Col_U_Disponible).Cells.Item(i).Specific).value;
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

                                                if(tipo != "")
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
                                            //BackOrder 01
                                            if (BackOrder == "01")
                                            {
                                                if(tipo != "A")
                                                {
                                                    //cambiamos tipo a alternativo
                                                    SAPbouiCOM.ComboBox oComboRef = (SAPbouiCOM.ComboBox)Col_Tipo.Cells.Item(i).Specific;
                                                    oComboRef.Select("A", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                                }
                                            }
                                            else
                                            {
                                                //Cambiamos back order a 01 y tipo a alternativo
                                            }
                                        }


                                    }
                                    return;
                                }
                                catch (Exception ex)
                                {
                                    Application.SBO_Application.SetStatusBarMessage("Error " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, false);
                                    return;
                                }
                            }


                        }

                    }

                    #endregion

                }

                #endregion

                //Despues
                #region  !pVal.BeforeAction item event

                else if (!pVal.BeforeAction)
                {

                }

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

using System;
using System.Runtime.InteropServices;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;

namespace SW_teste
{
    [ComVisible(true)]
    [Guid("12345678-1734-1734-1234-1834567890AB")]
    public class Addin : ISwAddin
    {
        private SldWorks swApp;
        private int addinID;

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            swApp = (SldWorks)ThisSW;
            addinID = cookie;

            // Cria uma PropertyManagerPage simples
            CreatePMPage();

            swApp.SendMsgToUser2("Suplemento SW_teste carregado com sucesso!", 
                (int)swMessageBoxIcon_e.swMbInformation, 
                (int)swMessageBoxBtn_e.swMbOk);

            return true;
        }

        public bool DisconnectFromSW()
        {
            swApp = null;
            return true;
        }

        private void CreatePMPage()
        {
            PropertyManagerPage2 pmPage = (PropertyManagerPage2)swApp.CreatePropertyManagerPage(
                "Exemplo Add-in", 
                (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_OkayButton, 
                null, 
                ref addinID);

            int controlId = 1;
            // Adiciona um botão (corrigido: todos os argumentos necessários)
            pmPage.AddControl(
                controlId, 
                (int)swPropertyManagerPageControlType_e.swControlType_Button, 
                "Clique aqui", 
                (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge, 
                0, 
                "Botão de exemplo"
            );
        }
    }
}
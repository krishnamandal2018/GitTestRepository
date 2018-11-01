//Login with Admin,operator,Customer and check the UI element
//Test Scenario: TS-16
function OperatorLoginScript()
{ 
//
      var Excel;
      Excel = Sys.OleObject("Excel.Application");
      Delay (3000); 
      // Wait until Excel starts
      Excel.Visible = true;
      Excel.Workbooks.Open("C:\\Users\\t89106\\Documents\\TestComplete 12 Projects\\OCRonWEB_UI\\Master.xlsx"); 
      for(var v=1;v<=3;v++)
      {
          var userId=VarToString(Excel.Cells.Item(v,2));
          var   pass=VarToString(Excel.Cells.Item(v,3));
          //TestedApps.https___st_rbpocloud_com.Run();
      Browsers.Item(btChrome).Run("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/");
      Aliases.browser.pageOcrtest1673629386ApNortheast.form.textboxTxtusername.SetText(userId);
      Aliases.browser.pageOcrtest1673629386ApNortheast.form.submitbuttonBtnlogin.ClickButton();
          //Waits until the browser loads the page and is ready to accept user input.
          Delay(3000);
          //Click on Fax-reception list
          Sys.Browser("chrome").Page("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/App_GeneralFunction/FS00004/FS00004002.aspx?sid=ow&svc=oow&lang=ja-JP").Panel("bd_inner").Panel("build_menu_page").Form("fm").Panel("phFormArea_headerArea").Panel(3).Panel("phFormArea_phContentsTabArea_contentsMenu_contentsMenu").Panel(0).Link("phFormArea_phContentsTabArea_contentsMenu_submenuMenu_Item1003").Click();
          // Aliases.browser.Page("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/App_GeneralFunction/FS00001/FS00001001.aspx?sid=ow&svc=oow&lang=ja-JP").Panel("bd_inner").Panel("build_menu_page").Form("fm").Panel("phFormArea_mainArea").Panel("storageSelectArea").Panel("FieldsetTab").TextNode(1).Click();
          Delay(4000)
                          
          if(v==2)
          {
            if(aqObject.CheckProperty(Sys.Browser("chrome").Page("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/App_GeneralFunction/FS00001/FS00001001.aspx?sid=ow&svc=oow&lang=ja-JP").Panel("bd_inner").Panel("build_menu_page").Form("fm").Panel("phFormArea_headerArea").Panel(0).Panel(1).Link(0).TextNode("phFormArea_phHeaderMenuArea_headerMenu_ltrUserId"), "contentText", cmpEqual,userId))
            Log.Message("Display User Name on the rigth top corner on the UI:-OK");
            else
            Log.Error("Display User Name on the rigth top corner on the UI:-NG")
          }
          if(v==1)
          {
            if(aqObject.CheckProperty(Sys.Browser("chrome").Page("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/App_GeneralFunction/FS00001/FS00001001.aspx?sid=ow&svc=oow&lang=ja-JP").Panel("bd_inner").Panel("build_menu_page").Form("fm").Panel("phFormArea_headerArea").Panel(0).Panel(1).Link(0).TextNode("phFormArea_phHeaderMenuArea_headerMenu_ltrUserId"), "contentText", cmpEqual,userId))
            Log.Message("Display User Name on the rigth top corner on the UI:-OK");
            else
            Log.Error("Display User Name on the rigth top corner on the UI:-NG")
          }
          //for customer login
          if(v==3)
          {
            if(aqObject.CheckProperty(Sys.Browser("chrome").Page("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/App_GeneralFunction/FS00001/FS00001001.aspx?sid=ow&svc=oow&lang=ja-JP").Panel("bd_inner").Panel("build_menu_page").Form("fm").Panel("phFormArea_headerArea").Panel(0).Panel(1).Link(0).TextNode("phFormArea_phHeaderMenuArea_headerMenu_ltrUserId"), "contentText", cmpEqual,userId))
            Log.Message("Display User Name on the rigth top corner on the UI:-OK");
            else
            Log.Error("Display User Name on the rigth top corner on the UI:-NG") 
            if(Sys.Browser("chrome").Page("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/App_GeneralFunction/FS00001/FS00001001.aspx?sid=ow&svc=oow&lang=ja-JP").Panel("bd_inner").Panel("build_menu_page").Form("fm").Panel("phFormArea_headerArea").Panel("phFormArea_headerAreaLogo").Panel(0).Image("kmui_header_kmlogo_png_v_5247917528372935018").Exists)
            Log.Message("Display Konica Minolta logo in UI:-OK");
            else
            Log.Error("Display Konica Minolta logo in UI-NG");
            //Compares the panel12 Stores item with the image of the Aliases.browser.pageHttpsWebRobobpoComAppGeneral.formFm.panelPhformareaHeaderarea.panel.panel object.
            if(Sys.Browser("chrome").Page("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/App_GeneralFunction/FS00001/FS00001001.aspx?sid=ow&svc=oow&lang=ja-JP").Panel("bd_inner").Panel("build_menu_page").Form("fm").Panel("phFormArea_headerArea").Panel(0).Panel(0).Image("phFormArea_phHeaderMenuArea_headerMenu_imgOcrOnWeb").Exists)
            Log.Message("Display robotic bpo for smart work image in UI:-OK");
            else
            Log.Error("Display robotic bpo for smart work image in UI:-NG")

            if(Sys.Browser("chrome").Page("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/App_GeneralFunction/FS00001/FS00001001.aspx?sid=ow&svc=oow&lang=ja-JP").Panel("bd_inner").Panel("build_menu_page").Form("fm").Panel("phFormArea_mainArea").Panel("storageSelectArea").Panel("ConfirmedFolderArea").Panel("NoNeedConfirmBox").Link(0).contentText=="取消済")
            Log.Message("Display 取消済 folder in Customer folder list:-OK");
            else
            Log.Error("Display 取消済 work Customer folder  list tab  :-NG")

            if(Sys.Browser("chrome").Page("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/App_GeneralFunction/FS00001/FS00001001.aspx?sid=ow&svc=oow&lang=ja-JP").Panel("bd_inner").Panel("build_menu_page").Form("fm").Panel("phFormArea_mainArea").Panel("storageSelectArea").Panel("UnconfirmedFolderArea").Panel("UnconfirmedFolderTree").Link(0).contentText=="すべてのファイル")
            Log.Checkpoint("すべてのファイル folder displayed :-OK");
            else
            Log.Error("すべてのファイル  folder displayed : NG");
            if(Sys.Browser("chrome").Page("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/App_GeneralFunction/FS00001/FS00001001.aspx?sid=ow&svc=oow&lang=ja-JP").Panel("bd_inner").Panel("build_menu_page").Form("fm").Panel("phFormArea_mainArea").Panel("storageSelectArea").Panel("UnconfirmedFolderArea").Panel("attentionUnconfirmBox").Link(0).contentText=="要注意")
            Log.Checkpoint("要注意 folder displayed :-OK");
            else
            Log.Error("要注意  folder displayed : NG");
            break;
          }  
          //Compares the imageKmuiHeaderKmlogoPngV52479173 Stores item with the image of the Aliases.browser.pageHttpsWebRobobpoComAppGeneral.formFm.panelPhformareaHeaderarea.panel.imageKmuiHeaderKmlogoPngV5247917 object.
          if(Sys.Browser("chrome").Page("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/App_GeneralFunction/FS00001/FS00001001.aspx?sid=ow&svc=oow&lang=ja-JP").Panel("bd_inner").Panel("build_menu_page").Form("fm").Panel("phFormArea_headerArea").Panel("phFormArea_headerAreaLogo").Panel(0).Image("kmui_header_kmlogo_png_v_5247917528372935018").Exists)
          Log.Message("Display Konica Minolta logo in UI:-OK");
          else
          Log.Error("Display Konica Minolta logo in UI-NG");
          //Compares the panel12 Stores item with the image of the Aliases.browser.pageHttpsWebRobobpoComAppGeneral.formFm.panelPhformareaHeaderarea.panel.panel object.
          if(Sys.Browser("chrome").Page("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/App_GeneralFunction/FS00001/FS00001001.aspx?sid=ow&svc=oow&lang=ja-JP").Panel("bd_inner").Panel("build_menu_page").Form("fm").Panel("phFormArea_headerArea").Panel(0).Panel(0).Image("phFormArea_phHeaderMenuArea_headerMenu_imgOcrOnWeb").Exists)
          Log.Message("Display robotic bpo for smart work image in UI:-OK");
          else
          Log.Error("Display robotic bpo for smart work image in UI:-NG")
          ///
          if(Sys.Browser("chrome").Page("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/App_GeneralFunction/FS00001/FS00001001.aspx?sid=ow&svc=oow&lang=ja-JP").Panel("bd_inner").Panel("build_menu_page").Form("fm").Panel("phFormArea_mainArea").Panel("storageSelectArea").Panel("UnconfirmedFolderArea").Panel("UnconfirmedFolderTree").Link(0).contentText=="すべてのファイル")
          Log.Checkpoint("すべてのファイル folder displayed in work list tab UI:-OK");
          else
          Log.Error("すべてのファイル  folder displayed work list tab UI: NG");
          if(Sys.Browser("chrome").Page("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/App_GeneralFunction/FS00001/FS00001001.aspx?sid=ow&svc=oow&lang=ja-JP").Panel("bd_inner").Panel("build_menu_page").Form("fm").Panel("phFormArea_mainArea").Panel("storageSelectArea").Panel("UnconfirmedFolderArea").Panel("attentionUnconfirmBox").Link(0).contentText=="要注意")
          Log.Checkpoint("要注意 folder displayed in work list tab UI:-OK");
          else
          Log.Error("要注意  folder displayed work list tab UI: NG");
          if(Sys.Browser("chrome").Page("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/App_GeneralFunction/FS00001/FS00001001.aspx?sid=ow&svc=oow&lang=ja-JP").Panel("bd_inner").Panel("build_menu_page").Form("fm").Panel("phFormArea_mainArea").Panel("storageSelectArea").Panel("UnconfirmedFolderArea").Panel("NotSubmittedBox").Link(0).contentText=="未提出")
          Log.Checkpoint("未提出 folder displayed in work list tab UI:-OK");
          else
          Log.Error("未提出  folder displayed work list tab UI: NG");
          if(Sys.Browser("chrome").Page("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/App_GeneralFunction/FS00001/FS00001001.aspx?sid=ow&svc=oow&lang=ja-JP").Panel("bd_inner").Panel("build_menu_page").Form("fm").Panel("phFormArea_mainArea").Panel("storageSelectArea").Panel("UnconfirmedFolderArea").Panel("UnconfirmedTrashBox").Link(0).contentText=="ゴミ箱")
          Log.Checkpoint("ゴミ箱 folder displayed in work list tab UI:-OK");
          else
          Log.Error("ゴミ箱  folder displayed work list tab UI: NG");
          if(Sys.Browser("chrome").Page("https://ocrtest-1673629386.ap-northeast-1.elb.amazonaws.com/App_GeneralFunction/FS00001/FS00001001.aspx?sid=ow&svc=oow&lang=ja-JP").Panel("bd_inner").Panel("build_menu_page").Form("fm").Panel("phFormArea_mainArea").Panel("storageSelectArea").Panel("UnconfirmedFolderArea").Panel("UnregisteredBox").Link(0).contentText=="未登録")
          Log.Message("未登録 folder displayed in work list tab UI:-OK");
          else
          Log.Error("未登録  folder displayed work list tab UI: NG");



          var browser =Sys.Browser("chrome");
          browser.Close();
        
    
      }
}
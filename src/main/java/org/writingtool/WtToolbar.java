/* WritingTool, a LibreOffice Extension based on LanguageTool
 * Copyright (C) 2024 Fred Kruse (https://writingtool.org)
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301
 * USA
 */
package org.writingtool;

import java.util.Arrays;
import java.util.ResourceBundle;

import org.apache.commons.lang3.SystemUtils;
import org.languagetool.JLanguageTool;
import org.languagetool.Language;
import org.writingtool.config.WtConfiguration;
import org.writingtool.sidebar.WtSidebarPanel;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;

import com.sun.star.awt.Point;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XWindow;
import com.sun.star.beans.Property;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XIndexContainer;
import com.sun.star.drawing.framework.XToolBar;
import com.sun.star.frame.XController;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XLayoutManager;
import com.sun.star.frame.XModel;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lib.uno.helper.ComponentBase;
import com.sun.star.lib.uno.helper.WeakBase;
import com.sun.star.ui.ConfigurationEvent;
import com.sun.star.ui.DockingArea;
import com.sun.star.ui.ItemStyle;
import com.sun.star.ui.ItemType;
import com.sun.star.ui.UIElementType;
import com.sun.star.ui.XModuleUIConfigurationManagerSupplier;
import com.sun.star.ui.XToolPanel;
import com.sun.star.ui.XUIConfigurationManager;
import com.sun.star.ui.XUIElement;
import com.sun.star.ui.XUIElementFactory;
import com.sun.star.ui.XUIElementFactoryRegistration;
import com.sun.star.ui.XUIElementSettings;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/** 
 *  Class to add a dynamic WritingTool Toolbar
 *  since 26.1
 */
public class WtToolbar {
  
  private static final ResourceBundle MESSAGES = WtOfficeTools.getMessageBundle();
  private static final String WRITER_SERVICE = "com.sun.star.text.TextDocument";
  private static final String WT_NEW_TOOLBAR_URL = "private:resource/toolbar/WtToolbarFactory.toolbar";
  private static final String WT_NEW_AI_TOOLBAR_URL = "private:resource/toolbar/WtToolbarFactory?aiGeneralCommand";
  private static final String WT_NEW_TOOLBAR_NAME = "WritingToolNew";
  private XComponentContext xContext;
  private WtSingleDocument document;

  WtToolbar(XComponentContext xContext, WtSingleDocument document) {
    this.xContext = xContext;
    this.document = document;
    createToolbar();
  }
  
  private void createToolbar() {
    try {
      XUIConfigurationManager confMan = getUIConfigManagerDoc(xContext);
      if (confMan == null) {
        WtMessageHandler.printToLogFile("WtToolbar: makeToolbar: Cannot create configuration manager");
        return;
      }
      
      String newToolbarName = WT_NEW_TOOLBAR_NAME;
      String newToolbarURL = WT_NEW_TOOLBAR_URL;
      
/*      
      XUIElementFactoryRegistration uiFactoryRegistration = getUIElementFactoryRegistration(xContext);
      uiFactoryRegistration.registerFactory("toolbar", "WtToolbarFactory", "", WtToolbarFactory.class.getName());
//      printRegisteredUIElementFactories(xContext);
      XUIElementFactory wtUIFactory = uiFactoryRegistration.getFactory(newToolbarURL, "");
      XUIElement wtToolbar = wtUIFactory.createUIElement(newToolbarURL, new PropertyValue[0]);
      XUIElementSettings xWtBarSettings = UnoRuntime.queryInterface(XUIElementSettings.class, wtToolbar);
      XIndexAccess UISettings = xWtBarSettings.getSettings(true);
      XPropertySet propset = UnoRuntime.queryInterface(XPropertySet.class, UISettings);
      propset.setPropertyValue("Persistent", true);
      propset.setPropertyValue("ConfigurationSource", confMan);
      propset.setPropertyValue("UIName", WT_NEW_TOOLBAR_NAME);
      xWtBarSettings.updateSettings();
      for (int i = 0; i < UISettings.getCount(); i++) {
        PropertyValue[] propVal = (PropertyValue[]) UISettings.getByIndex(i);
        for (int k = 0; k < propVal.length; k++) {
          WtMessageHandler.printToLogFile("WtToolbar: makeToolbar: " + i + ".: Property: Name: " + propVal[k].Name + ", Handle: " + propVal[k].Handle 
              + ", Value: " + propVal[k].Value + ", State: " + propVal[k].State);
        }
      }
*/      
      WtMessageHandler.printToLogFile("WtToolbar: makeToolbar: is running");
      XLayoutManager layoutManager = getLayoutManager();
//      layoutManager.destroyElement(toolbarName);

      boolean hasConfSettings = confMan.hasSettings(newToolbarURL);
      WtMessageHandler.printToLogFile("WtToolbar: makeToolbar: hasSettings: " + 
          (hasConfSettings ? "true" : "false") + ": " + newToolbarURL);
      
      XIndexAccess settings = confMan.createSettings();
      XPropertySet props = UnoRuntime.queryInterface(XPropertySet.class, settings);
      for (Property property : props.getPropertySetInfo().getProperties()) {
        WtMessageHandler.printToLogFile("Property: Name: " + property.Name + ", type: " + property.Type 
            + ", Value: " + props.getPropertyValue(property.Name));
      }
      XIndexContainer elementsContainer = UnoRuntime.queryInterface(XIndexContainer.class, settings);
      addToolbarButtons(elementsContainer, confMan);
      if (hasConfSettings) {
        confMan.replaceSettings(newToolbarURL, elementsContainer);
      } else {
        confMan.insertSettings(newToolbarURL, elementsContainer);
      }
      if (layoutManager.isElementVisible(newToolbarName)) {
        layoutManager.hideElement(newToolbarName);
        layoutManager.showElement(newToolbarName);
      }
      if (layoutManager.getElement(newToolbarURL) == null) {
        layoutManager.createElement(newToolbarURL);
        XUIElement toolbar = layoutManager.getElement(newToolbarURL);
        XUIElementSettings oLtBarSettings = UnoRuntime.queryInterface(XUIElementSettings.class, toolbar);
        XIndexAccess oLtBarAccess = oLtBarSettings.getSettings(true);
//        XPropertySet propset = UnoRuntime.queryInterface(XPropertySet.class, oLtBarAccess);
//        propset.setPropertyValue("UIName", WT_NEW_TOOLBAR_NAME);
        oLtBarSettings.updateSettings();
        for (int i = 0; i < oLtBarAccess.getCount(); i++) {
          WtMessageHandler.printToLogFile("");
          PropertyValue[] propVal = (PropertyValue[]) oLtBarAccess.getByIndex(i);
          for (int k = 0; k < propVal.length; k++) {
            WtMessageHandler.printToLogFile(i + ".: Property: Name: " + propVal[k].Name + ", Handle: " + propVal[k].Handle 
                + ", Value: " + propVal[k].Value + ", State: " + propVal[k].State);
//            WtMessageHandler.printToLogFile("Access (" + i + "): " + oLtBarAccess.getByIndex(i));
          }
        }
//        XToolPanel wtToolBar = UnoRuntime.queryInterface(XToolPanel.class, toolbar.getRealInterface());
        WtMessageHandler.printToLogFile("toolbar is " + (toolbar == null ? "NULL" : "NOT null"));
        for (XUIElement element : layoutManager.getElements()) {
          WtMessageHandler.printToLogFile("Element: Name: " + element.getResourceURL() + "(Type: " + element.getType() + ")");
        }
      }
    } catch (Throwable e) {
      WtMessageHandler.printException(e);
    }
  }
  
  public void resetToolbar() {
    try {
      XUIConfigurationManager confMan;
      WtMessageHandler.printToLogFile("WtToolBar: resetToobar called");
      confMan = getUIConfigManagerDoc(xContext);
      if (confMan == null) {
        WtMessageHandler.printToLogFile("WtToolbar: makeToolbar: Cannot create configuration manager");
        return;
      }
      XIndexAccess settings = confMan.createSettings();
      XIndexContainer elementsContainer = UnoRuntime.queryInterface(XIndexContainer.class, settings);
      addToolbarButtons(elementsContainer, confMan);
    } catch (Throwable e) {
      WtMessageHandler.printException(e);
    }
  }
  
  private void addToolbarButtons(XIndexContainer elementsContainer, XUIConfigurationManager confMan) throws Throwable {
    WtMessageHandler.printToLogFile("WtToolBar: addToolbarButtons called");
    int j = 0;
    Language lang = document.getLanguage();
    boolean hasStatisticalStyleRules;
    if (document.getMultiDocumentsHandler().isBackgroundCheckOff()) {
      hasStatisticalStyleRules = false;
    } else {
      hasStatisticalStyleRules = WtOfficeTools.hasStatisticalStyleRules(lang);
    }
    WtConfiguration config = document.getMultiDocumentsHandler().getConfiguration();
    boolean aiSupport = config.useAiSupport() || config.useAiImgSupport() || config.useAiTtsSupport();
    PropertyValue[] itemProps = makeBarItem(WtMenus.LT_NEXT_ERROR_COMMAND, MESSAGES.getString("loContextMenuNextError"));
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_CHECKDIALOG_COMMAND, MESSAGES.getString("loContextMenuGrammarCheck"));
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_CHECKAGAINDIALOG_COMMAND, MESSAGES.getString("loContextMenuGrammarCheckAgain"));
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_REFRESH_CHECK_COMMAND, MESSAGES.getString("loContextMenuRefreshCheck"));
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    if (document.getMultiDocumentsHandler().isBackgroundCheckOff()) {
      itemProps = makeBarItem(WtMenus.LT_BACKGROUND_CHECK_ON_COMMAND, MESSAGES.getString("loMenuEnableBackgroundCheck"));
    } else {
      itemProps = makeBarItem(WtMenus.LT_BACKGROUND_CHECK_OFF_COMMAND, MESSAGES.getString("loMenuDisableBackgroundCheck"));
    }
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_RESET_IGNORE_PERMANENT_COMMAND, MESSAGES.getString("loMenuResetIgnorePermanent"));
    elementsContainer.insertByIndex(j, itemProps);
/*        TODO: Add sub toolbars:
    if(!document.getMultiDocumentsHandler().getDisabledRulesMap(null).isEmpty()) {
      j++;
      itemProps = makeBarItem(WtMenus.LT_ACTIVATE_RULES_COMMAND, MESSAGES.getString("loContextMenuActivateRule"));
      elementsContainer.insertByIndex(j, itemProps);
    }
    if(config.getDefinedProfiles().size() > 1) {
      j++;
      itemProps = makeBarItem(WtMenus.LT_PROFILES_COMMAND, MESSAGES.getString("loMenuChangeProfiles"));
      elementsContainer.insertByIndex(j, itemProps);
    }
*/
    j++;
    itemProps = makeBarItem(WtMenus.LT_OPTIONS_COMMAND, MESSAGES.getString("loContextMenuOptions"));
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_ABOUT_COMMAND, MESSAGES.getString("loContextMenuAbout"));
    elementsContainer.insertByIndex(j, itemProps);
    if (hasStatisticalStyleRules) {
      j++;
      itemProps = makeBarItem(WtMenus.LT_STATISTICAL_ANALYSES_COMMAND, MESSAGES.getString("loStatisticalAnalysis"));
      elementsContainer.insertByIndex(j, itemProps);
    }
    if (aiSupport) {
      j++;
      XIndexAccess settings = confMan.createSettings();
      XIndexContainer aiElementsContainer = UnoRuntime.queryInterface(XIndexContainer.class, settings);
      addAiToolbarButtons(aiElementsContainer, config);
      if (confMan.hasSettings(WT_NEW_AI_TOOLBAR_URL)) {
        confMan.replaceSettings(WT_NEW_AI_TOOLBAR_URL, aiElementsContainer);
      } else {
        confMan.insertSettings(WT_NEW_AI_TOOLBAR_URL, aiElementsContainer);
      }
      itemProps = makeBarItem(WT_NEW_AI_TOOLBAR_URL, MESSAGES.getString("loMenuAiGeneralCommand"));
      elementsContainer.insertByIndex(j, itemProps);
    }
  }
  
  private void addAiToolbarButtons(XIndexContainer elementsContainer, WtConfiguration config) throws Throwable {
    int j = 0;
    PropertyValue[] itemProps;
    if (config.useAiSupport()) {
      itemProps = makeBarItem(WtMenus.LT_AI_BETTER_STYLE, MESSAGES.getString("loMenuAiStyleCommand"));
      elementsContainer.insertByIndex(j, itemProps);
      j++;
      itemProps = makeBarItem(WtMenus.LT_AI_REFORMULATE_TEXT, MESSAGES.getString("loMenuAiReformulateCommand"));
      elementsContainer.insertByIndex(j, itemProps);
      j++;
      itemProps = makeBarItem(WtMenus.LT_AI_EXPAND_TEXT, MESSAGES.getString("loMenuAiExpandCommand"));
      elementsContainer.insertByIndex(j, itemProps);
      j++;
      itemProps = makeBarItem(WtMenus.LT_AI_SYNONYMS_OF_WORD, MESSAGES.getString("loMenuAiSynnomsOfWordCommand"));
      elementsContainer.insertByIndex(j, itemProps);
      j++;
      itemProps = makeBarItem(WtMenus.LT_AI_SUMMARY, MESSAGES.getString("loMenuAiSummaryCommand"));
      elementsContainer.insertByIndex(j, itemProps);
      j++;
      itemProps = makeBarItem(WtMenus.LT_AI_TRANSLATE_TEXT, MESSAGES.getString("loMenuAiTranslateCommand"));
      elementsContainer.insertByIndex(j, itemProps);
      j++;
    }
    if (config.useAiTtsSupport()) {
      itemProps = makeBarItem(WtMenus.LT_AI_TEXT_TO_SPEECH, MESSAGES.getString("loMenuAiTextToSpeechCommand"));
      elementsContainer.insertByIndex(j, itemProps);
      j++;
    }
    if (config.useAiSupport() || config.useAiTtsSupport()) {
      itemProps = makeBarItem(WtMenus.LT_AI_GENERAL_COMMAND, MESSAGES.getString("loMenuAiGeneralCommand"));
      elementsContainer.insertByIndex(j, itemProps);
    }
  }
  
/*  
  private void setToolbarName(String toolbarName, XLayoutManager layoutManager) {
    XUIElement oLtBar = layoutManager.getElement(toolbarName);
    XUIElementSettings oLtBarSettings = UnoRuntime.queryInterface(XUIElementSettings.class, oLtBar);
    XIndexAccess oLtBarAccess = oLtBarSettings.getSettings(true);
    XIndexContainer oLtBarContainer = UnoRuntime.queryInterface(XIndexContainer.class, oLtBar);
    oLtBarContainer.insertByIndex(0, oLtBarContainer);
    
  }
*/  
  private int getIndexOfItem (XIndexContainer elementsContainer, String command) throws Throwable {
    for (int i = 0; i < elementsContainer.getCount(); i++) {
      PropertyValue[] itemProps = (PropertyValue[]) elementsContainer.getByIndex(i);
      for (PropertyValue prop : itemProps) {
        if ("CommandURL".equals(prop.Name)) {
          if (command.equals(prop.Value)) {
            return i;
          } else {
            break;
          }
        }
      }
    }
    return -1;
  }
  
  private PropertyValue[] makeToolbarProperties(String cmd) {
    //  properties of toolbar
    PropertyValue[] props = new PropertyValue[3];
    props[0] = new PropertyValue();
    props[0].Name = "CommandURL";
    props[0].Value = cmd;
    props[1] = new PropertyValue();
    props[1].Name = "Persistent";
    props[1].Value = true;
    props[2] = new PropertyValue();
    props[2].Name = "UIName";
    props[2].Value = "WritingToolNew";
    return props;
  }

  private PropertyValue[] makeBarItem(String cmd, String itemName) {
    // properties for a toolbar item using a name and an image
    // problem: image does not appear next to text on toolbar
    PropertyValue[] props = new PropertyValue[5];

    props[0] = new PropertyValue();
    props[0].Name = "CommandURL";
    props[0].Value = cmd;

    props[1] = new PropertyValue();
    props[1].Name = "Label";
    props[1].Value = itemName;

    props[2] = new PropertyValue();
    props[2].Name = "Type";
    props[2].Value = ItemType.DEFAULT;  // 0;

    props[3] = new PropertyValue();
    props[3].Name = "Visible";
    props[3].Value = true;

    props[4] = new PropertyValue();
    props[4].Name = "Style";
    props[4].Value = ItemStyle.DRAW_FLAT + ItemStyle.ALIGN_LEFT + 
                     ItemStyle.AUTO_SIZE + ItemStyle.ICON;
//                         + ItemStyle.TEXT;

    return props;
  }
  
  private PropertyValue[] makeContainerItem(String cmd, String itemName) {
    // properties for a toolbar item using a name and an image
    // problem: image does not appear next to text on toolbar
    PropertyValue[] props = new PropertyValue[5];

    props[0] = new PropertyValue();
    props[0].Name = "CommandURL";
    props[0].Value = cmd;

    props[1] = new PropertyValue();
    props[1].Name = "Label";
    props[1].Value = itemName;

    props[2] = new PropertyValue();
    props[2].Name = "Type";
    props[2].Value = ItemType.DEFAULT;  // 0;

    props[3] = new PropertyValue();
    props[3].Name = "Visible";
    props[3].Value = true;

    props[4] = new PropertyValue();
    props[4].Name = "Style";
    props[4].Value = ItemStyle.DRAW_FLAT + ItemStyle.ALIGN_LEFT + 
                     ItemStyle.AUTO_SIZE + ItemStyle.ICON;
//                         + ItemStyle.TEXT;

    return props;
  }
  
  private XUIConfigurationManager getUIConfigManagerDoc(XComponentContext xContext) throws Exception {

    XMultiComponentFactory mcFactory = xContext.getServiceManager();
  
    Object o = mcFactory.createInstanceWithContext("com.sun.star.ui.ModuleUIConfigurationManagerSupplier", xContext);     
    XModuleUIConfigurationManagerSupplier xSupplier = 
        UnoRuntime.queryInterface(XModuleUIConfigurationManagerSupplier.class, o);
  
    return xSupplier.getUIConfigurationManager(WRITER_SERVICE);
  }
  
  private XUIElementFactoryRegistration getUIElementFactoryRegistration(XComponentContext xContext) throws Exception {

    XMultiComponentFactory mcFactory = xContext.getServiceManager();
  
    Object o = mcFactory.createInstanceWithContext("com.sun.star.ui.UIElementFactoryManager", xContext);     
    return UnoRuntime.queryInterface(XUIElementFactoryRegistration.class, o);
    
  }

  private void printRegisteredUIElementFactories(XComponentContext xContext) throws Exception {
    XUIElementFactoryRegistration uiFactoryRegistration = getUIElementFactoryRegistration(xContext);
    PropertyValue[][] props = uiFactoryRegistration.getRegisteredFactories();
    for (int i = 0; i < props.length; i++) {
      for (int j = 0; j < props[i].length; j++) {
        WtMessageHandler.printToLogFile("Prop[" + i + "][" + j + "]: " + props[i][j].Name + ": " + props[i][j].Value);
      }
    }
  }
  
  private XFrame getComponentFrame() {
    XComponent xComponent = document.getXComponent();
    if (xComponent == null) {
      WtMessageHandler.printToLogFile("SingleDocument: setDokumentListener: XComponent not found!");
      return null;
    }
    XModel xModel = UnoRuntime.queryInterface(XModel.class, xComponent);
    if (xModel == null) {
      WtMessageHandler.printToLogFile("SingleDocument: setDokumentListener: XModel not found!");
      return null;
    }
    XController xController = xModel.getCurrentController();
    if (xController == null) {
      WtMessageHandler.printToLogFile("SingleDocument: setDokumentListener: XController not found!");
      return null;
    }
    return xController.getFrame();
  }

  private XWindow getComponentWindow() {
    XFrame frame = getComponentFrame();
    if (frame == null) {
      return null;
    }
    return frame.getComponentWindow();
  }

  private XLayoutManager getLayoutManager() {
    try {
      XFrame frame = getComponentFrame();
      if (frame == null) {
        return null;
      }
      XPropertySet propSet = UnoRuntime.queryInterface(XPropertySet.class, frame);
      if (propSet == null) {
        return null;
      }
      return UnoRuntime.queryInterface(XLayoutManager.class,  propSet.getPropertyValue("LayoutManager"));
    } catch (Exception e) {
      WtMessageHandler.printException(e);
    }
    return null;
  }

  private XLayoutManager getLayoutManager(XFrame frame) {
    try {
      if (frame == null) {
        frame = getComponentFrame();
      }
      if (frame == null) {
        return null;
      }
      XPropertySet propSet = UnoRuntime.queryInterface(XPropertySet.class, frame);
      if (propSet == null) {
        return null;
      }
      return UnoRuntime.queryInterface(XLayoutManager.class,  propSet.getPropertyValue("LayoutManager"));
    } catch (Exception e) {
      WtMessageHandler.printException(e);
    }
    return null;
  }

  public static void printUICmds(XUIConfigurationManager configMan, String uiElemName) 
      throws IllegalArgumentException, NoSuchElementException, IndexOutOfBoundsException, WrappedTargetException {
    // print every command used by the toolbar whose resource name is uiElemName
    XIndexAccess settings = configMan.getSettings(uiElemName, true);
    int numSettings = settings.getCount();
    WtMessageHandler.printToLogFile("No. of elements in \"" + uiElemName + "\" toolbar: " + numSettings);
    for (int i = 0; i < numSettings; i++) { 
      PropertyValue[] settingProps =  UnoRuntime.queryInterface(PropertyValue[].class, settings.getByIndex(i));
      // Props.showProps("Settings " + i, settingProps);
      WtMessageHandler.printToLogFile("Properties for \"Settings" + i + "\":");
      if (settingProps == null)
        WtMessageHandler.printToLogFile("  none found");
      else {
        for (PropertyValue prop : settingProps) {
          WtMessageHandler.printToLogFile("  " + prop.Name + ": " + propValueToString(prop.Value));
        }
        WtMessageHandler.printToLogFile("");
      }
    }
  //    Object val = Props.getValue("CommandURL", settingProps);
  //    WtMessageHandler.printToLogFile(i + ") " + propValueToString(val));
  }

  public static String propValueToString(Object val) {
    if (val == null) {
      return null;
    }
    if (val instanceof String[]) {
      return Arrays.toString((String[])val);
    } else if (val instanceof PropertyValue[]) {
      PropertyValue[] ps = (PropertyValue[])val;
      StringBuilder sb = new StringBuilder("[");
      for (PropertyValue p : ps) {
        sb.append("    " + p.Name + " = " + p.Value);
      }
      sb.append("  ]");
      return sb.toString();
    } else {
      return val.toString();
    }
  }
  /**
   * class to test for text changes in shapes 
   */
  private class WtDoLayout extends Thread {

    XLayoutManager layoutManager;

    WtDoLayout (XLayoutManager layoutManager) {
      this.layoutManager = layoutManager;
    }

    @Override
    public void run() {
      try {
        Thread.sleep(1000);
      } catch (Throwable e) {
        WtMessageHandler.showError(e);
      }
      layoutManager.doLayout();
    }
    
  }

  public class WtToolbarFactory extends WeakBase implements XUIElementFactory {
    
    @Override
    public XUIElement createUIElement(String  resourceURL, PropertyValue[] arguments)
        throws NoSuchElementException, IllegalArgumentException {
      if (!resourceURL.startsWith("private:resource/toolbar/WtToolbarFactory")) {
        throw new NoSuchElementException(resourceURL, this);
      }

      XWindow parentWindow = null;
      for (int i = 0; i < arguments.length; i++) {
        if (arguments[i].Name.equals("ParentWindow")) {
          parentWindow = UnoRuntime.queryInterface(XWindow.class, arguments[i].Value);
          break;
        }
      }

      return new WtToolbarContainer(xContext, parentWindow, resourceURL);
    }
    
  }
  
  public class WtToolbarContainer implements XUIElement {
    
    XWindow parentWindow;
    String resourceURL;

    public WtToolbarContainer(XComponentContext xContext, XWindow parentWindow, String resourceURL) {
      this.parentWindow = parentWindow;
      this.resourceURL = resourceURL;
    }

    @Override
    public XFrame getFrame() {
      return getComponentFrame();
    }

    @Override
    public Object getRealInterface() {
      return this;
    }

    @Override
    public String getResourceURL() {
      return resourceURL;
    }

    @Override
    public short getType() {
      return UIElementType.TOOLBAR;
    }
    
  }
  
}


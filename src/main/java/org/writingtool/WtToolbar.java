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

import org.languagetool.Language;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XIndexContainer;
import com.sun.star.frame.XController;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XLayoutManager;
import com.sun.star.frame.XModel;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.ui.ItemStyle;
import com.sun.star.ui.ItemType;
import com.sun.star.ui.XModuleUIConfigurationManagerSupplier;
import com.sun.star.ui.XUIConfigurationManager;
import com.sun.star.ui.XUIElement;
import com.sun.star.ui.XUIElementSettings;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/** 
 *  Class to add a dynamic WritingTool Toolbar
 *  since 26.1
 */
public class WtToolbar {
  
  private static boolean debugMode = false;
  
  private static final ResourceBundle MESSAGES = WtOfficeTools.getMessageBundle();
  private static final String WRITER_SERVICE = "com.sun.star.text.TextDocument";
  private static final String WT_TOOLBAR_URL = "private:resource/toolbar/WtToolbarFactory.toolbar";
  private static final String WT_AI_TOOLBAR_URL = "private:resource/toolbar/WtToolbarFactory?aiGeneralCommand";
  private XComponentContext xContext;
  private WtSingleDocument document;
  private XIndexAccess settings;

  WtToolbar(XComponentContext xContext, WtSingleDocument document) {
    this.xContext = xContext;
    this.document = document;
    createToolbar();
  }
  
  private void createToolbar() {
    try {
      
      XLayoutManager layoutManager = getLayoutManager();
      if (layoutManager.getElement(WT_TOOLBAR_URL) == null) {
        layoutManager.destroyElement(WT_TOOLBAR_URL);
        WtMessageHandler.printToLogFile("WtToolbar: createToolbar: XLayoutManager: destroy Element: " + WT_TOOLBAR_URL);
      }
      layoutManager.createElement(WT_TOOLBAR_URL);
      if (layoutManager.isElementVisible(WT_TOOLBAR_URL)) {
        layoutManager.hideElement(WT_TOOLBAR_URL);
        layoutManager.showElement(WT_TOOLBAR_URL);
      }

      XUIConfigurationManager confMan = getUIConfigManagerDoc(xContext);
      if (confMan == null) {
        WtMessageHandler.printToLogFile("WtToolbar: makeToolbar: Cannot create configuration manager");
        return;
      }
      settings = confMan.createSettings();
      addToolbarButtons((XIndexContainer)settings, null);
      if (confMan.hasSettings(WT_TOOLBAR_URL)) {
        confMan.replaceSettings(WT_TOOLBAR_URL, settings);
      } else {
        confMan.insertSettings(WT_TOOLBAR_URL, settings);
      }
      if (debugMode) {
        XUIElement wtToolbar = layoutManager.getElement(WT_TOOLBAR_URL);
        XUIElementSettings xWtBarSettings = UnoRuntime.queryInterface(XUIElementSettings.class, wtToolbar);
        XIndexAccess UISettings = ((XUIElementSettings) xWtBarSettings).getSettings(true);
        for (int i = 0; i < UISettings.getCount(); i++) {
          PropertyValue[] propVal = (PropertyValue[]) UISettings.getByIndex(i);
          for (int k = 0; k < propVal.length; k++) {
            WtMessageHandler.printToLogFile("WtToolbar: makeToolbar: " + i + ".: Property: Name: " + propVal[k].Name + ", Handle: " + propVal[k].Handle 
                + ", Value: " + propVal[k].Value + ", State: " + propVal[k].State);
          }
        }
        XPropertySet propset = UnoRuntime.queryInterface(XPropertySet.class, UISettings);
        WtMessageHandler.printToLogFile("WtToolbar: makeToolbar: propset.value (UIName): " + propset.getPropertyValue("UIName"));
      }
    } catch (Throwable e) {
      WtMessageHandler.printException(e);
    }
  }
  
  public void resetToolbar() {
    try {
      if (debugMode) {
        WtMessageHandler.printToLogFile("WtToolBar: resetToolbar called");
      }
      addToolbarButtons((XIndexContainer)settings, null);
      XUIConfigurationManager confMan = getUIConfigManagerDoc(xContext);
      confMan.replaceSettings(WT_TOOLBAR_URL, settings);
    } catch (Throwable e) {
      WtMessageHandler.printException(e);
    }
  }
    
  private void addToolbarButtons(XIndexContainer elementsContainer, XUIConfigurationManager confMan) throws Throwable {
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
    boolean isBackgroundCheckOff = document.getMultiDocumentsHandler().isBackgroundCheckOff();
    int num = elementsContainer.getCount();
    for (int i = num - 1; i >= 0; i--) {
      elementsContainer.removeByIndex(i);
    }
    PropertyValue[] itemProps = makeBarItem(WtMenus.LT_NEXT_ERROR_COMMAND, MESSAGES.getString("loContextMenuNextError"), true);
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_CHECKDIALOG_COMMAND, MESSAGES.getString("loContextMenuGrammarCheck"), true);
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_CHECKAGAINDIALOG_COMMAND, MESSAGES.getString("loContextMenuGrammarCheckAgain"), true);
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_REFRESH_CHECK_COMMAND, MESSAGES.getString("loContextMenuRefreshCheck"), true);
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_BACKGROUND_CHECK_ON_COMMAND, MESSAGES.getString("loMenuEnableBackgroundCheck"), isBackgroundCheckOff);
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_BACKGROUND_CHECK_OFF_COMMAND, MESSAGES.getString("loMenuDisableBackgroundCheck"), !isBackgroundCheckOff);
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_RESET_IGNORE_PERMANENT_COMMAND, MESSAGES.getString("loMenuResetIgnorePermanent"), true);
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
    itemProps = makeBarItem(WtMenus.LT_OPTIONS_COMMAND, MESSAGES.getString("loContextMenuOptions"), true);
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_ABOUT_COMMAND, MESSAGES.getString("loContextMenuAbout"), true);
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_STATISTICAL_ANALYSES_COMMAND, MESSAGES.getString("loStatisticalAnalysis"), hasStatisticalStyleRules);
    elementsContainer.insertByIndex(j, itemProps);
    if (confMan != null) {
      j++;
      XIndexAccess settings = confMan.createSettings();
      XIndexContainer aiElementsContainer = UnoRuntime.queryInterface(XIndexContainer.class, settings);
      addAiToolbarButtons(aiElementsContainer, config);
      if (confMan.hasSettings(WT_AI_TOOLBAR_URL)) {
        confMan.replaceSettings(WT_AI_TOOLBAR_URL, aiElementsContainer);
      } else {
        confMan.insertSettings(WT_AI_TOOLBAR_URL, aiElementsContainer);
      }
      itemProps = makeBarItem(WT_AI_TOOLBAR_URL, MESSAGES.getString("loMenuAiGeneralCommand"), aiSupport);
      elementsContainer.insertByIndex(j, itemProps);
    }
  }
  
  private void addAiToolbarButtons(XIndexContainer elementsContainer, WtConfiguration config) throws Throwable {
    int j = 0;
    PropertyValue[] itemProps;
    itemProps = makeBarItem(WtMenus.LT_AI_BETTER_STYLE, MESSAGES.getString("loMenuAiStyleCommand"), config.useAiSupport());
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_AI_REFORMULATE_TEXT, MESSAGES.getString("loMenuAiReformulateCommand"), config.useAiSupport());
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_AI_EXPAND_TEXT, MESSAGES.getString("loMenuAiExpandCommand"), config.useAiSupport());
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_AI_SYNONYMS_OF_WORD, MESSAGES.getString("loMenuAiSynnomsOfWordCommand"), config.useAiSupport());
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_AI_SUMMARY, MESSAGES.getString("loMenuAiSummaryCommand"), config.useAiSupport());
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_AI_TRANSLATE_TEXT, MESSAGES.getString("loMenuAiTranslateCommand"), config.useAiSupport());
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_AI_TEXT_TO_SPEECH, MESSAGES.getString("loMenuAiTextToSpeechCommand"), config.useAiTtsSupport());
    elementsContainer.insertByIndex(j, itemProps);
    j++;
    itemProps = makeBarItem(WtMenus.LT_AI_GENERAL_COMMAND, MESSAGES.getString("loMenuAiGeneralCommand"), config.useAiSupport() || config.useAiTtsSupport());
    elementsContainer.insertByIndex(j, itemProps);
  }
  
  private PropertyValue[] makeBarItem(String cmd, String itemName, boolean visible) {
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
    props[3].Name = "IsVisible";
    props[3].Value = visible;

    props[4] = new PropertyValue();
    props[4].Name = "Style";
    props[4].Value = ItemStyle.DRAW_FLAT + ItemStyle.ALIGN_LEFT + 
                     ItemStyle.AUTO_SIZE + ItemStyle.ICON;
//                         + ItemStyle.TEXT;

    return props;
  }
  
  @SuppressWarnings("unused")   //  TODO: test subcontainer
  private PropertyValue[] makeContainerItem(String cmd, String itemName, boolean visible, XIndexAccess ItemDescriptorContainer) {
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
    props[3].Name = "IsVisible";
    props[3].Value = true;

    props[4] = new PropertyValue();
    props[4].Name = "Style";
    props[4].Value = ItemStyle.DRAW_FLAT + ItemStyle.ALIGN_LEFT + 
                     ItemStyle.AUTO_SIZE + ItemStyle.ICON;
//                         + ItemStyle.TEXT;

    props[5] = new PropertyValue();
    props[5].Name = "ItemDescriptorContainer";
    props[5].Value = ItemDescriptorContainer;

    return props;
  }
  
  private XUIConfigurationManager getUIConfigManagerDoc(XComponentContext xContext) throws Exception {

    XMultiComponentFactory mcFactory = xContext.getServiceManager();
  
    Object o = mcFactory.createInstanceWithContext("com.sun.star.ui.ModuleUIConfigurationManagerSupplier", xContext);     
    XModuleUIConfigurationManagerSupplier xSupplier = 
        UnoRuntime.queryInterface(XModuleUIConfigurationManagerSupplier.class, o);
  
    return xSupplier.getUIConfigurationManager(WRITER_SERVICE);
  }
  
/*
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
*/  
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
  
}


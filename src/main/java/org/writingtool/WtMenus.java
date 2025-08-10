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

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.ResourceBundle;

import org.languagetool.Language;
import org.writingtool.aisupport.WtAiErrorDetection;
import org.writingtool.aisupport.WtAiParagraphChanging;
import org.writingtool.aisupport.WtAiTextToSpeech;
import org.writingtool.aisupport.WtAiTranslateDocument;
import org.writingtool.aisupport.WtAiErrorDetection.DetectionType;
import org.writingtool.aisupport.WtAiRemote.AiCommand;
import org.writingtool.config.WtConfiguration;
import org.writingtool.dialogs.WtAiSummaryDialog;
import org.writingtool.dialogs.WtStatAnDialog;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtViewCursorTools;
import org.writingtool.tools.WtOfficeTools.DocumentType;

import com.sun.star.awt.MenuEvent;
import com.sun.star.awt.MenuItemStyle;
import com.sun.star.awt.XMenuBar;
import com.sun.star.awt.XMenuListener;
import com.sun.star.awt.XPopupMenu;
import com.sun.star.beans.Property;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XIndexContainer;
import com.sun.star.frame.XController;
import com.sun.star.frame.XModel;
import com.sun.star.lang.DisposedException;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.text.TextMarkupType;
import com.sun.star.text.XTextRange;
import com.sun.star.ui.ActionTriggerSeparatorType;
import com.sun.star.ui.ContextMenuExecuteEvent;
import com.sun.star.ui.ContextMenuInterceptorAction;
import com.sun.star.ui.XContextMenuInterception;
import com.sun.star.ui.XContextMenuInterceptor;
import com.sun.star.uno.Any;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.view.XSelectionSupplier;

/**
 * Class of menus adding dynamic components 
 * to header menu and to context menu
 * @since 1.0
 * @author Fred Kruse
 */
public class WtMenus {
  
  public final static String LT_IGNORE_ONCE_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?ignoreOnce";
  public final static String LT_IGNORE_ALL_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?ignoreAll";
  public final static String LT_IGNORE_PERMANENT_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?ignorePermanent";
  public final static String LT_DEACTIVATE_RULE_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?deactivateRule";
  public final static String LT_MORE_INFO_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?moreInfo";
  public final static String LT_ACTIVATE_RULES_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?activateRules";
  public final static String LT_ACTIVATE_RULE_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?activateRule_";
  public final static String LT_REMOTE_HINT_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?remoteHint";   
  public final static String LT_RENEW_MARKUPS_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?renewMarkups";
  public final static String LT_ADD_TO_DICTIONARY_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?addToDictionary_";
  public final static String LT_NEXT_ERROR_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?nextError";
  public final static String LT_CHECKDIALOG_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?checkDialog";
  public final static String LT_CHECKAGAINDIALOG_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?checkAgainDialog";
  public static final String LT_STATISTICAL_ANALYSES_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?statisticalAnalyses";   
  public static final String LT_OFF_STATISTICAL_ANALYSES_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?offStatisticalAnalyses";   
  public static final String LT_RESET_IGNORE_PERMANENT_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?resetIgnorePermanent";   
  public static final String LT_PERMANENT_IGNORE_PARAGRAPH_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?permanentIgnoreParagraph";   
  public static final String LT_TOGGLE_BACKGROUND_CHECK_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?toggleNoBackgroundCheck";
  public static final String LT_BACKGROUND_CHECK_ON_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?backgroundCheckOn";
  public static final String LT_BACKGROUND_CHECK_OFF_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?backgroundCheckOff";
  public static final String LT_REFRESH_CHECK_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?refreshCheck";
  public static final String LT_ABOUT_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?about";
  public static final String LT_LANGUAGETOOL_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?lt";
  public static final String LT_OPTIONS_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?configure";
  public static final String LT_PROFILES_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?profiles";
  public static final String LT_PROFILE_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?profileChangeTo_";
  public static final String LT_NONE_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?noAction";
  public static final String LT_AI_MARK_ERRORS = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?aiAddErrorMarks";
  public static final String LT_AI_CORRECT_ERRORS = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?aiCorrectErrors";
  public static final String LT_AI_BETTER_STYLE = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?aiBetterStyle";
  public static final String LT_AI_REFORMULATE_TEXT = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?aiReformulateText";
  public static final String LT_AI_EXPAND_TEXT = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?aiAdvanceText";
  public static final String LT_AI_TRANSLATE_TEXT = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?aiTranslateText";
  public static final String LT_AI_TEXT_TO_SPEECH = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?aiTextToSpeech";
  public static final String LT_AI_GENERAL_COMMAND = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?aiGeneralCommand";
  public static final String LT_AI_REPLACE_WORD = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?aiReplaceWord_";
  public static final String LT_AI_SUMMARY = "service:" + WtOfficeTools.WT_SERVICE_NAME + "?aiSummary";
  
//  public static final String LT_MENU_REPLACE_COLON = "__|__";
  public static final String LT_MENU_REPLACE_COLON = ":";

  private static final ResourceBundle MESSAGES = WtOfficeTools.getMessageBundle();
  private static final int SUBMENU_ID_DIFF = 21;
  private static final short SUBMENU_ID_AI = 1000;
//  private static final String LT_TOOLBAR_URL = "private:resource/toolbar/addon_" + OfficeTools.WT_SERVICE_NAME + ".toolbar";
  
  // If anything on the position of LT menu is changed the following has to be changed
  private static final String TOOLS_COMMAND = ".uno:ToolsMenu";             //  Command to open tools menu
  private static final String COMMAND_BEFORE_LT_MENU = ".uno:LanguageMenu";   //  Command for Language Menu (LT menu is installed after)
                                                   
  private final static String IGNORE_ONCE_URL = "slot:201";
  private final static String ADD_TO_DICTIONARY_2 = "slot:2";
  private final static String ADD_TO_DICTIONARY_3 = "slot:3";
  private final static String SPEll_DIALOG_URL = "slot:4";
  
  private static boolean debugMode;   //  should be false except for testing
  private static boolean debugModeTm;  //  should be false except for testing
  private static boolean isRunning = false;
  
  private XComponentContext xContext;
  private XComponent xComponent;
  private WtSingleDocument document;
  private WtConfiguration config;
  private boolean isRemote;
  private LTHeadMenu ltHeadMenu;
  @SuppressWarnings("unused")
  private ContextMenuInterceptor ltContextMenu;

  WtMenus(XComponentContext xContext, WtSingleDocument document, WtConfiguration config) {
    try {
      debugMode = WtOfficeTools.DEBUG_MODE_LM;
      debugModeTm = WtOfficeTools.DEBUG_MODE_TM;
      this.document = document;
      this.xContext = xContext;
      this.xComponent = document.getXComponent();
      setConfigValues(config);
      if (document.getDocumentType() == DocumentType.WRITER) {
        ltHeadMenu = new LTHeadMenu(xComponent);
      }
      ltContextMenu = new ContextMenuInterceptor(xComponent);
      if (debugMode) {
        WtMessageHandler.printToLogFile("LanguageToolMenus initialised");
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }
  
  void setConfigValues(WtConfiguration config) {
    this.config = config;
    if (config != null) {
      isRemote = config.doRemoteCheck();
    }
  }
  
  void removeListener() {
    if (ltHeadMenu != null) {
      ltHeadMenu.removeListener();
    }
  }
  
  void addListener() {
    if (ltHeadMenu != null) {
      ltHeadMenu.addListener();
    }
  }
  
  String replaceColon (String str) {
    return str == null ? null : str.replace(":", LT_MENU_REPLACE_COLON);
  }
  
  /**
   * Class to add or change some items of the LT head menu
   */
  private class LTHeadMenu implements XMenuListener {
    XPopupMenu ltMenu = null;
    short toolsId = 0;
    short ltId = 0;
    short switchOffId = 0;
    short switchOffPos = 0;
    private XPopupMenu toolsMenu = null;
    private XPopupMenu xProfileMenu = null;
    private XPopupMenu xActivateRuleMenu = null;
    private XPopupMenu xAiSupportMenu = null;
    private List<String> definedProfiles = null;
    private String currentProfile = null;
    
    public LTHeadMenu(XComponent xComponent) {
      try {
        XMenuBar menubar = null;
        menubar = WtOfficeTools.getMenuBar(xComponent);
        if (menubar == null) {
          WtMessageHandler.printToLogFile("LanguageToolMenus: LTHeadMenu: Menubar is null");
          return;
        }
        for (short i = 0; i < menubar.getItemCount(); i++) {
          toolsId = menubar.getItemId(i);
          String command = menubar.getCommand(toolsId);
          if (TOOLS_COMMAND.equals(command)) {
            toolsMenu = menubar.getPopupMenu(toolsId);
            break;
          }
        }
        if (toolsMenu == null) {
          WtMessageHandler.printToLogFile("LanguageToolMenus: LTHeadMenu: Tools Menu is null");
          return;
        }
        for (short i = 0; i < toolsMenu.getItemCount(); i++) {
          String command = toolsMenu.getCommand(toolsMenu.getItemId(i));
          if (COMMAND_BEFORE_LT_MENU.equals(command)) {
            ltId = toolsMenu.getItemId((short) (i + 1));
            ltMenu = toolsMenu.getPopupMenu(ltId);
            break;
          }
        }
        if (ltMenu == null) {
          WtMessageHandler.printToLogFile("LanguageToolMenus: LTHeadMenu: LT Menu is null");
          return;
        }
        
        for (short i = 0; i < ltMenu.getItemCount(); i++) {
          String command = ltMenu.getCommand(ltMenu.getItemId(i));
          if (LT_OPTIONS_COMMAND.equals(command)) {
            switchOffId = (short)102;
            switchOffPos = (short)(i - 1);
            break;
          }
        }
        if (switchOffId == 0) {
          WtMessageHandler.printToLogFile("LanguageToolMenus: LTHeadMenu: switchOffId not found");
          return;
        }
        boolean hasStatisticalStyleRules = false;
        if (document.getDocumentType() == DocumentType.WRITER &&
            !document.getMultiDocumentsHandler().isBackgroundCheckOff()) {
          Language lang = document.getLanguage();
          if (lang != null) {
            hasStatisticalStyleRules = WtOfficeTools.hasStatisticalStyleRules(lang);
          }
        }
        if (hasStatisticalStyleRules) {
          ltMenu.insertItem((short)(switchOffId + 1), MESSAGES.getString("loStatisticalAnalysis") + " ...", 
              (short)0, switchOffPos);
          ltMenu.setCommand(switchOffId, LT_STATISTICAL_ANALYSES_COMMAND);
          switchOffPos++;
        }
        ltMenu.insertItem(switchOffId, MESSAGES.getString("loMenuResetIgnorePermanent"), (short)0, switchOffPos);
        ltMenu.setCommand(switchOffId, LT_RESET_IGNORE_PERMANENT_COMMAND);
        switchOffId--;
        switchOffPos++;
        ltMenu.insertItem(switchOffId, MESSAGES.getString("loMenuEnableBackgroundCheck"), (short)0, switchOffPos);
        if (document.getMultiDocumentsHandler().isBackgroundCheckOff()) {
          ltMenu.setCommand(switchOffId, LT_BACKGROUND_CHECK_ON_COMMAND);
        } else {
          ltMenu.setCommand(switchOffId, LT_BACKGROUND_CHECK_OFF_COMMAND);
        }
        toolsMenu.addMenuListener(this);
        ltMenu.addMenuListener(this);
        if (debugMode) {
          WtMessageHandler.printToLogFile("LanguageToolMenus: LTHeadMenu: Menu listener set");
        }
      } catch (Throwable t) {
        WtMessageHandler.showError(t);
      }
    }
    
    void removeListener() {
      if (toolsMenu != null) {
        toolsMenu.removeMenuListener(this);
      }
    }
    
    void addListener() {
      if (toolsMenu != null) {
        toolsMenu.addMenuListener(this);
      }
    }
    
    /**
     * Set the dynamic parts of the LT menu
     * placed as submenu at the LO/OO tools menu
     */
    private void setLtMenu() throws Throwable {
      if (document.getMultiDocumentsHandler().isBackgroundCheckOff()) {
        ltMenu.setItemText(switchOffId, MESSAGES.getString("loMenuEnableBackgroundCheck"));
      } else {
        ltMenu.setItemText(switchOffId, MESSAGES.getString("loMenuDisableBackgroundCheck"));
      }
      short profilesId = (short)(switchOffId + 10);
      short profilesPos = (short)(switchOffPos + 2);
      if (ltMenu.getItemId(profilesPos) != profilesId) {
        setProfileMenu(profilesId, profilesPos);
      }
      int nProfileItems = setProfileItems();
      setActivateRuleMenu((short)(switchOffPos + 3), (short)(switchOffId + 11), (short)(switchOffId + SUBMENU_ID_DIFF + nProfileItems));
      short nId = (short)(SUBMENU_ID_AI + 1);
//      short nPos = (short)(switchOffPos + 3);
      short aiPos = ltMenu.getItemPos((short)(nId + 1));
//      short aiAutoPos = ltMenu.getItemPos(nId);
/*      
      if (config.useAiSupport() && !config.aiAutoCorrect() && aiAutoPos < 1) {
        ltMenu.insertItem(nId, MESSAGES.getString("loMenuAiAddErrorMarks"), (short) 0, nPos);
        ltMenu.setCommand(nId, LT_AI_MARK_ERRORS);
        ltMenu.enableItem(nId , true);
        nPos++;
      } else if ((!config.useAiSupport() || config.aiAutoCorrect()) && aiAutoPos > 0) {
        ltMenu.removeItem(aiAutoPos, (short)1);
      }
*/      
      if ((config.useAiSupport() || config.useAiImgSupport()) && aiPos < 1) {
        setAIMenu((short)(switchOffPos + 3), SUBMENU_ID_AI, (short)(SUBMENU_ID_AI + 1));
/*
        nId++;
        ltMenu.insertItem(nId, MESSAGES.getString("loMenuAiGeneralCommand"), (short) 0, nPos);
        ltMenu.setCommand(nId, LT_AI_GENERAL_COMMAND);
        ltMenu.enableItem(nId , true);
*/
      } else if ((!config.useAiSupport() && !config.useAiImgSupport()) && aiPos > 0) {
        ltMenu.removeItem(aiPos, (short)1);
      }
    }
      
    /**
     * Set the profile menu
     * if there are more than one profiles defined at the LT configuration file
     */
    private void setProfileMenu(short profilesId, short profilesPos) throws Throwable {
      ltMenu.insertItem(profilesId, MESSAGES.getString("loMenuChangeProfiles"), MenuItemStyle.AUTOCHECK, profilesPos);
      xProfileMenu = WtOfficeTools.getPopupMenu(xContext);
      if (xProfileMenu == null) {
        WtMessageHandler.printToLogFile("LanguageToolMenus: setProfileMenu: Profile menu == null");
        return;
      }
      
      xProfileMenu.addMenuListener(this);

      ltMenu.setPopupMenu(profilesId, xProfileMenu);
    }
    
    /**
     * Set the items for different profiles 
     * if there are more than one defined at the LT configuration file
     */
    private int setProfileItems() throws Throwable {
      currentProfile = config.getCurrentProfile();
      definedProfiles = config.getDefinedProfiles();
      definedProfiles.sort(null);
      if (xProfileMenu != null) {
        xProfileMenu.removeItem((short)0, xProfileMenu.getItemCount());
        short nId = (short) (switchOffId + SUBMENU_ID_DIFF);
        short nPos = 0;
        xProfileMenu.insertItem(nId, MESSAGES.getString("guiUserProfile"), (short) 0, nPos);
        xProfileMenu.setCommand(nId, LT_PROFILE_COMMAND);
        if (currentProfile == null || currentProfile.isEmpty()) {
          xProfileMenu.enableItem(nId , false);
        } else {
          xProfileMenu.enableItem(nId , true);
        }
        if (definedProfiles != null) {
          for (int i = 0; i < definedProfiles.size(); i++) {
            nId++;
            nPos++;
            xProfileMenu.insertItem(nId, definedProfiles.get(i), (short) 0, nPos);
            xProfileMenu.setCommand(nId, LT_PROFILE_COMMAND + replaceColon(definedProfiles.get(i)));
            if (currentProfile != null && currentProfile.equals(definedProfiles.get(i))) {
              xProfileMenu.enableItem(nId , false);
            } else {
              xProfileMenu.enableItem(nId , true);
            }
          }
        }
      }
      return (definedProfiles == null ? 1 : definedProfiles.size() + 1);
    }

    /**
     * Run the actions defined in the profile menu
     */
    private void runProfileAction(String profile) throws Throwable {
      List<String> definedProfiles = config.getDefinedProfiles();
      if (profile != null && (definedProfiles == null || !definedProfiles.contains(profile))) {
        WtMessageHandler.showMessage("profile '" + profile + "' not found");
      } else {
        try {
          List<String> saveProfiles = new ArrayList<>();
          saveProfiles.addAll(config.getDefinedProfiles());
          config.initOptions();
          config.loadConfiguration(profile == null ? "" : profile);
          config.setCurrentProfile(profile);
          config.addProfiles(saveProfiles);
          config.saveConfiguration(document.getLanguage());
          document.getMultiDocumentsHandler().resetConfiguration();
        } catch (IOException e) {
          WtMessageHandler.showError(e);
        }
      }
    }
    
    /**
     * Set Activate Rule Submenu
     */
    private void setActivateRuleMenu(short pos, short id, short submenuStartId) throws Throwable {
      Map<String, String> deactivatedRulesMap = document.getMultiDocumentsHandler().getDisabledRulesMap(null);
      if (!deactivatedRulesMap.isEmpty()) {
        if (ltMenu.getItemPos(id) < 1) {
          ltMenu.insertItem(id, MESSAGES.getString("loContextMenuActivateRule"), MenuItemStyle.AUTOCHECK, pos);
          xActivateRuleMenu = WtOfficeTools.getPopupMenu(xContext);
          if (xActivateRuleMenu == null) {
            WtMessageHandler.printToLogFile("LanguageToolMenus: setActivateRuleMenu: activate rule menu == null");
            return;
          }
          xActivateRuleMenu.addMenuListener(this);
          ltMenu.setPopupMenu(id, xActivateRuleMenu);
        }
        xActivateRuleMenu.removeItem((short) 0, xActivateRuleMenu.getItemCount());
        short nId = submenuStartId;
        short nPos = 0;
        for (String ruleId : deactivatedRulesMap.keySet()) {
          xActivateRuleMenu.insertItem(nId, deactivatedRulesMap.get(ruleId), (short) 0, nPos);
          xActivateRuleMenu.setCommand(nId, LT_ACTIVATE_RULE_COMMAND + ruleId);
          xActivateRuleMenu.enableItem(nId , true);
          nId++;
          nPos++;
        }
      } else if (xActivateRuleMenu != null) {
        pos = ltMenu.getItemPos(id);
        ltMenu.removeItem(pos, (short)1);
        xActivateRuleMenu.removeItem((short) 0, xActivateRuleMenu.getItemCount());
        xActivateRuleMenu = null;
      }
    }

    /**
     * Set AI Submenu
     */
    private void setAIMenu(short pos, short id, short submenuStartId) throws Throwable {
      if (config.useAiSupport() || config.useAiImgSupport() || config.useAiTtsSupport()) {
        if (ltMenu.getItemPos(id) < 1) {
          ltMenu.insertItem(id, MESSAGES.getString("loMenuAiSupport"), MenuItemStyle.AUTOCHECK, pos);
          xAiSupportMenu = WtOfficeTools.getPopupMenu(xContext);
          if (xAiSupportMenu == null) {
            WtMessageHandler.printToLogFile("LanguageToolMenus: setAIMenu: AI support menu == null");
            return;
          }
          xAiSupportMenu.addMenuListener(this);
          ltMenu.setPopupMenu(id, xAiSupportMenu);
        }
        xAiSupportMenu.removeItem((short) 0, xAiSupportMenu.getItemCount());
        short nId = submenuStartId;
        short nPos = 0;
        if (config.useAiSupport()) {
          if (!config.aiAutoCorrect()) {
            xAiSupportMenu.insertItem(nId, MESSAGES.getString("loMenuAiAddErrorMarks"), (short) 0, nPos);
            xAiSupportMenu.setCommand(nId, LT_AI_MARK_ERRORS);
            xAiSupportMenu.enableItem(nId , true);
            nPos++;
          }
          nId++;
          xAiSupportMenu.insertItem(nId, MESSAGES.getString("loMenuAiStyleCommand"), (short) 0, nPos);
          xAiSupportMenu.setCommand(nId, LT_AI_BETTER_STYLE);
          xAiSupportMenu.enableItem(nId , true);
          nId++;
          nPos++;
          xAiSupportMenu.insertItem(nId, MESSAGES.getString("loMenuAiReformulateCommand"), (short) 0, nPos);
          xAiSupportMenu.setCommand(nId, LT_AI_REFORMULATE_TEXT);
          xAiSupportMenu.enableItem(nId , true);
          nId++;
          nPos++;
          xAiSupportMenu.insertItem(nId, MESSAGES.getString("loMenuAiExpandCommand"), (short) 0, nPos);
          xAiSupportMenu.setCommand(nId, LT_AI_EXPAND_TEXT);
          xAiSupportMenu.enableItem(nId , true);
          nId++;
          nPos++;
          xAiSupportMenu.insertItem(nId, MESSAGES.getString("loMenuAiSummaryCommand"), (short) 0, nPos);
          xAiSupportMenu.setCommand(nId, LT_AI_SUMMARY);
          xAiSupportMenu.enableItem(nId , true);
          nId++;
          nPos++;
          xAiSupportMenu.insertItem(nId, MESSAGES.getString("loMenuAiTranslateCommand"), (short) 0, nPos);
          xAiSupportMenu.setCommand(nId, LT_AI_TRANSLATE_TEXT);
          xAiSupportMenu.enableItem(nId , true);
          nId++;
          nPos++;
        }
        if (config.useAiTtsSupport()) {
          nId = SUBMENU_ID_AI + 8;
          xAiSupportMenu.insertItem(nId, MESSAGES.getString("loMenuAiTextToSpeechCommand"), (short) 0, nPos);
          xAiSupportMenu.setCommand(nId, LT_AI_TEXT_TO_SPEECH);
          xAiSupportMenu.enableItem(nId , true);
          nPos++;
        }
        if (config.useAiSupport() || config.useAiImgSupport()) {
          nId = SUBMENU_ID_AI + 9;
          xAiSupportMenu.insertItem(nId, MESSAGES.getString("loMenuAiGeneralCommand"), (short) 0, nPos);
          xAiSupportMenu.setCommand(nId, LT_AI_GENERAL_COMMAND);
          xAiSupportMenu.enableItem(nId , true);
        }
      } else if (xAiSupportMenu != null) {
        pos = ltMenu.getItemPos(id);
        ltMenu.removeItem(pos, (short)1);
        xAiSupportMenu.removeItem((short) 0, xActivateRuleMenu.getItemCount());
        xAiSupportMenu = null;
      }
    }

    @Override
    public void disposing(EventObject event) {
    }

    @Override
    public void itemActivated(MenuEvent event) {
      if (event.MenuId == 0) {
        try {
          setLtMenu();
        } catch (Throwable e) {
          WtMessageHandler.showError(e);
        }
      }
    }

    @Override
    public void itemDeactivated(MenuEvent event) {
    }

    @Override
    public void itemHighlighted(MenuEvent event) {
    }

    @Override
    public void itemSelected(MenuEvent event) {
      try {
        if (debugMode) {
          WtMessageHandler.printToLogFile("LanguageToolMenus: itemSelected: event id: " + ((int)event.MenuId));
        }
        if (event.MenuId == switchOffId) {
          if (document.getMultiDocumentsHandler().toggleNoBackgroundCheck()) {
            document.getMultiDocumentsHandler().resetCheck(); 
          }
        } else if (event.MenuId == switchOffId + 1) {
          document.resetIgnorePermanent();
        } else if (event.MenuId == switchOffId + 2) {
          WtStatAnDialog statAnDialog = new WtStatAnDialog(document);
          statAnDialog.start();
          return;
        } else if (event.MenuId > SUBMENU_ID_AI && event.MenuId < SUBMENU_ID_AI + 10) {
//          if (debugMode) {
            WtMessageHandler.printToLogFile("LanguageToolMenus: itemSelected: AI support: " + (event.MenuId - SUBMENU_ID_AI));
//          }
          if (event.MenuId == SUBMENU_ID_AI + 1) {
            WtAiErrorDetection aiError = new WtAiErrorDetection(document, config, 
                document.getMultiDocumentsHandler().getLanguageTool());
            aiError.addAiRuleMatchesForParagraph(DetectionType.GRAMMAR);
          } else if (event.MenuId == SUBMENU_ID_AI + 2) {
            WtAiParagraphChanging aiChange = new WtAiParagraphChanging(document, config, AiCommand.ImproveStyle);
            aiChange.start();
          } else if (event.MenuId == SUBMENU_ID_AI + 3) {
            WtAiParagraphChanging aiChange = new WtAiParagraphChanging(document, config, AiCommand.ReformulateText);
            aiChange.start();
          } else if (event.MenuId == SUBMENU_ID_AI + 4) {
            WtAiSummaryDialog dialog = document.getMultiDocumentsHandler().getAiSummaryDialog();
            if (dialog != null) {
              dialog.toFront();
            } else {
              dialog = new WtAiSummaryDialog(document, MESSAGES);
              document.getMultiDocumentsHandler().setAiSummaryDialog(dialog);
              dialog.start();
            }
          } else if (event.MenuId == SUBMENU_ID_AI + 5) {
            WtAiTranslateDocument aiTranslate = new WtAiTranslateDocument(document, MESSAGES);
            aiTranslate.start();
          } else if (event.MenuId == SUBMENU_ID_AI + 6) {
            WtAiTranslateDocument aiTranslate = new WtAiTranslateDocument(document, MESSAGES);
            aiTranslate.start();
          } else if (event.MenuId == SUBMENU_ID_AI + 8) {
            WtAiTextToSpeech aiTextToSpeech = new WtAiTextToSpeech(document, MESSAGES);
            aiTextToSpeech.start();
          } else if (event.MenuId == SUBMENU_ID_AI + 9){
            WtAiParagraphChanging aiChange = new WtAiParagraphChanging(document, config, AiCommand.GeneralAi);
            aiChange.start();
          }
        } else if (event.MenuId == switchOffId + SUBMENU_ID_DIFF) {
          runProfileAction(null);
        } else if (definedProfiles != null && event.MenuId > switchOffId + SUBMENU_ID_DIFF 
            && event.MenuId <= switchOffId + SUBMENU_ID_DIFF + definedProfiles.size()) {
          runProfileAction(definedProfiles.get(event.MenuId - switchOffId - 22));
        } else if (definedProfiles != null && event.MenuId > switchOffId + SUBMENU_ID_DIFF + definedProfiles.size()) {
          Map<String, String> deactivatedRulesMap = document.getMultiDocumentsHandler().getDisabledRulesMap(null);
          short j = (short)(switchOffId + SUBMENU_ID_DIFF + definedProfiles.size() + 1);
          for (String ruleId : deactivatedRulesMap.keySet()) {
            if(event.MenuId == j) {
              if (debugMode) {
                WtMessageHandler.printToLogFile("LanguageToolMenus: itemSelected: activate rule: " + ruleId);
              }
              document.getMultiDocumentsHandler().activateRule(ruleId);
              return;
            }
            j++;
          }
        }
      } catch (Throwable e) {
        WtMessageHandler.showError(e);
      }
    }

  }

  /** 
   * Class to add a LanguageTool Options item to the context menu
   * since 4.6
   */
  private class ContextMenuInterceptor implements XContextMenuInterceptor {
    
    public ContextMenuInterceptor() {}
    
    public ContextMenuInterceptor(XComponent xComponent) {
      try {
        XModel xModel = UnoRuntime.queryInterface(XModel.class, xComponent);
        if (xModel == null) {
          WtMessageHandler.printToLogFile("LanguageToolMenus: ContextMenuInterceptor: XModel not found!");
          return;
        }
        XController xController = xModel.getCurrentController();
        if (xController == null) {
          WtMessageHandler.printToLogFile("LanguageToolMenus: ContextMenuInterceptor: xController == null");
          return;
        }
        XContextMenuInterception xContextMenuInterception = UnoRuntime.queryInterface(XContextMenuInterception.class, xController);
        if (xContextMenuInterception == null) {
          WtMessageHandler.printToLogFile("LanguageToolMenus: ContextMenuInterceptor: xContextMenuInterception == null");
          return;
        }
        ContextMenuInterceptor aContextMenuInterceptor = new ContextMenuInterceptor();
        XContextMenuInterceptor xContextMenuInterceptor = 
            UnoRuntime.queryInterface(XContextMenuInterceptor.class, aContextMenuInterceptor);
        if (xContextMenuInterceptor == null) {
          WtMessageHandler.printToLogFile("LanguageToolMenus: ContextMenuInterceptor: xContextMenuInterceptor == null");
          return;
        }
        xContextMenuInterception.registerContextMenuInterceptor(xContextMenuInterceptor);
      } catch (DisposedException e) {
        WtMessageHandler.printToLogFile("WtMenus: Document is disposed");
      } catch (Throwable t) {
        WtMessageHandler.printException(t);
      }
    }
  
    /**
     * Add LT items to context menu
     */
    @Override
    public ContextMenuInterceptorAction notifyContextMenuExecute(ContextMenuExecuteEvent aEvent) {
      try {
        if (isRunning) {
          WtMessageHandler.printToLogFile("LanguageToolMenus: notifyContextMenuExecute: is running: no change in Menu");
          return ContextMenuInterceptorAction.IGNORED;
        }
        isRunning = true;
        long startTime = 0;
        if (debugModeTm) {
          startTime = System.currentTimeMillis();
          WtMessageHandler.printToLogFile("Generate context menu started");
        }
        XIndexContainer xContextMenu = aEvent.ActionTriggerContainer;
        if (xContextMenu == null) {
          return ContextMenuInterceptorAction.IGNORED;
        }
        if (debugMode) {
          WtMessageHandler.printToLogFile("LanguageToolMenus: notifyContextMenuExecute: get xContextMenu");
        }
        
        document.setMenuDocId();
        if (document.getDocumentType() == DocumentType.IMPRESS) {
          int nId = 0;
          XMultiServiceFactory xMenuElementFactory = UnoRuntime.queryInterface(XMultiServiceFactory.class, xContextMenu);
          XPropertySet xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
              xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
          xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("loContextMenuGrammarCheck"));
          xNewMenuEntry.setPropertyValue("CommandURL", LT_CHECKDIALOG_COMMAND);
          xContextMenu.insertByIndex(nId, xNewMenuEntry);
          nId++;
          if (config.useAiSupport() || config.useAiImgSupport()) {
            addAIMenuEntry(nId, xContextMenu, xMenuElementFactory);
            nId++;
          }
          XPropertySet xSeparator = UnoRuntime.queryInterface(XPropertySet.class,
              xMenuElementFactory.createInstance("com.sun.star.ui.ActionTriggerSeparator"));
          xSeparator.setPropertyValue("SeparatorType", ActionTriggerSeparatorType.LINE);
          xContextMenu.insertByIndex(nId, xSeparator);
          if (debugModeTm) {
            long runTime = System.currentTimeMillis() - startTime;
            if (runTime > WtOfficeTools.TIME_TOLERANCE) {
              WtMessageHandler.printToLogFile("Time to generate context menu (Impress): " + runTime);
            }
          }
          isRunning = false;
          if (debugMode) {
            WtMessageHandler.printToLogFile("LanguageToolMenus: notifyContextMenuExecute: execute modified for Impress");
          }
          return ContextMenuInterceptorAction.EXECUTE_MODIFIED;
        }
        
        int count = xContextMenu.getCount();
        
        if (debugMode) {
          for (int i = 0; i < count; i++) {
            Any a = (Any) xContextMenu.getByIndex(i);
            XPropertySet props = (XPropertySet) a.getObject();
            printProperties(props);
          }
        }

        //  Add LT Options Item if a Grammar or Spell error was detected
        XMultiServiceFactory xMenuElementFactory = UnoRuntime.queryInterface(XMultiServiceFactory.class, xContextMenu);
        for (int i = 0; i < count; i++) {
          Any a = (Any) xContextMenu.getByIndex(i);
          XPropertySet props = (XPropertySet) a.getObject();
          String str = null;
          if (props.getPropertySetInfo().hasPropertyByName("CommandURL")) {
            str = props.getPropertyValue("CommandURL").toString();
          }
          if (str != null && IGNORE_ONCE_URL.equals(str)) {
            int n;
            boolean isSpellError = false;
            for (n = i + 1; n < count; n++) {
              a = (Any) xContextMenu.getByIndex(n);
              XPropertySet tmpProps = (XPropertySet) a.getObject();
              if (tmpProps.getPropertySetInfo().hasPropertyByName("CommandURL")) {
                str = tmpProps.getPropertyValue("CommandURL").toString();
              }
              if (ADD_TO_DICTIONARY_2.equals(str) || ADD_TO_DICTIONARY_3.equals(str)) {
                isSpellError = true;
                String wrongWord = getSelectedWord(aEvent);
                if (debugMode) {
                  WtMessageHandler.printToLogFile("LanguageToolMenus: notifyContextMenuExecute: wrong word: " + wrongWord);
                }
                if (wrongWord != null && !wrongWord.isEmpty()) {
                  if (wrongWord.charAt(wrongWord.length() - 1) == '.') {
                    wrongWord= wrongWord.substring(0, wrongWord.length() - 1);
                  }
                  if (!wrongWord.isEmpty()) {
                    XIndexContainer xSubMenuContainer = (XIndexContainer)UnoRuntime.queryInterface(XIndexContainer.class,
                        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTriggerContainer"));
                    int j = 0;
                    for (String dict : WtDictionary.getUserDictionaries(xContext)) {
                      XPropertySet xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
                          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
                      xNewSubMenuEntry.setPropertyValue("Text", dict);
                      xNewSubMenuEntry.setPropertyValue("CommandURL", LT_ADD_TO_DICTIONARY_COMMAND + dict + ":" + wrongWord);
                      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
                      j++;
                    }
                    XPropertySet xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
                        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
                    xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("loContextMenuAddToDictionary"));
                    xNewMenuEntry.setPropertyValue( "SubContainer", (Object)xSubMenuContainer );
                    xContextMenu.removeByIndex(n);
                    xContextMenu.insertByIndex(n, xNewMenuEntry);
                  }
                  // add AI suggestion to the top of context menu (if there is any)
                  String aiSuggestion = getAiSuggestion();
                  if (aiSuggestion != null) {
                    for (int k = 0; k < count; k++) {
                      Any aa = (Any) xContextMenu.getByIndex(k);
                      XPropertySet ttProps = (XPropertySet) aa.getObject();
                      if (!ttProps.getPropertySetInfo().hasPropertyByName("Text")) {
                        break;
                      }
                      String suggestion = ttProps.getPropertyValue("Text").toString();
//                      WtMessageHandler.printToLogFile("WtMenus: suggestion: " + (suggestion == null ? "null" : suggestion)
//                          + ", CommandURL: " + ttProps.getPropertyValue("CommandURL").toString());
                      if (aiSuggestion.equals(suggestion)) {
                        xContextMenu.removeByIndex(k);
                        break;
                      }
                    }
                    XPropertySet xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
                        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
                    xNewMenuEntry.setPropertyValue("Text", aiSuggestion);
                    xNewMenuEntry.setPropertyValue( "CommandURL", LT_AI_REPLACE_WORD + aiSuggestion);
                    xContextMenu.insertByIndex(0, xNewMenuEntry);
                  }
                }
              } else if (SPEll_DIALOG_URL.equals(str)) {
                XPropertySet xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
                    xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
                xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("checkTextShortDesc"));
                xNewMenuEntry.setPropertyValue("CommandURL", LT_CHECKDIALOG_COMMAND);
                xContextMenu.removeByIndex(n);
                xContextMenu.insertByIndex(n, xNewMenuEntry);
                break;
              }
            }
            if (!isSpellError) {
              if (!config.filterOverlappingMatches()) {
                document.getErrorAndChangeRange(aEvent, false);
              }
              addLTMenus(i, count, null, xContextMenu, xMenuElementFactory, aEvent);
              if (document.getCurrentNumberOfParagraph() >= 0) {
                props.setPropertyValue("CommandURL", LT_IGNORE_ONCE_COMMAND);
              }

              if (debugModeTm) {
                long runTime = System.currentTimeMillis() - startTime;
                if (runTime > WtOfficeTools.TIME_TOLERANCE) {
                  WtMessageHandler.printToLogFile("Time to generate context menu (grammar error): " + runTime);
                }
              }
              isRunning = false;
              if (debugMode) {
                WtMessageHandler.printToLogFile("LanguageToolMenus: notifyContextMenuExecute: execute modified for Writer");
              }
              return ContextMenuInterceptorAction.EXECUTE_MODIFIED;
            }
          }
        }

        //  Workaround for LO 24.x
        WtProofreadingError error = document.getErrorAndChangeRange(aEvent, true);
        if (error != null) {
          addLTMenus(0, count, error, xContextMenu, xMenuElementFactory, aEvent);
          isRunning = false;
          return ContextMenuInterceptorAction.EXECUTE_MODIFIED;
        }
        
        //  Add LT Options Item for context menu without grammar error
        XPropertySet xSeparator = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTriggerSeparator"));
        xSeparator.setPropertyValue("SeparatorType", ActionTriggerSeparatorType.LINE);
        xContextMenu.insertByIndex(count, xSeparator);
        
        int nId = count + 1;
        if (isRemote) {
          XPropertySet xNewMenuEntry2 = UnoRuntime.queryInterface(XPropertySet.class,
              xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
          xNewMenuEntry2.setPropertyValue("Text", MESSAGES.getString("loMenuRemoteInfo"));
          xNewMenuEntry2.setPropertyValue("CommandURL", LT_REMOTE_HINT_COMMAND);
          xContextMenu.insertByIndex(nId, xNewMenuEntry2);
          nId++;
        }

        XPropertySet xNewMenuEntry4 = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewMenuEntry4.setPropertyValue("Text", MESSAGES.getString("loContextMenuRenewMarkups"));
        xNewMenuEntry4.setPropertyValue("CommandURL", LT_RENEW_MARKUPS_COMMAND);
        xContextMenu.insertByIndex(nId, xNewMenuEntry4);
        nId++;

        addLTMenuEntry(nId, xContextMenu, xMenuElementFactory, aEvent, true);
        if (config.useAiSupport() || config.useAiImgSupport()) {
          nId++;
          addAIMenuEntry(nId, xContextMenu, xMenuElementFactory);
        }
/*        
        XPropertySet xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("loContextMenuOptions"));
        xNewMenuEntry.setPropertyValue("CommandURL", LT_OPTIONS_URL);
        xContextMenu.insertByIndex(nId, xNewMenuEntry);
*/
        if (debugModeTm) {
          long runTime = System.currentTimeMillis() - startTime;
          if (runTime > WtOfficeTools.TIME_TOLERANCE) {
            WtMessageHandler.printToLogFile("Time to generate context menu (no grammar error): " + runTime);
          }
        }
        isRunning = false;
        if (debugMode) {
          WtMessageHandler.printToLogFile("LanguageToolMenus: notifyContextMenuExecute: execute modified for Writer (no grammar error)");
        }
        return ContextMenuInterceptorAction.CONTINUE_MODIFIED;

      } catch (Throwable t) {
        WtMessageHandler.printException(t);
      }
      isRunning = false;
      WtMessageHandler.printToLogFile("LanguageToolMenus: notifyContextMenuExecute: no change in Menu");
      return ContextMenuInterceptorAction.IGNORED;
    }
    
    private void addLTMenus(int n, int count, WtProofreadingError error, XIndexContainer xContextMenu, 
        XMultiServiceFactory xMenuElementFactory, ContextMenuExecuteEvent aEvent) throws Throwable {
      if (error != null) {
        for (int i = count - 1; i >= 0; i--) {
          xContextMenu.removeByIndex(i);
        }
        n = 0;
        XPropertySet xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewMenuEntry.setPropertyValue("Text", error.aShortComment);
        xNewMenuEntry.setPropertyValue("CommandURL", LT_NONE_COMMAND);
        xContextMenu.insertByIndex(n, xNewMenuEntry);

        n++;
        XPropertySet xSeparator = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTriggerSeparator"));
        xSeparator.setPropertyValue("SeparatorType", ActionTriggerSeparatorType.LINE);
        xContextMenu.insertByIndex(n, xSeparator);

        n++;
        xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("guiOOoIgnoreButton"));
        xNewMenuEntry.setPropertyValue("CommandURL", LT_IGNORE_ONCE_COMMAND);
        xContextMenu.insertByIndex(n, xNewMenuEntry);
      }

      XPropertySet xNewMenuEntry3 = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewMenuEntry3.setPropertyValue("Text", MESSAGES.getString("loContextMenuIgnorePermanent"));
      xNewMenuEntry3.setPropertyValue("CommandURL", LT_IGNORE_PERMANENT_COMMAND);
      xContextMenu.insertByIndex(n + 1, xNewMenuEntry3);
      
      if (error != null) {
        XPropertySet xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("guiOOoIgnoreAllButton"));
        xNewMenuEntry.setPropertyValue("CommandURL", LT_IGNORE_ALL_COMMAND);
        xContextMenu.insertByIndex(n + 2, xNewMenuEntry);
      }
      
      XPropertySet xNewMenuEntry1 = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewMenuEntry1.setPropertyValue("Text", MESSAGES.getString("loContextMenuDeactivateRule"));
      xNewMenuEntry1.setPropertyValue("CommandURL", LT_DEACTIVATE_RULE_COMMAND);
      xContextMenu.insertByIndex(n + 3, xNewMenuEntry1);
      
      int nId = n + 4;
      Map<String, String> deactivatedRulesMap = document.getMultiDocumentsHandler().getDisabledRulesMap(null);
      if (!deactivatedRulesMap.isEmpty()) {
        xContextMenu.insertByIndex(nId, createActivateRuleProfileItems(deactivatedRulesMap, xMenuElementFactory));
        nId++;
      }
      
      if (isRemote) {
        XPropertySet xNewMenuEntry2 = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewMenuEntry2.setPropertyValue("Text", MESSAGES.getString("loMenuRemoteInfo"));
        xNewMenuEntry2.setPropertyValue("CommandURL", LT_REMOTE_HINT_COMMAND);
        xContextMenu.insertByIndex(nId, xNewMenuEntry2);
        nId++;
      }
      
      XPropertySet xNewMenuEntry4 = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewMenuEntry4.setPropertyValue("Text", MESSAGES.getString("loContextMenuRenewMarkups"));
      xNewMenuEntry4.setPropertyValue("CommandURL", LT_RENEW_MARKUPS_COMMAND);
      xContextMenu.insertByIndex(nId, xNewMenuEntry4);
      nId++;

      List<String> definedProfiles = config.getDefinedProfiles();
      if (definedProfiles.size() > 1) {
        xContextMenu.insertByIndex(nId, createProfileItems(definedProfiles, xMenuElementFactory));
        nId++;
      }
      addLTMenuEntry(nId, xContextMenu, xMenuElementFactory, aEvent, false);
      nId++;
      addAIMenuEntry(nId, xContextMenu, xMenuElementFactory);
      
      XPropertySet xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("guiMore"));
      xNewMenuEntry.setPropertyValue("CommandURL", LT_MORE_INFO_COMMAND);
      xContextMenu.insertByIndex(1, xNewMenuEntry);
    }
    
    private void addLTMenuEntry(int nId, XIndexContainer xContextMenu, XMultiServiceFactory xMenuElementFactory,
            ContextMenuExecuteEvent aEvent, boolean showAll) throws Throwable {
      XIndexContainer xSubMenuContainer = (XIndexContainer)UnoRuntime.queryInterface(XIndexContainer.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTriggerContainer"));
      boolean hasStatisticalStyleRules;
      if (document.getDocumentType() == DocumentType.WRITER &&
          !document.getMultiDocumentsHandler().isBackgroundCheckOff()) {
        hasStatisticalStyleRules = WtOfficeTools.hasStatisticalStyleRules(document.getLanguage());
      } else {
        hasStatisticalStyleRules = false;
      }
      XPropertySet xNewSubMenuEntry;
      int j = 0;
      if (showAll) {
        xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("checkTextShortDesc"));
        xNewSubMenuEntry.setPropertyValue("CommandURL", LT_CHECKDIALOG_COMMAND);
        xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
        j++;
      }
      if (hasStatisticalStyleRules) {
        xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loStatisticalAnalysis"));
        xNewSubMenuEntry.setPropertyValue("CommandURL", LT_STATISTICAL_ANALYSES_COMMAND);
        xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
        j++;
      }
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      if (isSelectedRange(aEvent)) {
        xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuPermanentIgnoreRange"));
      } else {
        xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuPermanentIgnoreParagraph"));
      }
      xNewSubMenuEntry.setPropertyValue("CommandURL", LT_PERMANENT_IGNORE_PARAGRAPH_COMMAND);
      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
      j++;
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuResetIgnorePermanent"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", LT_RESET_IGNORE_PERMANENT_COMMAND);
      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
      j++;
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      if (document.getMultiDocumentsHandler().isBackgroundCheckOff()) {
        xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuEnableBackgroundCheck"));
        xNewSubMenuEntry.setPropertyValue("CommandURL", LT_BACKGROUND_CHECK_ON_COMMAND);
      } else {
        xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuDisableBackgroundCheck"));
        xNewSubMenuEntry.setPropertyValue("CommandURL", LT_BACKGROUND_CHECK_OFF_COMMAND);
      }
      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
      j++;
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loContextMenuRefreshCheck"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", LT_REFRESH_CHECK_COMMAND);
      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
      
      Map<String, String> deactivatedRulesMap = document.getMultiDocumentsHandler().getDisabledRulesMap(null);
      if (showAll && !deactivatedRulesMap.isEmpty()) {
        j++;
        xSubMenuContainer.insertByIndex(j, createActivateRuleProfileItems(deactivatedRulesMap, xMenuElementFactory));
      }
      
      List<String> definedProfiles = config.getDefinedProfiles();
      if (showAll && definedProfiles.size() > 1) {
        j++;
        xSubMenuContainer.insertByIndex(j, createProfileItems(definedProfiles, xMenuElementFactory));
      }

      j++;
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loContextMenuOptions"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", LT_OPTIONS_COMMAND);
      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
      j++;
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loContextMenuAbout"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", LT_ABOUT_COMMAND);
      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
      

      XPropertySet xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewMenuEntry.setPropertyValue("Text", WtOfficeTools.WT_NAME);
      xNewMenuEntry.setPropertyValue("CommandURL", LT_LANGUAGETOOL_COMMAND);
      xNewMenuEntry.setPropertyValue("SubContainer", (Object)xSubMenuContainer);
      xContextMenu.insertByIndex(nId, xNewMenuEntry);
    }
    
    private void addAIMenuEntry(int nId, XIndexContainer xContextMenu, XMultiServiceFactory xMenuElementFactory) throws Throwable {
      if (!config.useAiSupport() && !config.useAiImgSupport() && !config.useAiTtsSupport()) {
        return;
      }
/*      
      XPropertySet xNewMenuEntry;
      int j = nId;
      if (config.useAiSupport() && !config.aiAutoCorrect()) {
        xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiAddErrorMarks"));
        xNewMenuEntry.setPropertyValue("CommandURL", LT_AI_MARK_ERRORS);
        xContextMenu.insertByIndex(j, xNewMenuEntry);
        j++;
      }
      if (config.useAiSupport()) {
        xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiStyleCommand"));
        xNewMenuEntry.setPropertyValue("CommandURL", LT_AI_BETTER_STYLE);
        xContextMenu.insertByIndex(j, xNewMenuEntry);
        j++;
      }
      xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiGeneralCommand"));
      xNewMenuEntry.setPropertyValue("CommandURL", LT_AI_GENERAL_COMMAND);
      xContextMenu.insertByIndex(j, xNewMenuEntry);
*/
      XIndexContainer xSubMenuContainer = (XIndexContainer)UnoRuntime.queryInterface(XIndexContainer.class,
        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTriggerContainer"));
      XPropertySet xNewSubMenuEntry;
      int j = 0;
      if (config.useAiSupport() && !config.aiAutoCorrect()) {
        xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiAddErrorMarks"));
        xNewSubMenuEntry.setPropertyValue("CommandURL", LT_AI_MARK_ERRORS);
        xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
        j++;
      }
      if (config.useAiSupport()) {
        xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiStyleCommand"));
        xNewSubMenuEntry.setPropertyValue("CommandURL", LT_AI_BETTER_STYLE);
        xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
        j++;
        xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiReformulateCommand"));
        xNewSubMenuEntry.setPropertyValue("CommandURL", LT_AI_REFORMULATE_TEXT);
        xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
        j++;
        xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiExpandCommand"));
        xNewSubMenuEntry.setPropertyValue("CommandURL", LT_AI_EXPAND_TEXT);
        xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
        j++;
        xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiSummaryCommand"));
        xNewSubMenuEntry.setPropertyValue("CommandURL", LT_AI_SUMMARY);
        xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
        j++;
        xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
              xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiTranslateCommand"));
        xNewSubMenuEntry.setPropertyValue("CommandURL", LT_AI_TRANSLATE_TEXT);
        xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
        j++;
      }
      if (config.useAiTtsSupport()) {
        xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiTextToSpeechCommand"));
        xNewSubMenuEntry.setPropertyValue("CommandURL", LT_AI_TEXT_TO_SPEECH);
        xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
        j++;
      }
      if (config.useAiSupport() || config.useAiTtsSupport()) {
        xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiGeneralCommand"));
        xNewSubMenuEntry.setPropertyValue("CommandURL", LT_AI_GENERAL_COMMAND);
        xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
      }
      XPropertySet xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewMenuEntry.setPropertyValue("Text",  MESSAGES.getString("loMenuAiSupport"));
      xNewMenuEntry.setPropertyValue("CommandURL", LT_AI_GENERAL_COMMAND);
      xNewMenuEntry.setPropertyValue("SubContainer", (Object)xSubMenuContainer);
      xContextMenu.insertByIndex(nId, xNewMenuEntry);

   }
      
   private XPropertySet createActivateRuleProfileItems(Map<String, String> deactivatedRulesMap, 
        XMultiServiceFactory xMenuElementFactory) throws Throwable {
      XPropertySet xNewSubMenuEntry;
      XIndexContainer xRuleMenuContainer = (XIndexContainer)UnoRuntime.queryInterface(XIndexContainer.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTriggerContainer"));
      int nPos = 0;
      for (String ruleId : deactivatedRulesMap.keySet()) {
        xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewSubMenuEntry.setPropertyValue("Text", deactivatedRulesMap.get(ruleId));
        xNewSubMenuEntry.setPropertyValue("CommandURL", LT_ACTIVATE_RULE_COMMAND + ruleId);
        xRuleMenuContainer.insertByIndex(nPos, xNewSubMenuEntry);
        nPos++;
      }
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loContextMenuActivateRule"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", LT_ACTIVATE_RULES_COMMAND);
      xNewSubMenuEntry.setPropertyValue("SubContainer", (Object)xRuleMenuContainer);
      return xNewSubMenuEntry;
    }

    private XPropertySet createProfileItems(List<String> definedProfiles, 
        XMultiServiceFactory xMenuElementFactory) throws Exception {
      XPropertySet xNewSubMenuEntry;
      definedProfiles.sort(null);
      XIndexContainer xRuleMenuContainer = (XIndexContainer)UnoRuntime.queryInterface(XIndexContainer.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTriggerContainer"));
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("guiUserProfile"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", LT_PROFILE_COMMAND);
      xRuleMenuContainer.insertByIndex(0, xNewSubMenuEntry);
      for (int i = 0; i < definedProfiles.size(); i++) {
        xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
        xNewSubMenuEntry.setPropertyValue("Text", definedProfiles.get(i));
        xNewSubMenuEntry.setPropertyValue("CommandURL", LT_PROFILE_COMMAND + replaceColon(definedProfiles.get(i)));
        xRuleMenuContainer.insertByIndex(i + 1, xNewSubMenuEntry);
      }
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuChangeProfiles"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", LT_PROFILES_COMMAND);
      xNewSubMenuEntry.setPropertyValue("SubContainer", (Object)xRuleMenuContainer);
      return xNewSubMenuEntry;
    }

    /**
     * get selected word
     */
    private String getSelectedWord(ContextMenuExecuteEvent aEvent) {
      try {
        XSelectionSupplier xSelectionSupplier = aEvent.Selection;
        Object selection = xSelectionSupplier.getSelection();
        XIndexAccess xIndexAccess = UnoRuntime.queryInterface(XIndexAccess.class, selection);
        if (xIndexAccess == null) {
          WtMessageHandler.printToLogFile("LanguageToolMenus: getSelectedWord: xIndexAccess == null");
          return null;
        }
        XTextRange xTextRange = UnoRuntime.queryInterface(XTextRange.class, xIndexAccess.getByIndex(0));
        if (xTextRange == null) {
          WtMessageHandler.printToLogFile("LanguageToolMenus: getSelectedWord: xTextRange == null");
          return null;
        }
        return xTextRange.getString();
      } catch (Throwable t) {
        WtMessageHandler.printException(t);
      }
      return null;
    }

    /**
     * is selected range
     */
    private boolean isSelectedRange(ContextMenuExecuteEvent aEvent) {
      String selection = getSelectedWord(aEvent);
      if (selection == null || selection.length() == 0) {
        return false;
      }
      return true;
    }

    /**
     * get AI suggestion
     */
    private String getAiSuggestion() {
      try {
        if (config.useAiSupport() && config.aiAutoCorrect() && !document.getMultiDocumentsHandler().isBackgroundCheckOff()) {
          WtViewCursorTools viewCursor = new WtViewCursorTools(xComponent);
          int y = document.getDocumentCache().getFlatParagraphNumber(viewCursor.getViewCursorParagraph());
          int x = viewCursor.getViewCursorCharacter();
          List<WtProofreadingError> errors = document.getParagraphsCache().get(WtOfficeTools.CACHE_AI).getErrorsAtPosition(y, x);
          if (errors != null) {
            for (WtProofreadingError error : errors) {
              if (error.nErrorType == TextMarkupType.SPELLCHECK 
                  && error.aSuggestions.length > 0 && !error.aSuggestions[0].isBlank()) {
                  return error.aSuggestions[0];
              }
            }
          }
        }
      } catch (Throwable t) {
        WtMessageHandler.printException(t);
      }
      return null;
    }

    /**
     * Print properties in debug mode
     */
    private void printProperties(XPropertySet props) throws Throwable {
      Property[] propInfo = props.getPropertySetInfo().getProperties();
      for (Property property : propInfo) {
        WtMessageHandler.printToLogFile("LanguageToolMenus: Property: Name: " + property.Name + ", Type: " + property.Type);
      }
      if (props.getPropertySetInfo().hasPropertyByName("Text")) {
        WtMessageHandler.printToLogFile("LanguageToolMenus: Property: Name: " + props.getPropertyValue("Text"));
      }
      if (props.getPropertySetInfo().hasPropertyByName("CommandURL")) {
        WtMessageHandler.printToLogFile("LanguageToolMenus: Property: CommandURL: " + props.getPropertyValue("CommandURL"));
      }
    }

  }

}

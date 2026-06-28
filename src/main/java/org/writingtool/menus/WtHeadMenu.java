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
package org.writingtool.menus;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.ResourceBundle;

import org.languagetool.Language;
import org.writingtool.WtProtocolHandler;
import org.writingtool.WtSingleDocument;
import org.writingtool.WtDocumentsHandler.BackgroundCheck;
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
import org.writingtool.tools.WtOfficeTools.DocumentType;

import com.sun.star.awt.MenuEvent;
import com.sun.star.awt.MenuItemStyle;
import com.sun.star.awt.XMenuBar;
import com.sun.star.awt.XMenuListener;
import com.sun.star.awt.XPopupMenu;
import com.sun.star.lang.EventObject;
import com.sun.star.uno.XComponentContext;

/**
 * Class to add or change some items of the LT head menu
 * @since 26.7
 * @author Fred Kruse
 */
public class WtHeadMenu implements XMenuListener {

  private static final ResourceBundle MESSAGES = WtOfficeTools.getMessageBundle();
  // If anything on the position of LT menu is changed the following has to be changed
  private static final String TOOLS_COMMAND = ".uno:ToolsMenu";             //  Command to open tools menu
  private static final String COMMAND_BEFORE_WT_MENU = ".uno:LanguageMenu";   //  Command for Language Menu (LT menu is installed after)
                                                   
  //private static final short SWITCH_OFF_ID = 102;
  private static final short SUBMENU_ID_DIFF = 21;
  private static final short SUBMENU_ID_AI = 1000;

  private static boolean debugMode;   //  should be false except for testing

  private XComponentContext xContext;
  private WtSingleDocument document;
  private WtConfiguration config;
  private XPopupMenu ltMenu = null;
  private short toolsId = 0;
  private short ltId = 0;
  private short switchOffId = 0;
  private short switchOffPos = 0;
  private XPopupMenu toolsMenu = null;
  private XPopupMenu xProfileMenu = null;
  private XPopupMenu xActivateRuleMenu = null;
  private XPopupMenu xAiSupportMenu = null;
  private List<String> definedProfiles = null;
  private String currentProfile = null;
  
  public WtHeadMenu(WtSingleDocument document, WtConfiguration config, XComponentContext xContext) {
    this.document = document;
    this.config = config;
    this.xContext = xContext;
    debugMode = WtOfficeTools.DEBUG_MODE_LM;
    try {
      XMenuBar menubar = null;
      menubar = WtOfficeTools.getMenuBar(document.getXComponent());
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
        if (COMMAND_BEFORE_WT_MENU.equals(command)) {
          ltId = toolsMenu.getItemId((short) (i + 1));
          ltMenu = toolsMenu.getPopupMenu(ltId);
          break;
        }
      }
      if (ltMenu == null) {
        WtMessageHandler.printToLogFile("LanguageToolMenus: LTHeadMenu: LT Menu is null");
        return;
      }
      switchOffId = 0; 
      for (short i = 0; i < ltMenu.getItemCount(); i++) {
        String command = ltMenu.getCommand(ltMenu.getItemId(i));
        short nId = ltMenu.getItemId(i);
        if (nId >= switchOffId) {
          switchOffId += (short)1;
        }
        if (WtProtocolHandler.WT_OPTIONS_COMMAND.equals(command)) {
//          switchOffId = SWITCH_OFF_ID;
          switchOffPos = (short)(i - 1);
//          break;
        }
      }
      if (switchOffId == 0) {
        WtMessageHandler.printToLogFile("LanguageToolMenus: LTHeadMenu: switchOffId not found");
        return;
      }
      
      switchOffId += 5;
      
      boolean hasStatisticalStyleRules = false;
      if (document.getDocumentType() == DocumentType.WRITER &&
          !document.getMultiDocumentsHandler().isBackgroundCheckOff()) {
        Language lang = document.getLanguage();
        if (lang != null) {
          hasStatisticalStyleRules = WtOfficeTools.hasStatisticalStyleRules(lang);
        }
      }
      if (hasStatisticalStyleRules) {
        ltMenu.insertItem(switchOffId, MESSAGES.getString("statAnalysisDialog") + " ...", 
            (short)0, switchOffPos);
        ltMenu.setCommand(switchOffId, WtProtocolHandler.WT_STATISTICAL_ANALYSES_COMMAND);
        switchOffId++;
        switchOffPos++;
      }
      ltMenu.insertItem(switchOffId, MESSAGES.getString("loMenuChangeQuotes"), (short)0, switchOffPos);
      ltMenu.setCommand(switchOffId, WtProtocolHandler.WT_CHANGE_QUOTES_COMMAND);
      switchOffId++;
      switchOffPos++;
      ltMenu.insertItem(switchOffId, MESSAGES.getString("loMenuResetIgnorePermanent"), (short)0, switchOffPos);
      ltMenu.setCommand(switchOffId, WtProtocolHandler.WT_RESET_IGNORE_PERMANENT_COMMAND);
      switchOffId++;
      switchOffPos++;
      ltMenu.insertItem(switchOffId, MESSAGES.getString("loMenuEnableBackgroundCheck"), (short)0, switchOffPos);
      if (document.getMultiDocumentsHandler().isBackgroundCheckOff()) {
        ltMenu.setCommand(switchOffId, WtProtocolHandler.WT_BACKGROUND_CHECK_ON_COMMAND);
      } else {
        ltMenu.setCommand(switchOffId, WtProtocolHandler.WT_BACKGROUND_CHECK_OFF_COMMAND);
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
  
  /**
   * Set configuration values
   */
  public void setConfigValues(WtConfiguration config) {
    this.config = config;
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
    boolean isActivateRuleMenue = setActivateRuleMenu((short)(switchOffPos + 3), (short)(switchOffId + 11), (short)(switchOffId + SUBMENU_ID_DIFF + nProfileItems));
    short aiPos = ltMenu.getItemPos(SUBMENU_ID_AI);
    if ((config.useAiSupport() || config.useAiImgSupport()) && aiPos < 1) {
      short aiPosNew = (short)(switchOffPos + 3);
      if (isActivateRuleMenue) {
        aiPosNew++;
      }
      setAIMenu(aiPosNew, SUBMENU_ID_AI, (short)(SUBMENU_ID_AI + 1));
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
    ltMenu.setCommand(profilesId, WtProtocolHandler.WT_PROFILES_COMMAND);
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
      xProfileMenu.insertItem(nId, MESSAGES.getString("allDialogDefaultProfile"), (short) 0, nPos);
      xProfileMenu.setCommand(nId, WtProtocolHandler.WT_PROFILE_COMMAND);
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
          xProfileMenu.setCommand(nId, WtProtocolHandler.WT_PROFILE_COMMAND + WtMenus.replaceColon(definedProfiles.get(i)));
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
        document.getMultiDocumentsHandler().resetAiResultCaches();
        document.getMultiDocumentsHandler().resetGrammarCheckConfiguration();
      } catch (IOException e) {
        WtMessageHandler.showError(e);
      }
    }
  }
  
  /**
   * Set Activate Rule Submenu
   */
  private boolean setActivateRuleMenu(short pos, short id, short submenuStartId) throws Throwable {
    Map<String, String> deactivatedRulesMap = document.getMultiDocumentsHandler().getDisabledRulesMap(null);
    if (!deactivatedRulesMap.isEmpty()) {
      if (ltMenu.getItemPos(id) < 1 || xActivateRuleMenu == null) {
        ltMenu.insertItem(id, MESSAGES.getString("loContextMenuActivateRule"), MenuItemStyle.AUTOCHECK, pos);
        xActivateRuleMenu = WtOfficeTools.getPopupMenu(xContext);
        if (xActivateRuleMenu == null) {
          WtMessageHandler.printToLogFile("LanguageToolMenus: setActivateRuleMenu: activate rule menu == null");
          return false;
        }
        xActivateRuleMenu.addMenuListener(this);
        ltMenu.setPopupMenu(id, xActivateRuleMenu);
        ltMenu.setCommand(id, WtProtocolHandler.WT_ACTIVATE_RULES_COMMAND);
      }
      xActivateRuleMenu.removeItem((short) 0, xActivateRuleMenu.getItemCount());
      short nId = submenuStartId;
      short nPos = 0;
      for (String ruleId : deactivatedRulesMap.keySet()) {
        xActivateRuleMenu.insertItem(nId, deactivatedRulesMap.get(ruleId), (short) 0, nPos);
        xActivateRuleMenu.setCommand(nId, WtProtocolHandler.WT_ACTIVATE_RULE_COMMAND + ruleId);
        xActivateRuleMenu.enableItem(nId , true);
        nId++;
        nPos++;
      }
      return true;
    } else if (xActivateRuleMenu != null) {
      pos = ltMenu.getItemPos(id);
      ltMenu.removeItem(pos, (short)1);
      xActivateRuleMenu.removeItem((short) 0, xActivateRuleMenu.getItemCount());
      xActivateRuleMenu = null;
    }
    return false;
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
          xAiSupportMenu.setCommand(nId, WtProtocolHandler.WT_AI_MARK_ERRORS_COMMAND);
          xAiSupportMenu.enableItem(nId , true);
          nPos++;
        }
        nId++;
        xAiSupportMenu.insertItem(nId, MESSAGES.getString("loMenuAiStyleCommand"), (short) 0, nPos);
        xAiSupportMenu.setCommand(nId, WtProtocolHandler.WT_AI_BETTER_STYLE_COMMAND);
        xAiSupportMenu.enableItem(nId , true);
        nId++;
        nPos++;
        xAiSupportMenu.insertItem(nId, MESSAGES.getString("loMenuAiReformulateCommand"), (short) 0, nPos);
        xAiSupportMenu.setCommand(nId, WtProtocolHandler.WT_AI_REFORMULATE_TEXT_COMMAND);
        xAiSupportMenu.enableItem(nId , true);
        nId++;
        nPos++;
        xAiSupportMenu.insertItem(nId, MESSAGES.getString("loMenuAiExpandCommand"), (short) 0, nPos);
        xAiSupportMenu.setCommand(nId, WtProtocolHandler.WT_AI_EXPAND_TEXT_COMMAND);
        xAiSupportMenu.enableItem(nId , true);
        nId++;
        nPos++;
        xAiSupportMenu.insertItem(nId, MESSAGES.getString("loMenuAiSummaryCommand"), (short) 0, nPos);
        xAiSupportMenu.setCommand(nId, WtProtocolHandler.WT_AI_SUMMARY_COMMAND);
        xAiSupportMenu.enableItem(nId , true);
        nId++;
        nPos++;
        xAiSupportMenu.insertItem(nId, MESSAGES.getString("loMenuAiTranslateCommand"), (short) 0, nPos);
        xAiSupportMenu.setCommand(nId, WtProtocolHandler.WT_AI_TRANSLATE_TEXT_COMMAND);
        xAiSupportMenu.enableItem(nId , true);
        nId++;
        nPos++;
      }
      if (config.useAiTtsSupport()) {
        nId = SUBMENU_ID_AI + 8;
        xAiSupportMenu.insertItem(nId, MESSAGES.getString("loMenuAiTextToSpeechCommand"), (short) 0, nPos);
        xAiSupportMenu.setCommand(nId, WtProtocolHandler.WT_AI_TEXT_TO_SPEECH_COMMAND);
        xAiSupportMenu.enableItem(nId , true);
        nPos++;
      }
      if (config.useAiSupport() || config.useAiImgSupport()) {
        nId = SUBMENU_ID_AI + 9;
        xAiSupportMenu.insertItem(nId, MESSAGES.getString("loMenuAiGeneralCommand"), (short) 0, nPos);
        xAiSupportMenu.setCommand(nId, WtProtocolHandler.WT_AI_GENERAL_COMMAND);
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
      document.setMenuDocId();
      if (event.MenuId == switchOffId) {
        BackgroundCheck check = document.getMultiDocumentsHandler().isBackgroundCheckOff() ? BackgroundCheck.ON : BackgroundCheck.OFF;
        if (document.getMultiDocumentsHandler().toggleNoBackgroundCheck(check)) {
          document.getMultiDocumentsHandler().resetCheck(); 
        }
      } else if (event.MenuId == switchOffId - 1) {
        document.resetIgnorePermanent();
      } else if (event.MenuId == switchOffId - 2) {
        document.getMultiDocumentsHandler().runChangeQuotes();
        return;
      } else if (event.MenuId == switchOffId - 3) {
        WtStatAnDialog statAnDialog = new WtStatAnDialog(document);
        statAnDialog.start();
        return;
      } else if (event.MenuId > SUBMENU_ID_AI && event.MenuId < SUBMENU_ID_AI + 10) {
//        if (debugMode) {
          WtMessageHandler.printToLogFile("LanguageToolMenus: itemSelected: AI support: " + (event.MenuId - SUBMENU_ID_AI));
//        }
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
          WtAiParagraphChanging aiChange = new WtAiParagraphChanging(document, config, AiCommand.ExpandText);
          aiChange.start();
        } else if (event.MenuId == SUBMENU_ID_AI + 5) {
          WtAiSummaryDialog dialog = document.getMultiDocumentsHandler().getAiSummaryDialog();
          if (dialog != null) {
            dialog.toFront();
          } else {
            dialog = new WtAiSummaryDialog(document, MESSAGES);
            document.getMultiDocumentsHandler().setAiSummaryDialog(dialog);
            dialog.start();
          }
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

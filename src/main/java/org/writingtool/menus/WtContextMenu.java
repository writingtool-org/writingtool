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

import java.util.List;
import java.util.Map;
import java.util.ResourceBundle;

import org.writingtool.WritingTool;
import org.writingtool.WtDictionary;
import org.writingtool.WtDocumentsHandler;
import org.writingtool.WtProofreadingError;
import org.writingtool.WtProtocolHandler;
import org.writingtool.WtSingleDocument;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtViewCursorTools;
import org.writingtool.tools.WtOfficeTools.DocumentType;

import com.sun.star.beans.Property;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XIndexContainer;
import com.sun.star.frame.XController;
import com.sun.star.frame.XModel;
import com.sun.star.lang.DisposedException;
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
import com.sun.star.view.XSelectionSupplier;

/** 
 * Class to add WritingTool items to the context menu
 * @since 24.7
 * @author Fred Kruse
 */
public class WtContextMenu implements XContextMenuInterceptor {
  
  private static final ResourceBundle MESSAGES = WtOfficeTools.getMessageBundle();

  private static boolean debugMode;   //  should be false except for testing
  private static boolean debugModeTm;  //  should be false except for testing
  
  private final static String IGNORE_ONCE_URL = "slot:201";
  private final static String ADD_TO_DICTIONARY_2 = "slot:2";
  private final static String ADD_TO_DICTIONARY_3 = "slot:3";
  private final static String SPEll_DIALOG_URL = "slot:4";
  
  private static boolean isRunning = false;

  private boolean isRemote = false;
  
  private WtSingleDocument document;
  private WtConfiguration config;

  public WtContextMenu() {}
  
  public WtContextMenu(WtSingleDocument document, WtConfiguration config) {
    this.document = document;
    WtMessageHandler.printToLogFile("WtContextMenu: document " + (document == null ? "==" : "!=" ) + " NULL");
    this.config = config;
    isRemote = config.doRemoteCheck();
    try {
      debugMode = WtOfficeTools.DEBUG_MODE_LM;
      debugModeTm = WtOfficeTools.DEBUG_MODE_TM;
      XModel xModel = UnoRuntime.queryInterface(XModel.class, document.getXComponent());
      if (xModel == null) {
        WtMessageHandler.printToLogFile("WtContextMenu: ContextMenuInterceptor: XModel not found!");
        return;
      }
      XController xController = xModel.getCurrentController();
      if (xController == null) {
        WtMessageHandler.printToLogFile("WtContextMenu: ContextMenuInterceptor: xController == null");
        return;
      }
      XContextMenuInterception xContextMenuInterception = UnoRuntime.queryInterface(XContextMenuInterception.class, xController);
      if (xContextMenuInterception == null) {
        WtMessageHandler.printToLogFile("WtContextMenu: ContextMenuInterceptor: xContextMenuInterception == null");
        return;
      }
      WtContextMenu aContextMenuInterceptor = new WtContextMenu();
      XContextMenuInterceptor xContextMenuInterceptor = 
          UnoRuntime.queryInterface(XContextMenuInterceptor.class, aContextMenuInterceptor);
      if (xContextMenuInterceptor == null) {
        WtMessageHandler.printToLogFile("WtContextMenu: ContextMenuInterceptor: xContextMenuInterceptor == null");
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
   * Set configuration values
   */
  public void setConfigValues(WtConfiguration config) {
    this.config = config;
    if (config != null) {
      isRemote = config.doRemoteCheck();
    }
  }
  
  /**
   * Add LT items to context menu
   */
  @Override
  public ContextMenuInterceptorAction notifyContextMenuExecute(ContextMenuExecuteEvent aEvent) {
    try {
      if (isRunning) {
        WtMessageHandler.printToLogFile("WtContextMenu: notifyContextMenuExecute: is running: no change in Menu");
        return ContextMenuInterceptorAction.IGNORED;
      }
      WtDocumentsHandler documents = WritingTool.getDocumentsHandler();
      if (documents == null) {
        WtMessageHandler.printToLogFile("WtContextMenu: notifyContextMenuExecute: DocumentsHandler == NULL: no change in Menu");
        return ContextMenuInterceptorAction.IGNORED;
      }
      document = documents.getCurrentDocument();
      config = documents.getConfiguration();
      isRemote = config.doRemoteCheck();
      
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
        WtMessageHandler.printToLogFile("WtContextMenu: notifyContextMenuExecute: get xContextMenu");
      }
      WtMessageHandler.printToLogFile("WtContextMenu: document " + (document == null ? "==" : "!=" ) + " NULL");
      document.setMenuDocId();
      if (document.getDocumentType() == DocumentType.IMPRESS) {
        try {
          int nId = 0;
          XMultiServiceFactory xMenuElementFactory = UnoRuntime.queryInterface(XMultiServiceFactory.class, xContextMenu);
          XPropertySet xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
              xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
          xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("loContextMenuGrammarCheck"));
          xNewMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_CHECKDIALOG_COMMAND);
          xContextMenu.insertByIndex(nId, xNewMenuEntry);
          nId++;
          if (config.useAiSupport() || config.useAiImgSupport()) {
            xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
                xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
            xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiGeneralCommand"));
            xNewMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_AI_GENERAL_COMMAND);
            xContextMenu.insertByIndex(nId, xNewMenuEntry);
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
            WtMessageHandler.printToLogFile("WtContextMenu: notifyContextMenuExecute: execute modified for Impress");
          }
          return ContextMenuInterceptorAction.EXECUTE_MODIFIED;
        } catch (Throwable t) {
          WtMessageHandler.printException(t);
          return ContextMenuInterceptorAction.IGNORED;
        }
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
                WtMessageHandler.printToLogFile("WtContextMenu: notifyContextMenuExecute: wrong word: " + wrongWord);
              }
              if (wrongWord != null && !wrongWord.isEmpty()) {
                if (wrongWord.charAt(wrongWord.length() - 1) == '.') {
                  wrongWord= wrongWord.substring(0, wrongWord.length() - 1);
                }
                if (!wrongWord.isEmpty()) {
                  XIndexContainer xSubMenuContainer = (XIndexContainer)UnoRuntime.queryInterface(XIndexContainer.class,
                      xMenuElementFactory.createInstance("com.sun.star.ui.ActionTriggerContainer"));
                  int j = 0;
                  for (String dict : WtDictionary.getUserDictionaries(document.getMultiDocumentsHandler().getContext())) {
                    XPropertySet xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
                        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
                    xNewSubMenuEntry.setPropertyValue("Text", dict);
                    xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_ADD_TO_DICTIONARY_COMMAND + dict + ":" + wrongWord);
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
//                    WtMessageHandler.printToLogFile("WtMenus: suggestion: " + (suggestion == null ? "null" : suggestion)
//                        + ", CommandURL: " + ttProps.getPropertyValue("CommandURL").toString());
                    if (aiSuggestion.equals(suggestion)) {
                      xContextMenu.removeByIndex(k);
                      break;
                    }
                  }
                  XPropertySet xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
                      xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
                  xNewMenuEntry.setPropertyValue("Text", aiSuggestion);
                  xNewMenuEntry.setPropertyValue( "CommandURL", WtProtocolHandler.WT_AI_REPLACE_WORD_COMMAND + aiSuggestion);
                  xContextMenu.insertByIndex(0, xNewMenuEntry);
                }
              }
            } else if (SPEll_DIALOG_URL.equals(str)) {
              XPropertySet xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
                  xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
              xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("loContextMenuCheckText"));
              xNewMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_CHECKDIALOG_COMMAND);
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
              props.setPropertyValue("CommandURL", WtProtocolHandler.WT_IGNORE_ONCE_COMMAND);
            }

            if (debugModeTm) {
              long runTime = System.currentTimeMillis() - startTime;
              if (runTime > WtOfficeTools.TIME_TOLERANCE) {
                WtMessageHandler.printToLogFile("Time to generate context menu (grammar error): " + runTime);
              }
            }
            isRunning = false;
            if (debugMode) {
              WtMessageHandler.printToLogFile("WtContextMenu: notifyContextMenuExecute: execute modified for Writer");
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
        xNewMenuEntry2.setPropertyValue("CommandURL", WtProtocolHandler.WT_REMOTE_HINT_COMMAND);
        xContextMenu.insertByIndex(nId, xNewMenuEntry2);
        nId++;
      }

      XPropertySet xNewMenuEntry5 = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      if (isSelectedRange(aEvent)) {
        xNewMenuEntry5.setPropertyValue("Text", MESSAGES.getString("loMenuPermanentIgnoreRange"));
      } else {
        xNewMenuEntry5.setPropertyValue("Text", MESSAGES.getString("loMenuPermanentIgnoreParagraph"));
      }
      xNewMenuEntry5.setPropertyValue("CommandURL", WtProtocolHandler.WT_PERMANENT_IGNORE_PARAGRAPH_COMMAND);
      xContextMenu.insertByIndex(nId, xNewMenuEntry5);
      nId++;

      XPropertySet xNewMenuEntry4 = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewMenuEntry4.setPropertyValue("Text", MESSAGES.getString("loContextMenuRenewMarkups"));
      xNewMenuEntry4.setPropertyValue("CommandURL", WtProtocolHandler.WT_RENEW_MARKUPS_COMMAND);
      xContextMenu.insertByIndex(nId, xNewMenuEntry4);
      nId++;

      addLTMenuEntry(nId, xContextMenu, xMenuElementFactory, aEvent, true);
      if (config.useAiSupport() || config.useAiImgSupport() || config.useAiTtsSupport()) {
        nId++;
        addAIMenuEntry(nId, xContextMenu, xMenuElementFactory);
      }
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
        if (runTime > WtOfficeTools.TIME_TOLERANCE) {
          WtMessageHandler.printToLogFile("Time to generate context menu (no grammar error): " + runTime);
        }
      }
      isRunning = false;
      if (debugMode) {
        WtMessageHandler.printToLogFile("WtContextMenu: notifyContextMenuExecute: execute modified for Writer (no grammar error)");
      }
      return ContextMenuInterceptorAction.CONTINUE_MODIFIED;

    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
    isRunning = false;
    WtMessageHandler.printToLogFile("WtContextMenu: notifyContextMenuExecute: no change in Menu");
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
      xNewMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_NONE_COMMAND);
      xContextMenu.insertByIndex(n, xNewMenuEntry);

      n++;
      XPropertySet xSeparator = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTriggerSeparator"));
      xSeparator.setPropertyValue("SeparatorType", ActionTriggerSeparatorType.LINE);
      xContextMenu.insertByIndex(n, xSeparator);

      n++;
      xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("checkDialogIgnoreButton"));
      xNewMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_IGNORE_ONCE_COMMAND);
      xContextMenu.insertByIndex(n, xNewMenuEntry);
    }

    XPropertySet xNewMenuEntry3 = UnoRuntime.queryInterface(XPropertySet.class,
        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
    xNewMenuEntry3.setPropertyValue("Text", MESSAGES.getString("loContextMenuIgnorePermanent"));
    xNewMenuEntry3.setPropertyValue("CommandURL", WtProtocolHandler.WT_IGNORE_PERMANENT_COMMAND);
    xContextMenu.insertByIndex(n + 1, xNewMenuEntry3);
    
    if (error != null) {
      XPropertySet xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("checkDialogIgnoreAllButton"));
      xNewMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_IGNORE_ALL_COMMAND);
      xContextMenu.insertByIndex(n + 2, xNewMenuEntry);
    }
    
    XPropertySet xNewMenuEntry1 = UnoRuntime.queryInterface(XPropertySet.class,
        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
    xNewMenuEntry1.setPropertyValue("Text", MESSAGES.getString("loContextMenuDeactivateRule"));
    xNewMenuEntry1.setPropertyValue("CommandURL", WtProtocolHandler.WT_DEACTIVATE_RULE_COMMAND);
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
      xNewMenuEntry2.setPropertyValue("CommandURL", WtProtocolHandler.WT_REMOTE_HINT_COMMAND);
      xContextMenu.insertByIndex(nId, xNewMenuEntry2);
      nId++;
    }
    
    XPropertySet xNewMenuEntry4 = UnoRuntime.queryInterface(XPropertySet.class,
        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
    xNewMenuEntry4.setPropertyValue("Text", MESSAGES.getString("loContextMenuRenewMarkups"));
    xNewMenuEntry4.setPropertyValue("CommandURL", WtProtocolHandler.WT_RENEW_MARKUPS_COMMAND);
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
    xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("allDialogButtonMore"));
    xNewMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_MORE_INFO_COMMAND);
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
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loContextMenuCheckText"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_CHECKDIALOG_COMMAND);
      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
      j++;
    }
    if (hasStatisticalStyleRules) {
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("statAnalysisDialog"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_STATISTICAL_ANALYSES_COMMAND);
      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
      j++;
    }
    xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
    xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuChangeQuotes"));
    xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_CHANGE_QUOTES_COMMAND);
    xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
    j++;
/*      
    xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
    if (isSelectedRange(aEvent)) {
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuPermanentIgnoreRange"));
    } else {
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuPermanentIgnoreParagraph"));
    }
    xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_PERMANENT_IGNORE_PARAGRAPH_COMMAND);
    xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
    j++;
*/
    xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
    xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuResetIgnorePermanent"));
    xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_RESET_IGNORE_PERMANENT_COMMAND);
    xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
    j++;
    xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
    if (document.getMultiDocumentsHandler().isBackgroundCheckOff()) {
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuEnableBackgroundCheck"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_BACKGROUND_CHECK_ON_COMMAND);
    } else {
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuDisableBackgroundCheck"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_BACKGROUND_CHECK_OFF_COMMAND);
    }
    xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
    j++;
    xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
    xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loContextMenuRefreshCheck"));
    xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_REFRESH_CHECK_COMMAND);
    xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
    j++;
    xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
    xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loContextMenuRefreshMarkups"));
    xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_REFRESH_MARKUPS_COMMAND);
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
    xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_OPTIONS_COMMAND);
    xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
    j++;
    xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
    xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loContextMenuAbout"));
    xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_ABOUT_COMMAND);
    xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
    

    XPropertySet xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
    xNewMenuEntry.setPropertyValue("Text", WtOfficeTools.WT_NAME);
    xNewMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_LANGUAGETOOL_COMMAND);
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
      xNewMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_AI_MARK_ERRORS);
      xContextMenu.insertByIndex(j, xNewMenuEntry);
      j++;
    }
    if (config.useAiSupport()) {
      xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiStyleCommand"));
      xNewMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_AI_BETTER_STYLE);
      xContextMenu.insertByIndex(j, xNewMenuEntry);
      j++;
    }
    xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
    xNewMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiGeneralCommand"));
    xNewMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_AI_GENERAL_COMMAND);
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
      xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_AI_MARK_ERRORS_COMMAND);
      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
      j++;
    }
    if (config.useAiSupport()) {
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiStyleCommand"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_AI_BETTER_STYLE_COMMAND);
      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
      j++;
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiReformulateCommand"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_AI_REFORMULATE_TEXT_COMMAND);
      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
      j++;
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiExpandCommand"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_AI_EXPAND_TEXT_COMMAND);
      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
      j++;
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiSynnomsOfWordCommand"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_AI_SYNONYMS_OF_WORD_COMMAND);
      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
      j++;
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiSummaryCommand"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_AI_SUMMARY_COMMAND);
      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
      j++;
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
            xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiTranslateCommand"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_AI_TRANSLATE_TEXT_COMMAND);
      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
      j++;
    }
    if (config.useAiTtsSupport()) {
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiTextToSpeechCommand"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_AI_TEXT_TO_SPEECH_COMMAND);
      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
      j++;
    }
    if (config.useAiSupport() || config.useAiImgSupport()) {
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuAiGeneralCommand"));
      xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_AI_GENERAL_COMMAND);
      xSubMenuContainer.insertByIndex(j, xNewSubMenuEntry);
    }
    XPropertySet xNewMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
      xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
    xNewMenuEntry.setPropertyValue("Text",  MESSAGES.getString("loMenuAiSupport"));
    xNewMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_AI_GENERAL_COMMAND);
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
      xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_ACTIVATE_RULE_COMMAND + ruleId);
      xRuleMenuContainer.insertByIndex(nPos, xNewSubMenuEntry);
      nPos++;
    }
    xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
    xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loContextMenuActivateRule"));
    xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_ACTIVATE_RULES_COMMAND);
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
    xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("allDialogDefaultProfile"));
    xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_PROFILE_COMMAND);
    xRuleMenuContainer.insertByIndex(0, xNewSubMenuEntry);
    for (int i = 0; i < definedProfiles.size(); i++) {
      xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
          xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
      xNewSubMenuEntry.setPropertyValue("Text", definedProfiles.get(i));
      xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_PROFILE_COMMAND + WtMenus.replaceColon(definedProfiles.get(i)));
      xRuleMenuContainer.insertByIndex(i + 1, xNewSubMenuEntry);
    }
    xNewSubMenuEntry = UnoRuntime.queryInterface(XPropertySet.class,
        xMenuElementFactory.createInstance("com.sun.star.ui.ActionTrigger"));
    xNewSubMenuEntry.setPropertyValue("Text", MESSAGES.getString("loMenuChangeProfiles"));
    xNewSubMenuEntry.setPropertyValue("CommandURL", WtProtocolHandler.WT_PROFILES_COMMAND);
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
        WtMessageHandler.printToLogFile("WtContextMenu: getSelectedWord: xIndexAccess == null");
        return null;
      }
      XTextRange xTextRange = UnoRuntime.queryInterface(XTextRange.class, xIndexAccess.getByIndex(0));
      if (xTextRange == null) {
        WtMessageHandler.printToLogFile("WtContextMenu: getSelectedWord: xTextRange == null");
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
        WtViewCursorTools viewCursor = new WtViewCursorTools(document.getXComponent());
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
      WtMessageHandler.printToLogFile("WtContextMenu: Property: Name: " + property.Name + ", Type: " + property.Type);
    }
    if (props.getPropertySetInfo().hasPropertyByName("Text")) {
      WtMessageHandler.printToLogFile("WtContextMenu: Property: Name: " + props.getPropertyValue("Text"));
    }
    if (props.getPropertySetInfo().hasPropertyByName("CommandURL")) {
      WtMessageHandler.printToLogFile("WtContextMenu: Property: CommandURL: " + props.getPropertyValue("CommandURL"));
    }
  }

}

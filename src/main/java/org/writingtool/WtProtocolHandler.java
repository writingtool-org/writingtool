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

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.ResourceBundle;

import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtVersionInfo;

import com.sun.star.awt.Point;
import com.sun.star.awt.Rectangle;
import com.sun.star.awt.Size;
import com.sun.star.beans.NamedValue;
import com.sun.star.beans.PropertyValue;
import com.sun.star.frame.ControlCommand;
import com.sun.star.frame.DispatchDescriptor;
import com.sun.star.frame.FeatureStateEvent;
import com.sun.star.frame.XDispatch;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XLayoutManager;
import com.sun.star.frame.XStatusListener;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.lib.uno.helper.WeakBase;
import com.sun.star.ui.DockingArea;
import com.sun.star.uno.XComponentContext;
import com.sun.star.util.URL;

/**
 * Class to handle the WritingTool toobar
 * @since 26.4
 * @author Fred Kruse
 */
public class WtProtocolHandler extends WeakBase implements XDispatchProvider, XDispatch, XServiceInfo {

  private static final ResourceBundle MESSAGES = WtOfficeTools.getMessageBundle();

  public static final String WT_PROTOCOL = "org.writingtool.WritingTool:";

  public static final String WT_IGNORE_ONCE = "ignoreOnce";
  public static final String WT_IGNORE_ALL = "ignoreAll";
  public static final String WT_IGNORE_PERMANENT = "ignorePermanent";
  public static final String WT_DEACTIVATE_RULE = "deactivateRule";
  public static final String WT_MORE_INFO = "moreInfo";
  public static final String WT_ACTIVATE_RULES = "activateRules";
  public static final String WT_ACTIVATE_RULE = "activateRule_";
  public static final String WT_REMOTE_HINT = "remoteHint";   
  public static final String WT_RENEW_MARKUPS = "renewMarkups";
  public static final String WT_ADD_TO_DICTIONARY = "addToDictionary_";
  public static final String WT_NEXT_ERROR = "nextError";
  public static final String WT_CHECKDIALOG = "checkDialog";
  public static final String WT_CHECKAGAINDIALOG = "checkAgainDialog";
  public static final String WT_STATISTICAL_ANALYSES = "statisticalAnalyses";   
  public static final String WT_CHANGE_QUOTES = "changeQuotes";   
  public static final String WT_OFF_STATISTICAL_ANALYSES = "offStatisticalAnalyses";   
  public static final String WT_RESET_IGNORE_PERMANENT = "resetIgnorePermanent";   
  public static final String WT_PERMANENT_IGNORE_PARAGRAPH = "permanentIgnoreParagraph";   
  public static final String WT_TOGGLE_BACKGROUND_CHECK = "toggleNoBackgroundCheck";
  public static final String WT_BACKGROUND_CHECK_ON = "backgroundCheckOn";
  public static final String WT_BACKGROUND_CHECK_OFF = "backgroundCheckOff";
  public static final String WT_REFRESH_CHECK = "refreshCheck";
  public static final String WT_ABOUT = "about";
  public static final String WT_LANGUAGETOOL = "lt";
  public static final String WT_OPTIONS = "configure";
  public static final String WT_PROFILES = "profiles";
  public static final String WT_PROFILE = "profileChangeTo_";
  public static final String WT_NONE = "noAction";
  public static final String WT_AI_MARK_ERRORS = "aiAddErrorMarks";
  public static final String WT_AI_CORRECT_ERRORS = "aiCorrectErrors";
  public static final String WT_AI_BETTER_STYLE = "aiBetterStyle";
  public static final String WT_AI_REFORMULATE_TEXT = "aiReformulateText";
  public static final String WT_AI_EXPAND_TEXT = "aiAdvanceText";
  public static final String WT_AI_SYNONYMS_OF_WORD = "aiSynonymsOfWord";
  public static final String WT_AI_TRANSLATE_TEXT = "aiTranslateText";
  public static final String WT_AI_TEXT_TO_SPEECH = "aiTextToSpeech";
  public static final String WT_AI_GENERAL = "aiGeneralCommand";
  public static final String WT_AI_REPLACE_WORD = "aiReplaceWord_";
  public static final String WT_AI_SUMMARY = "aiSummary";
  
  public static final String WT_IGNORE_ONCE_COMMAND = WT_PROTOCOL + WT_IGNORE_ONCE;
  public static final String WT_IGNORE_ALL_COMMAND = WT_PROTOCOL + WT_IGNORE_ALL;
  public static final String WT_IGNORE_PERMANENT_COMMAND = WT_PROTOCOL + WT_IGNORE_PERMANENT;
  public static final String WT_DEACTIVATE_RULE_COMMAND = WT_PROTOCOL + WT_DEACTIVATE_RULE;
  public static final String WT_MORE_INFO_COMMAND = WT_PROTOCOL + WT_MORE_INFO;
  public static final String WT_ACTIVATE_RULES_COMMAND = WT_PROTOCOL + WT_ACTIVATE_RULES;
  public static final String WT_ACTIVATE_RULE_COMMAND = WT_PROTOCOL + WT_ACTIVATE_RULE;
  public static final String WT_REMOTE_HINT_COMMAND = WT_PROTOCOL + WT_REMOTE_HINT;   
  public static final String WT_RENEW_MARKUPS_COMMAND = WT_PROTOCOL + WT_RENEW_MARKUPS;
  public static final String WT_ADD_TO_DICTIONARY_COMMAND = WT_PROTOCOL + WT_ADD_TO_DICTIONARY;
  public static final String WT_NEXT_ERROR_COMMAND = WT_PROTOCOL + WT_NEXT_ERROR;
  public static final String WT_CHECKDIALOG_COMMAND = WT_PROTOCOL + WT_CHECKDIALOG;
  public static final String WT_CHECKAGAINDIALOG_COMMAND = WT_PROTOCOL + WT_CHECKAGAINDIALOG;
  public static final String WT_STATISTICAL_ANALYSES_COMMAND = WT_PROTOCOL + WT_STATISTICAL_ANALYSES;   
  public static final String WT_CHANGE_QUOTES_COMMAND = WT_PROTOCOL + WT_CHANGE_QUOTES;   
  public static final String WT_OFF_STATISTICAL_ANALYSES_COMMAND = WT_PROTOCOL + WT_OFF_STATISTICAL_ANALYSES;   
  public static final String WT_RESET_IGNORE_PERMANENT_COMMAND = WT_PROTOCOL + WT_RESET_IGNORE_PERMANENT;   
  public static final String WT_PERMANENT_IGNORE_PARAGRAPH_COMMAND = WT_PROTOCOL + WT_PERMANENT_IGNORE_PARAGRAPH;   
  public static final String WT_TOGGLE_BACKGROUND_CHECK_COMMAND = WT_PROTOCOL + WT_TOGGLE_BACKGROUND_CHECK;
  public static final String WT_BACKGROUND_CHECK_ON_COMMAND = WT_PROTOCOL + WT_BACKGROUND_CHECK_ON;
  public static final String WT_BACKGROUND_CHECK_OFF_COMMAND = WT_PROTOCOL + WT_BACKGROUND_CHECK_OFF;
  public static final String WT_REFRESH_CHECK_COMMAND = WT_PROTOCOL + WT_REFRESH_CHECK;
  public static final String WT_ABOUT_COMMAND = WT_PROTOCOL + WT_ABOUT;
  public static final String WT_LANGUAGETOOL_COMMAND = WT_PROTOCOL + WT_LANGUAGETOOL;
  public static final String WT_OPTIONS_COMMAND = WT_PROTOCOL + WT_OPTIONS;
  public static final String WT_PROFILES_COMMAND = WT_PROTOCOL + WT_PROFILES;
  public static final String WT_PROFILE_COMMAND = WT_PROTOCOL + WT_PROFILE;
  public static final String WT_NONE_COMMAND = WT_PROTOCOL + WT_NONE;
  public static final String WT_AI_MARK_ERRORS_COMMAND = WT_PROTOCOL + WT_AI_MARK_ERRORS;
  public static final String WT_AI_CORRECT_ERRORS_COMMAND = WT_PROTOCOL + WT_AI_CORRECT_ERRORS;
  public static final String WT_AI_BETTER_STYLE_COMMAND = WT_PROTOCOL + WT_AI_BETTER_STYLE;
  public static final String WT_AI_REFORMULATE_TEXT_COMMAND = WT_PROTOCOL + WT_AI_REFORMULATE_TEXT;
  public static final String WT_AI_EXPAND_TEXT_COMMAND = WT_PROTOCOL + WT_AI_EXPAND_TEXT;
  public static final String WT_AI_SYNONYMS_OF_WORD_COMMAND = WT_PROTOCOL + WT_AI_SYNONYMS_OF_WORD;
  public static final String WT_AI_TRANSLATE_TEXT_COMMAND = WT_PROTOCOL + WT_AI_TRANSLATE_TEXT;
  public static final String WT_AI_TEXT_TO_SPEECH_COMMAND = WT_PROTOCOL + WT_AI_TEXT_TO_SPEECH;
  public static final String WT_AI_GENERAL_COMMAND = WT_PROTOCOL + WT_AI_GENERAL;
  public static final String WT_AI_REPLACE_WORD_COMMAND = WT_PROTOCOL + WT_AI_REPLACE_WORD;
  public static final String WT_AI_SUMMARY_COMMAND = WT_PROTOCOL + WT_AI_SUMMARY;
  
  private static final String[] serviceNames = { "com.sun.star.frame.ProtocolHandler" };
  
  private final Map<String, List<XStatusListener>> xListenersMap = new HashMap<>();
  
  private static boolean isConfigRead = false;
  
  private XComponentContext xContext;
  
  private boolean statAnEnabled = false;
  private boolean backgroundCheckState = true;
  private Map<String, String> deactivatedRulesMap = null;
  private List<String> definedProfiles = null;
  private WtConfiguration conf = null;
  String currentProfile = null;
  

  public WtProtocolHandler(XComponentContext xContext) {
    this.xContext = xContext;
    if (WritingTool.getDocumentsHandler() != null) {
      WritingTool.getDocumentsHandler().setProtocolHandler(this);
    }
  }
  
  /**
   * Interface XDispatch
   * add status listener
   */
  @Override
  public void addStatusListener(XStatusListener statusListener, URL url) {
    try {
      if (statusListener != null && url.Protocol.equals(WT_PROTOCOL)) {
        String path = url.Path;
        if (path.equals(WT_STATISTICAL_ANALYSES) || path.equals(WT_BACKGROUND_CHECK_OFF) || path.equals(WT_BACKGROUND_CHECK_ON) 
            || path.equals(WT_PROFILES) || path.equals(WT_ACTIVATE_RULES)) {
          List<XStatusListener> xListeners;
          if (xListenersMap.containsKey(path)) {
            xListeners = xListenersMap.get(path);
          } else {
            xListeners = new ArrayList<>();
          }
          xListeners.add(statusListener);
  //        for (XStatusListener listener : xListeners) {
  //          WtMessageHandler.printToLogFile("WtProtocolHandler: add StatusListener: listener is " + (listener == null ? "" : "NOT ") + "null");
  //        }
          xListenersMap.put(path, xListeners);
  //        WtMessageHandler.printToLogFile("WtProtocolHandler: add StatusListener: URL.Path: " + path);
          setButtonState();
          if (path.equals(WT_PROFILES)){
            conf = WritingTool.getDocumentsHandler().getLastConfiguration();
//            WtMessageHandler.printToLogFile("WtProtocolHandler: add StatusListener: URL.Path: " + path);
            makeProfileList(statusListener, url);
          } else if (path.equals(WT_PROFILES)){
//            WtMessageHandler.printToLogFile("WtProtocolHandler: add StatusListener: URL.Path: " + path);
            makeActivateRulesList(statusListener, url);
          }
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }

  /**
   * Interface XDispatch
   * remove status listener
   */
  @Override
  public void removeStatusListener(XStatusListener statusListener, URL url) {
    if (statusListener != null && url.Protocol.equals(WT_PROTOCOL)) {
      String path = url.Path;
      if (xListenersMap.containsKey(path)) {
        List<XStatusListener> xListeners = xListenersMap.get(path);
        if (xListeners.contains(statusListener)) {
          xListeners.remove(statusListener);
          xListenersMap.put(path, xListeners);
//          WtMessageHandler.printToLogFile("WtProtocolHandler: remove StatusListener: URL.Path: " + path);
        }
      }
    }
  }

  /**
   * Interface XDispatch
   * dispatch query
   */
  @Override
  public void dispatch(URL url, PropertyValue[] props) {
//    WtMessageHandler.printToLogFile("WtProtocolHandler: dispatch: URL.Protocol: " + url.Protocol);
//    WtMessageHandler.printToLogFile("WtProtocolHandler: dispatch: URL.Path: " + url.Path);
    if (url.Protocol.equals(WT_PROTOCOL)) {
//      WtMessageHandler.printToLogFile("WtProtocolHandler: dispatch: URL: " + url.Complete);
      WritingTool.getDocumentsHandler().setMenuDocId(null);
      if (url.Path.equals(WT_PROFILES)) {
        WritingTool.getDocumentsHandler().trigger(handleProfiles(url, props));
        return;
      } else if (url.Path.equals(WT_ACTIVATE_RULES)) {
        WritingTool.getDocumentsHandler().trigger(handleActivateRules(url, props));
        changeListOfButton(url);
        return;
      }
      WritingTool.getDocumentsHandler().trigger(url.Path);
    }
  }

  /**
   * Interface XDispatchProvider
   * query dispatch
   */
  @Override
  public XDispatch queryDispatch(URL url, String targetFrameName, int searchFlags) {
//    WtMessageHandler.printToLogFile("WtProtocolHandler: queryDispatch: URL: " + url.Complete);
//    WtMessageHandler.printToLogFile("WtProtocolHandler: queryDispatch: URL.Protocol: " + url.Protocol);
    if (url.Protocol.equals(WT_PROTOCOL)) {
//      WtMessageHandler.printToLogFile("WtProtocolHandler: queryDispatch: return this");
      return this;
    }
    return null;
  }

  /**
   * Interface XDispatchProvider
   * query multiple dispatches
   */
  @Override
  public XDispatch[] queryDispatches(DispatchDescriptor[] requests) {
    int nCount = requests.length;
    XDispatch[] lDispatcher = new XDispatch[nCount];
    for (int i = 0; i < nCount; ++i) {
      lDispatcher[i] = queryDispatch( requests[i].FeatureURL, requests[i].FrameName, requests[i].SearchFlags );
    }
    return lDispatcher;
  }

  @Override
  public String getImplementationName() {
    return WtProtocolHandler.class.getName();
  }

  @Override
  public String[] getSupportedServiceNames() {
    return getServiceNames();
  }

  public static String[] getServiceNames() {
    return serviceNames;
}

  @Override
  public boolean supportsService(String service) {
    for (int i = 0; i < serviceNames.length; i++) {
      if (service.equals(serviceNames[i])) {
        return true;
      }
    }
    return false;
  }
  
  public void setButtonState() {
    if (!WritingTool.getDocumentsHandler().isOpenOffice) {
      statAnEnabled = WtOfficeTools.hasStatisticalStyleRules(xContext);
      conf = WritingTool.getDocumentsHandler().getLastConfiguration();
      if (conf != null) {
        backgroundCheckState = conf.noBackgroundCheck();
        definedProfiles = conf.getDefinedProfiles();
        deactivatedRulesMap = WritingTool.getDocumentsHandler().getDisabledRulesMap(null);
      }
      changeButtonStatus();
    }
  }
  
  private void changeButtonStatus() {
    URL url = WtOfficeTools.createUrl(xContext, WT_STATISTICAL_ANALYSES_COMMAND);
    changeStateOfButton(url, statAnEnabled, false);
    url = WtOfficeTools.createUrl(xContext, WT_BACKGROUND_CHECK_OFF_COMMAND);
    changeStateOfButton(url, !backgroundCheckState, false);
    url = WtOfficeTools.createUrl(xContext, WT_BACKGROUND_CHECK_ON_COMMAND);
    changeStateOfButton(url, backgroundCheckState, false);
    url = WtOfficeTools.createUrl(xContext, WT_PROFILES_COMMAND);
    if (definedProfiles == null || definedProfiles.isEmpty()) {
      changeStateOfButton(url, false, false);
    } else {
      changeStateOfButton(url, true, false);
      changeListOfButton(url);
    }
    url = WtOfficeTools.createUrl(xContext, WT_ACTIVATE_RULES_COMMAND);
    if (deactivatedRulesMap == null || deactivatedRulesMap.isEmpty()) {
      changeStateOfButton(url, false, false);
    } else {
      changeStateOfButton(url, true, false);
      changeListOfButton(url);
    }
  }
  
  private void changeStateOfButton(URL url, boolean enabled, boolean state) {
    try {
      String path = url.Path;
//      WtMessageHandler.printToLogFile("WtProtocolHandler: changeStateOfButton: Path: " + path);
      if (xListenersMap.containsKey(path)) {
//        WtMessageHandler.printToLogFile("WtProtocolHandler: changeStateOfButton: contains key: " + path);
        List<XStatusListener> xListeners = xListenersMap.get(path);
//        WtMessageHandler.printToLogFile("WtProtocolHandler: changeStateOfButton: number of listeners: " + xListeners.size());
        if (!xListeners.isEmpty()) {
//          WtMessageHandler.printToLogFile("WtProtocolHandler: changeStateOfButton: number of listeners: " + xListeners.size());
          for (XStatusListener xStaLis : xListeners) {
//            WtMessageHandler.printToLogFile("WtProtocolHandler: changeStateOfButton: set new status!");
            if (xStaLis != null) {
              FeatureStateEvent xEvent = new FeatureStateEvent();
              xEvent.FeatureURL = url;
              xEvent.IsEnabled = enabled;
              xEvent.State = state;
              xStaLis.statusChanged(xEvent);
//              WtMessageHandler.printToLogFile("WtProtocolHandler: changeStateOfButton: new status is set!");
//            } else {
//              WtMessageHandler.printToLogFile("WtProtocolHandler: changeStateOfButton: listener == null");
            }
          }
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }
  
  private void changeListOfButton(URL url) {
    try {
      String path = url.Path;
      if (xListenersMap.containsKey(path)) {
        List<XStatusListener> xListeners = xListenersMap.get(path);
        if (!xListeners.isEmpty()) {
          for (XStatusListener xStaLis : xListeners) {
            if (xStaLis != null) {
              if (url.Path.equals(WT_PROFILES)) {
                makeProfileList(xStaLis, url);
              } else if (url.Path.equals(WT_ACTIVATE_RULES)) {
                makeActivateRulesList(xStaLis, url);
              }
            }
          }
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }
  
  private String WT_TB_FILE_NAME = "toolbar.cfg";
  private String WT_TOOLBAR_URL = "private:resource/toolbar/addon_org.writingtool.WritingTool.toolbar";

  private String WT_TB_ISVISIBLE = "WT_TB_IsVisible";
  private String WT_TB_ISDOCKED = "WT_TB_IsDocked";
  private String WT_TB_ISLOCKED = "WT_TB_IsLocked";
  private String WT_TB_POS = "WT_TB_Pos";
  private String WT_TB_SIZE = "WT_TB_Size";
  private String WT_TB_DOCK_AREA = "WT_TB_DockingArea";

  public void writeCurrentConfiguration() {
    if (!WritingTool.getDocumentsHandler().isOpenOffice) {
      try {
        Properties props = new Properties();
        XLayoutManager layoutManager = WtOfficeTools.getLayoutManager(xContext);
        
        Rectangle rect = layoutManager.getCurrentDockingArea();
  //      String sRect = rect.X + "," + rect.Y + "," + rect.Width + "," + rect.Height;
    
        props.setProperty(WT_TB_ISVISIBLE, Boolean.toString(layoutManager.isElementVisible(WT_TOOLBAR_URL)));
        props.setProperty(WT_TB_ISDOCKED, Boolean.toString(layoutManager.isElementDocked(WT_TOOLBAR_URL)));
        props.setProperty(WT_TB_ISLOCKED, Boolean.toString(layoutManager.isElementLocked(WT_TOOLBAR_URL)));
        Point point = layoutManager.getElementPos(WT_TOOLBAR_URL);
        String sPoint = point.X + "," + point.Y;
        props.setProperty(WT_TB_POS, sPoint);
        Size size = layoutManager.getElementSize(WT_TOOLBAR_URL);
        sPoint = size.Width + "," + size.Height;
        props.setProperty(WT_TB_SIZE, sPoint);
        int dockingArea = 0;
        if (size.Width > size.Height) {
          if (rect.Y < point.Y * size.Height) {
            dockingArea = 1;
          }
        } else {
          if (rect.X >= size.Width) {
            dockingArea = 2;
          } else {
            dockingArea = 3;
          }
        }
        props.setProperty(WT_TB_DOCK_AREA, Integer.toString(dockingArea));
        
  //      props.setProperty("rect", sRect);
    
        File tbConfigFile = new File(WtOfficeTools.getWtConfigDir(xContext), WT_TB_FILE_NAME);
        
        try (FileOutputStream fos = new FileOutputStream(tbConfigFile)) {
          props.store(fos, WtOfficeTools.WT_NAME + " toolbar configuration (" + WtVersionInfo.getWtInformation() + ")");
        } catch (Throwable t) {
          WtMessageHandler.printException(t);
        }
      } catch (Throwable t) {
        WtMessageHandler.printException(t);
      }
    }
  }

  public void readConfiguration() {
    if (!isConfigRead && !WritingTool.getDocumentsHandler().isOpenOffice) {
      isConfigRead = true;
      readConfig();
    }
  }
  
  private DockingArea getDockingArea(int iDockingArea) {
    DockingArea dockingArea;
    if (iDockingArea == 0) {
      dockingArea = DockingArea.DOCKINGAREA_TOP;
    } else if (iDockingArea == 1) {
      dockingArea = DockingArea.DOCKINGAREA_BOTTOM;
    } else if (iDockingArea == 2) {
      dockingArea = DockingArea.DOCKINGAREA_LEFT;
    } else if (iDockingArea == 3) {
      dockingArea = DockingArea.DOCKINGAREA_RIGHT;
    } else {
      dockingArea = DockingArea.DOCKINGAREA_DEFAULT;
    }
    return dockingArea;
  }

  private void readConfig() {
    
    File tbConfigFile = new File(WtOfficeTools.getWtConfigDir(xContext), WT_TB_FILE_NAME);

    try (FileInputStream fis = new FileInputStream(tbConfigFile)) {
      WtMessageHandler.printToLogFile("WtProtocolHandler: readConfig: read started");
      XLayoutManager layoutManager = WtOfficeTools.getLayoutManager(xContext);
      Properties props = new Properties();
      props.load(fis);
      String sProp = (String) props.get(WT_TB_POS);
      Point point = null;
      if (sProp != null) {
        String[] sPoint = sProp.split(",");
        if (sPoint.length == 2) {
          point = new Point(Integer.parseInt(sPoint[0]), Integer.parseInt(sPoint[1]));
          layoutManager.setElementPos(WT_TOOLBAR_URL, point);
        }
      }
      int dockingArea = -1;
      sProp = (String) props.get(WT_TB_DOCK_AREA);
      if (sProp != null) {
        dockingArea = Integer.parseInt(sProp);
      }
      sProp = (String) props.get(WT_TB_ISDOCKED);
      if (sProp != null) {
        if(Boolean.parseBoolean(sProp)) {
          if (point != null && dockingArea >= 0) {
            layoutManager.dockWindow(WT_TOOLBAR_URL, getDockingArea(dockingArea), point);
          }
          layoutManager.lockWindow(WT_TOOLBAR_URL);
        }
      }
      sProp = (String) props.get(WT_TB_ISVISIBLE);
      if (sProp != null) {
        if(Boolean.parseBoolean(sProp)) {
          layoutManager.showElement(WT_TOOLBAR_URL);
        } else {
          layoutManager.hideElement(WT_TOOLBAR_URL);
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
  }

  private void sendCommandTo(XStatusListener statusListener, URL url, String cmd, NamedValue[] args, boolean enabled) throws Throwable {
    FeatureStateEvent aEvent = new FeatureStateEvent();
    aEvent.FeatureURL = url;
    aEvent.Source     = (XDispatch) this;
    aEvent.IsEnabled  = enabled;
    aEvent.Requery    = false;

    ControlCommand ctrlCmd = new ControlCommand();
    ctrlCmd.Command   = cmd;
    ctrlCmd.Arguments = args;

    aEvent.State = ctrlCmd;
    statusListener.statusChanged(aEvent);
  }

  private void makeProfileList(XStatusListener statusListener, URL url) throws Throwable {
    currentProfile = conf.getCurrentProfile();
    List<String> definedProfiles = new ArrayList<>(conf.getDefinedProfiles());
    definedProfiles.sort(null);
    definedProfiles.add(0, MESSAGES.getString("guiUserProfile"));
    if (currentProfile != null && !currentProfile.isEmpty()) {
      definedProfiles.remove(currentProfile);
      definedProfiles.add(0, currentProfile);
    }
    String[] contextMenu = new String[definedProfiles.size()];
    for (int i = 0; i < definedProfiles.size(); i++) {
      contextMenu[i] = definedProfiles.get(i);
    }
    NamedValue[] args = new NamedValue[1];
    // set list
    args[0] = new NamedValue();
    args[0].Name = "List";
    args[0].Value = contextMenu;
    sendCommandTo(statusListener, url, "SetList", args, true);
    // set check to actual profile
    args[0].Name = "Pos";
    args[0].Value = 0;
    sendCommandTo(statusListener, url, "CheckItemPos", args, true);
//    WtMessageHandler.printToLogFile("WtProtocolHandler: add StatusListener: " + contextMenu.length + " profiles set");
  }

  private String handleProfiles(URL url, PropertyValue[] props) {
    for (PropertyValue propVal : props) {
      if (propVal.Name.equals("Text")) {
        String text = new String((String) propVal.Value);
        if (text.equals(MESSAGES.getString("guiUserProfile"))) {
          text = "";
        }
        conf = WritingTool.getDocumentsHandler().getLastConfiguration();
        changeListOfButton(url);
        return WT_PROFILE + text;
      }
    }
    return WT_PROFILES;
  }

  private void makeActivateRulesList(XStatusListener statusListener, URL url) throws Throwable {
    deactivatedRulesMap = WritingTool.getDocumentsHandler().getDisabledRulesMap(null);
    String[] contextMenu = new String[deactivatedRulesMap.keySet().size()];
    int i = 0;
    for (String ruleId : deactivatedRulesMap.keySet()) {
      contextMenu[i] = deactivatedRulesMap.get(ruleId);
      i++;
    }
    NamedValue[] args = new NamedValue[1];
    // set list
    args[0] = new NamedValue();
    args[0].Name = "List";
    args[0].Value = contextMenu;
    sendCommandTo(statusListener, url, "SetList", args, true);
//    WtMessageHandler.printToLogFile("WtProtocolHandler: add StatusListener: " + contextMenu.length + " profiles set");
  }

  private String handleActivateRules(URL url, PropertyValue[] props) {
    for (PropertyValue propVal : props) {
      if (propVal.Name.equals("Text")) {
        String text = new String((String) propVal.Value);
        deactivatedRulesMap = WritingTool.getDocumentsHandler().getDisabledRulesMap(null);
        for (String ruleId : deactivatedRulesMap.keySet()) {
          if (text.equals(deactivatedRulesMap.get(ruleId))) {
            return WT_ACTIVATE_RULE + ruleId;
          }
        }
      }
    }
    return WT_ACTIVATE_RULES;
  }

  
}

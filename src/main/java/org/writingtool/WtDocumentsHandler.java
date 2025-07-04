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

import java.awt.*;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.io.File;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.ResourceBundle;
import java.util.Set;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.UIManager;

import org.apache.commons.lang3.StringUtils;
import org.languagetool.Language;
import org.languagetool.Languages;
import org.languagetool.UserConfig;
import org.languagetool.rules.Rule;
import org.languagetool.tools.Tools;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.WtSingleDocument.RuleDesc;
import org.writingtool.aisupport.WtAiCheckQueue;
import org.writingtool.aisupport.WtAiErrorDetection;
import org.writingtool.aisupport.WtAiParagraphChanging;
import org.writingtool.aisupport.WtAiRemote.AiCommand;
import org.writingtool.aisupport.WtAiTextToSpeech;
import org.writingtool.aisupport.WtAiTranslateDocument;
import org.writingtool.aisupport.WtAiErrorDetection.DetectionType;
import org.writingtool.config.WtConfiguration;
import org.writingtool.dialogs.WtAboutDialog;
import org.writingtool.dialogs.WtAiDialog;
import org.writingtool.dialogs.WtCheckDialog;
import org.writingtool.dialogs.WtConfigurationDialog;
import org.writingtool.dialogs.WtMoreInfoDialog;
import org.writingtool.dialogs.WtStatAnDialog;
import org.writingtool.dialogs.WtCheckDialog.LtCheckDialog;
import org.writingtool.sidebar.WtSidebarContent;
import org.writingtool.config.WtConfigThread;
import org.writingtool.tools.WtGeneralTools;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeDrawTools;
import org.writingtool.tools.WtOfficeSpreadsheetTools;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtViewCursorTools;
import org.writingtool.tools.WtOfficeTools.DocumentType;
import org.writingtool.tools.WtOfficeTools.LoErrorType;
import org.writingtool.tools.WtVersionInfo;

import com.sun.star.beans.PropertyValue;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.Locale;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XEventListener;
import com.sun.star.linguistic2.LinguServiceEvent;
import com.sun.star.linguistic2.LinguServiceEventFlags;
import com.sun.star.linguistic2.ProofreadingResult;
import com.sun.star.linguistic2.XLinguServiceEventListener;
import com.sun.star.linguistic2.XProofreader;
import com.sun.star.text.XTextDocument;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


/**
 * Class to handle multiple LO documents for checking
 * @since 1.0
 * @author Fred Kruse, Marcin Miłkowski
 */
public class WtDocumentsHandler {

  // LibreOffice (since 4.2.0) special tag for locale with variant 
  // e.g. language ="qlt" country="ES" variant="ca-ES-valencia":
  private static final String LIBREOFFICE_SPECIAL_LANGUAGE_TAG = "qlt";
  
  private static final ResourceBundle messages = WtOfficeTools.getMessageBundle();

  private static final int HEAP_CHECK_INTERVAL = 1000;

  private final List<XLinguServiceEventListener> xEventListeners;
  private boolean docReset = false;

  private static boolean debugMode = false;   //  should be false except for testing
  private static boolean debugModeTm = false;   //  should be false except for testing

  public final boolean isOpenOffice;
  
  private WtLanguageTool lt = null;
  private Language docLanguage = null;
  private Language fixedLanguage = null;
  private Language langForShortName;
  private Locale locale;                            //  locale for grammar check
  private final XEventListener xEventListener;
  private final XProofreader xProofreader;
  private static File configDir;
  private static String configFile;
  private static WtConfiguration config = null;
  private WtLinguisticServices linguServices = null;
  private WtSidebarContent sidebarContent;
  private static Map<String, Set<String>> disabledRulesUI; //  Rules disabled by context menu or spell dialog
  private final List<Rule> extraRemoteRules;        //  store of rules supported by remote server but not locally
  private LtCheckDialog ltDialog = null;            //  WT spelling and grammar check dialog
  private WtConfigurationDialog cfgDialog = null;   //  configuration dialog (show only one configuration panel)
  private static WtAboutDialog aboutDialog = null;  //  about dialog (show only one about panel)
  private static WtMoreInfoDialog infoDialog = null;//  more info about a rule dialog (show only one info panel)
  private boolean dialogIsRunning = false;          //  The dialog was started
  private WaitDialogThread waitDialog = null;

  
  private static XComponentContext xContext;               //  The context of the document
  private final List<WtSingleDocument> documents;   //  The List of LO documents to be checked
  private final List<String> disposedIds; //  The List of IDs of disposed documents
  private boolean recheck = true;                   //  if true: recheck the whole document at next iteration
  private int docNum;                               //  number of the current document
  
  private int numSinceHeapTest = 0;                 //  number of checks since last heap test
  private boolean heapLimitReached = false;         //  heap limit is reached

  private boolean noBackgroundCheck = false;        //  is LT switched off by config
  private boolean useQueue = false;                  //  will be overwritten by config
  private boolean noLtSpeller = false;              //  true if LT spell checker can't be used

  private String menuDocId = null;                    //  Id of document at which context menu was called 
  private WtTextLevelCheckQueue textLevelQueue = null;// Queue to check text level rules
  private WtAiCheckQueue aiQueue = null;              // Queue to check by AI support
  private ShapeChangeCheck shapeChangeCheck = null;   // Thread for test changes in shape texts
  private boolean doShapeCheck = false;               // do the test for changes in shape texts
  
  private boolean useOrginalCheckDialog = false;    // use original spell and grammar dialog (LT check dialog does not work for OO)
  private boolean checkImpressDocument = false;     //  the document to check is Impress
//  private final HandleLtDictionary handleDictionary;
  private boolean isNotTextDocument = false;
  private int heapCheckInterval = HEAP_CHECK_INTERVAL;
  private boolean testMode = false;
  private static int javaLookAndFeelSet = -1;
  private boolean isHelperDisposed = false;
  private boolean statAnDialogRunning = false;

  
  WtDocumentsHandler(XComponentContext xContext, XProofreader xProofreader, XEventListener xEventListener) {
    WtDocumentsHandler.xContext = xContext;
    this.xEventListener = xEventListener;
    this.xProofreader = xProofreader;
    xEventListeners = new ArrayList<>();
    WtVersionInfo.init(xContext);
    WtOfficeTools.renameOldLtFiles();     // This has to be deleted in later versions 
    if (WtVersionInfo.ooName == null || WtVersionInfo.ooName.equals("OpenOffice")) {
      isOpenOffice = true;
      useOrginalCheckDialog = true;
      configFile = WtOfficeTools.OOO_CONFIG_FILE;
    } else {
      isOpenOffice = false;
      configFile = WtOfficeTools.CONFIG_FILE;
    }
    configDir = WtOfficeTools.getWtConfigDir(xContext);
    WtMessageHandler.init(xContext, false);
    documents = new ArrayList<>();
    disposedIds = new ArrayList<>();
    disabledRulesUI = new HashMap<>();
    extraRemoteRules = new ArrayList<>();
    if (WtVersionInfo.osArch == null || WtVersionInfo.osArch.equals("x86")
        || !WtSpellChecker.runLTSpellChecker(xContext)) {
      noLtSpeller = true;
    }
//    handleDictionary = new HandleLtDictionary();
//    handleDictionary.start();
    LtHelper ltHelper = new LtHelper();
    ltHelper.start();
  }
  
  /**
   * Runs the grammar checker on paragraph text.
   *
   * @param docID document ID
   * @param paraText paragraph text
   * @param locale Locale the text Locale
   * @param startOfSentencePos start of sentence position
   * @param nSuggestedBehindEndOfSentencePosition end of sentence position
   * @return ProofreadingResult containing the results of the check.
   */
  public final ProofreadingResult doProofreading(String docID,
      String paraText, Locale locale, int startOfSentencePos,
      int nSuggestedBehindEndOfSentencePosition,
      PropertyValue[] propertyValues) {
    ProofreadingResult paRes = new ProofreadingResult();
    paRes.nStartOfSentencePosition = startOfSentencePos;
    paRes.nStartOfNextSentencePosition = nSuggestedBehindEndOfSentencePosition;
    paRes.nBehindEndOfSentencePosition = paRes.nStartOfNextSentencePosition;
    paRes.xProofreader = xProofreader;
    paRes.aLocale = locale;
    paRes.aDocumentIdentifier = docID;
    paRes.aText = paraText;
    paRes.aProperties = propertyValues;
    try {
      paRes = getCheckResults(paraText, locale, paRes, propertyValues, docReset);
      docReset = false;
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return paRes;
  }

  /**
   * distribute the check request to the concerned document
   */
  ProofreadingResult getCheckResults(String paraText, Locale locale, ProofreadingResult paRes, 
      PropertyValue[] propertyValues, boolean docReset) throws Throwable {
    if (!hasLocale(locale)) {
      docLanguage = null;
      WtMessageHandler.printToLogFile("MultiDocumentsHandler: getCheckResults: Sorry, don't have locale: " + WtOfficeTools.localeToString(locale));
      return paRes;
    }
//    LinguisticServices.setLtAsSpellService(xContext, true);
    if (!isBackgroundCheckOff()) {
      boolean isSameLanguage = true;
      if (fixedLanguage == null || langForShortName == null) {
        langForShortName = getLanguage(locale);
        isSameLanguage = langForShortName.equals(docLanguage) && lt != null;
      }
      if (!isSameLanguage || recheck || checkImpressDocument) {
        boolean initDocs = (lt == null || recheck || checkImpressDocument);
        if (debugMode && initDocs) {
          WtMessageHandler.printToLogFile("initDocs: lt " + (lt == null ? "=" : "!") + "= null, recheck: " + recheck 
              + ", Impress: " + checkImpressDocument);
        }
        checkImpressDocument = false;
        if (!isSameLanguage) {
          docLanguage = langForShortName;
          this.locale = locale;
          extraRemoteRules.clear();
        }
        lt = initLanguageTool();
        if (javaLookAndFeelSet < 0) {
          setJavaLookAndFeel();
        }
        if (initDocs) {
          initDocuments(true);
        }
      } else {
        if (textLevelQueue == null && useQueue) {
          textLevelQueue = new WtTextLevelCheckQueue(this);
        }
        if (aiQueue == null && config.getNumParasToCheck() != 0 && config.useAiSupport() && config.aiAutoCorrect()) {
          aiQueue = new WtAiCheckQueue(this);
        }
      }
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("MultiDocumentsHandler: getCheckResults: Start getNumDoc!");
    }
    docNum = getNumDoc(paRes.aDocumentIdentifier, propertyValues);
    if (isBackgroundCheckOff()) {
      return paRes;
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("MultiDocumentsHandler: getCheckResults: Start testHeapSpace!");
    }
    if (!isBackgroundCheckOff()) {
      testHeapSpace();
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("MultiDocumentsHandler: getCheckResults: Start getCheckResults at single document: " + paraText);
    }
//    handleLtDictionary(paraText);
    paRes = documents.get(docNum).getCheckResults(paraText, locale, paRes, propertyValues, docReset, lt, LoErrorType.GRAMMAR);
    if (!isBackgroundCheckOff() && lt.doReset()) {
      // langTool.doReset() == true: if server connection is broken ==> switch to internal check
      WtMessageHandler.showMessage(messages.getString("loRemoteSwitchToLocal"));
      config.setRemoteCheck(false);
      try {
        config.saveConfiguration(docLanguage);
      } catch (IOException e) {
        WtMessageHandler.showError(e);
      }
      resetDocument();
    }
    if (debugMode) {
      WtMessageHandler.printToLogFile("MultiDocumentsHandler: getCheckResults: return to LO/OO!");
    }
    return paRes;
  }

  /**
   *  Get the current used document
   */
  public WtSingleDocument getCurrentDocument() {
    try {
      XComponent xComponent = WtOfficeTools.getCurrentComponent(xContext);
      isNotTextDocument = false;
      if (xComponent != null) {
        for (WtSingleDocument document : documents) {
          if (xComponent.equals(document.getXComponent())) {
            return document;
          }
        }
        XTextDocument curDoc = UnoRuntime.queryInterface(XTextDocument.class, xComponent);
        if (curDoc == null) {
          String prefix = null;
          if (WtOfficeDrawTools.isImpressDocument(xComponent)) {
            prefix = "I";
            checkImpressDocument = true;
          } else if (WtOfficeSpreadsheetTools.isSpreadsheetDocument(xComponent)) {
            prefix = "C";
          }
          if (prefix != null) {
            String docID = createOtherDocId(prefix);
            try {
              xComponent.addEventListener(xEventListener);
            } catch (Throwable t1) {
              WtMessageHandler.printToLogFile("MultiDocumentsHandler: getCurrentDocument: Error: Document (ID: " + docID + ") has no XComponent -> Internal space will not be deleted when document disposes");
              xComponent = null;
            }
            if (config == null) {
              if (docLanguage == null) {
                Locale loc;
                if (prefix.equals("I")) {
                  loc = WtOfficeDrawTools.getDocumentLocale(xComponent);
                } else {
                  loc = WtOfficeSpreadsheetTools.getDocumentLocale(xComponent);
                }
                docLanguage = getLanguage(loc);
                if (docLanguage == null) {
                  //  language is not supported --> use default for configuration
                  loc = new Locale("en","US","");
                  docLanguage = getLanguage(loc);
                }
              }
              config = getConfiguration();
            }
            WtSingleDocument newDocument = new WtSingleDocument(xContext, config, docID, xComponent, this, null);
            documents.add(newDocument);
            WtMessageHandler.printToLogFile("Document " + (documents.size() - 1) + " created; docID = " + docID);
            return newDocument;
          }
          WtMessageHandler.printToLogFile("MultiDocumentsHandler: getCurrentDocument: Is document, but not a text document!");
          isNotTextDocument = true;
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return null;
  }
  
  /**
   * create new Impress document id
   */
  private String createOtherDocId(String prefix) {
    String docID;
    if (documents.size() == 0) {
      return prefix + "1";
    }
    for (int n = 1; n < documents.size() + 1; n++) {
      docID = prefix + n;
      boolean isValid = true;
      for (WtSingleDocument document : documents) {
        if (docID.equals(document.getDocID())) {
          isValid = false;
          break;
        }
      }
      if (isValid) {
        return docID;
      }
    }
    return prefix + (documents.size() + 1);
  }
  
  /**
   * return true, if a document was found but is not a text document
   */
  public boolean isNotTextDocument() {
    return isNotTextDocument;
  }
  
  /**
   * return true, if the LT spell checker is not be used
   */
  public boolean noLtSpeller() {
    return this.noLtSpeller;
  }
  
  /**
   * return true, if the document to check is an Impress document
   */
  public boolean isCheckImpressDocument() {
    return checkImpressDocument;
  }
  
  /**
   * set the checkImpressDocument flag
   */
  public void setCheckImpressDocument(boolean checkImpressDocument) {
    this.checkImpressDocument = checkImpressDocument;
  }
  
  /**
   * set sidebar content
   */
  public void setSidebarContent(WtSidebarContent sidebarContent) {
    this.sidebarContent = sidebarContent;
  }
  
  /**
   * get sidebar content
   */
  public WtSidebarContent getSidebarContent() {
    return sidebarContent;
  }
  
  /**
   * set sidebar content
   */
  public void setTextToSidebarBox(XComponent xComponent) {
    if (sidebarContent != null) {
      sidebarContent.setCursorTextToBox(xComponent);
    }
  }
  
  /**
   *  Set all documents to be checked again
   */
  void setRecheck() {
    recheck = true;
  }
  
  /**
   *  Set XComponentContext
   */
  void setComponentContext(XComponentContext xContxt) {
    if (xContext != null && !xContext.equals(xContxt)) {
      setRecheck();
    }
    xContext = xContxt;
  }
  
  /**
   *  Get XComponentContext
   */
  public static XComponentContext getComponentContext() {
    return xContext;
  }
  
  /**
   *  Set pointer to configuration dialog
   */
  public void setConfigurationDialog(WtConfigurationDialog dialog) {
    cfgDialog = dialog;
  }
  
  /**
   *  Set pointer to LT spell and grammar check dialog
   */
  public void setLtDialog(LtCheckDialog dialog) {
    ltDialog = dialog;
  }
  
  /**
   *  Set Information LT spell and grammar check dialog was started
   */
  public void setLtDialogIsRunning(boolean running) {
    this.dialogIsRunning = running;
  }
  
  /**
   *  Set Information LT spell and grammar check dialog was started
   */
  public void setConfigFileName(String name) {
    configFile = name;
  }
  
  /**
   *  Set dialog for statisical analysis running
   */
  public void setStatAnDialogRunning(boolean running) {
    statAnDialogRunning = running;
  }
  
  /**
   *  use analyzed sentences cache
   */
  public boolean useAnalyzedSentencesCache() {
    return !config.doRemoteCheck() || statAnDialogRunning;
  }
  
  /**
   *  close configuration dialog
   * @throws Throwable 
   */
  private void closeDialogs() throws Throwable {
    if (ltDialog != null) {
      ltDialog.closeDialog();
    } 
    if (cfgDialog != null) {
      cfgDialog.close();
      cfgDialog = null;
    }
    WtAiDialog aiDialog = WtAiParagraphChanging.getAiDialog();
    if (aiDialog != null) {
      aiDialog.closeDialog();
    } 
  }
  
  /**
   *  Set a document as closed
   */
  private void setContextOfClosedDoc(XComponent xComponent) {
    boolean found = false;
    try {
      for (WtSingleDocument document : documents) {
        if (xComponent.equals(document.getXComponent())) {
          found = true;
          document.dispose(true);
          disposedIds.add(document.getDocID());
          if (documents.size() < 2) {
 /*
            if (textLevelQueue != null) {
              textLevelQueue.setStop();
              textLevelQueue = null;
            }
            if (aiQueue != null) {
              aiQueue.setStop();
              aiQueue = null;
            }
*/            
            isHelperDisposed = true;
          }
//          document.removeDokumentListener(xComponent);
//          document.setXComponent(xContext, null);
          if (document.getDocumentCache().hasNoContent(false)) {
            //  The delay seems to be necessary as workaround for a GDK bug (Linux) to stabilizes
            //  the load of a document from an empty document 
            WtMessageHandler.printToLogFile("Disposing document has no content: Wait for 1000 milliseconds");
            try {
              Thread.sleep(1000);
            } catch (InterruptedException e) {
              WtMessageHandler.printException(e);
            }
          }
        }
      }
      if (!found) {
        WtMessageHandler.printToLogFile("MultiDocumentsHandler: setContextOfClosedDoc: Error: Disposed Document not found - Cache not deleted");
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }
  
  /**
   *  stop all queues before unload LO
   */
  public void prepareUnload(String docId) {
    try {
      if (debugMode) {
        WtMessageHandler.printToLogFile("MultiDocumentsHandler: prepareUnload for docId: " + docId);
      }
      if (documents.size() < 2) {
        isHelperDisposed = true;
        if (textLevelQueue != null) {
          textLevelQueue.setStop();
          textLevelQueue = null;
        }
        if (aiQueue != null) {
          aiQueue.setStop();
          aiQueue = null;
        }
        if (shapeChangeCheck != null) {
          shapeChangeCheck.stopLoop();
          shapeChangeCheck = null;
        }
      } else if (documents.size() > 1) {
        if (textLevelQueue != null) {
          textLevelQueue.interruptCheck(docId, false);
        }
        if (aiQueue != null) {
          aiQueue.interruptCheck(docId, false);
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }
  
  /**
   *  Add a rule to disabled rules by context menu or spell dialog
   */
  void addDisabledRule(String langCode, String ruleId) {
    if (disabledRulesUI.containsKey(langCode)) {
      disabledRulesUI.get(langCode).add(ruleId);
    } else {
      Set<String >rulesIds = new HashSet<>();
      rulesIds.add(ruleId);
      disabledRulesUI.put(langCode, rulesIds);
    }
  }
  
  /**
   *  Remove a rule from disabled rules by spell dialog
   */
  public void removeDisabledRule(String langCode, String ruleId) {
    if (disabledRulesUI.containsKey(langCode)) {
      Set<String >rulesIds = disabledRulesUI.get(langCode);
      rulesIds.remove(ruleId);
      if (rulesIds.isEmpty()) {
        disabledRulesUI.remove(langCode);
      } else {
        disabledRulesUI.put(langCode, rulesIds);
      }
    }
  }
  
  /**
   *  remove all disabled rules by context menu or spell dialog
   */
  void resetDisabledRules() {
    disabledRulesUI = new HashMap<>();
  }
  
  /**
   *  get disabled rules for a language code by context menu or spell dialog
   */
  public static Set<String> getDisabledRules(String langCode) {
    if (langCode == null || !disabledRulesUI.containsKey(langCode)) {
      return new HashSet<String>();
    }
    return disabledRulesUI.get(langCode);
  }
  
  /**
   *  get all disabled rules
   */
  Map<String, Set<String>> getAllDisabledRules() {
    return disabledRulesUI;
  }
  
  /**
   *  get all disabled rules
   */
  void setAllDisabledRules(Map<String, Set<String>> disabledRulesUI) {
    WtDocumentsHandler.disabledRulesUI = disabledRulesUI;
  }
  
  /**
   *  get all disabled rules by context menu or spell dialog
   */
  public Map<String, String> getDisabledRulesMap(String langCode) {
    if (langCode == null) {
      langCode = WtOfficeTools.localeToString(locale);
    }
    Map<String, String> disabledRulesMap = new HashMap<>();
    if (langCode != null && lt != null && config != null) {
      List<Rule> allRules = lt.getAllRules();
      List<String> disabledRules = new ArrayList<String>(getDisabledRules(langCode));
      for (int i = disabledRules.size() - 1; i >= 0; i--) {
        String disabledRule = disabledRules.get(i);
        String ruleDesc = null;
        for (Rule rule : allRules) {
          if (disabledRule.equals(rule.getId())) {
            if (!rule.isDefaultOff() || rule.isOfficeDefaultOn()) {
              ruleDesc = rule.getDescription();
            } else {
              removeDisabledRule(langCode, disabledRule);
            }
            break;
          }
        }
        if (ruleDesc != null) {
          disabledRulesMap.put(disabledRule, ruleDesc);
        }
      }
      disabledRules = new ArrayList<String>(config.getDisabledRuleIds());
      for (int i = disabledRules.size() - 1; i >= 0; i--) {
        String disabledRule = disabledRules.get(i);
        String ruleDesc = null;
        for (Rule rule : allRules) {
          if (disabledRule.equals(rule.getId())) {
            if (!rule.isDefaultOff() || rule.isOfficeDefaultOn()) {
              ruleDesc = rule.getDescription();
            } else {
              config.removeDisabledRuleId(disabledRule);
            }
            break;
          }
        }
        if (ruleDesc != null) {
          disabledRulesMap.put(disabledRule, ruleDesc);
        }
      }
    }
    return disabledRulesMap;
  }
  
  /**
   *  set disabled rules by context menu or spell dialog
   */
  public void setDisabledRules(String langCode, Set<String> ruleIds) {
    disabledRulesUI.put(langCode, new HashSet<>(ruleIds));
  }
  
  /**
   *  get LanguageTool
   */
  public WtLanguageTool getLanguageTool() {
    if (lt == null) {
      if (docLanguage == null) {
        docLanguage = getCurrentLanguage();
      }
      lt = initLanguageTool();
    }
    return lt;
  }
  
  /**
   *  get Configuration
   */
  public WtConfiguration getConfiguration() {
    try {
      if (config == null || recheck) {
        if (docLanguage == null) {
          docLanguage = getCurrentLanguage();
        }
        initLanguageTool();
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return config;
  }
  
  /**
   *  get Configuration for language
   *  @throws IOException 
   */
  public static WtConfiguration getConfiguration(Language lang) throws IOException {
    return new WtConfiguration(configDir, configFile, lang, true);
  }
  
  private void disableLTSpellChecker(XComponentContext xContext, Language lang) {
    try {
      config.setUseLtSpellChecker(false);
      config.saveConfiguration(lang);
    } catch (IOException e) {
      WtMessageHandler.printToLogFile("Can't read configuration: LT spell checker not used!");
    }
  }

  /**
   *  get LinguisticServices
   */
  public WtLinguisticServices getLinguisticServices() {
    if (linguServices == null) {
      linguServices = new WtLinguisticServices(xContext);
      WtMessageHandler.printToLogFile("MultiDocumentsHandler: getLinguisticServices: linguServices set: is " 
            + (linguServices == null ? "" : "NOT ") + "null");
      if (noLtSpeller) {
        Tools.setLinguisticServices(linguServices);
        WtMessageHandler.printToLogFile("MultiDocumentsHandler: getLinguisticServices: linguServices set to tools");
      }
    }
    return linguServices;
  }
  
  /**
   * Allow xContext == null for test cases
   */
  void setTestMode(boolean mode) {
    testMode = mode;
    WtMessageHandler.setTestMode(mode);
    if (mode) {
      configFile = "dummy_xxxx.cfg";
      File dummy = new File(configDir, configFile);
      if (dummy.exists()) {
        dummy.delete();
      }
    }
  }

  /**
   * proofs if test cases
   */
  public boolean isTestMode() {
    return testMode;
  }

  /**
   * Checks the language under the cursor. Used for opening the configuration dialog.
   * @return the language under the visible cursor
   * if null or not supported returns the most used language of the document
   * if there is no supported language use en-US as default
   */
  public Language getCurrentLanguage() {
    Locale locale = WtOfficeTools.getCursorLocale(xContext);
    if (locale == null || locale.Language.equals(WtOfficeTools.IGNORE_LANGUAGE) || !hasLocale(locale)) {
      WtSingleDocument document = getCurrentDocument();
      if (document != null) {
        Language lang = document.getLanguage();
        if (lang != null) {
          return lang;
        }
      }
      locale = new Locale("en","US","");
    }
    return getLanguage(locale);
  }
  
  /**
   * @return true if LT supports the language of a given locale
   * @param locale The Locale to check
   */
  public final static boolean hasLocale(Locale locale) {
    try {
      for (Language element : Languages.get()) {
        if (locale.Language.equalsIgnoreCase(LIBREOFFICE_SPECIAL_LANGUAGE_TAG)
            && element.getShortCodeWithCountryAndVariant().equals(locale.Variant)) {
          return true;
        }
        if (element.getShortCode().equals(locale.Language)) {
          return true;
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return false;
  }
  
  /**
   *  Set configuration Values for all documents
   */
  private void setConfigValues(WtConfiguration config, WtLanguageTool lt) throws Throwable {
    WtDocumentsHandler.config = config;
    this.lt = lt;
    if (textLevelQueue != null && (heapLimitReached || config.getNumParasToCheck() == 0 || !config.useTextLevelQueue())) {
      textLevelQueue.setStop();
      textLevelQueue = null;
    }
    if (aiQueue != null && (config.getNumParasToCheck() == 0 || !config.useAiSupport() || !config.aiAutoCorrect())) {
      aiQueue.setStop();
      aiQueue = null;
    }
    useQueue = !isBackgroundCheckOff() && !heapLimitReached && !testMode && config.getNumParasToCheck() == 0 ? false : config.useTextLevelQueue();
    for (WtSingleDocument document : documents) {
      if (!document.isDisposed()) {
        document.setConfigValues(config);
      }
    }
  }

  /**
   * Get language from locale
   */
  public static Language getLanguage(Locale locale) {
    try {
      if (locale.Language.equals(WtOfficeTools.IGNORE_LANGUAGE)) {
        return null;
      }
      if (locale.Language.equalsIgnoreCase(LIBREOFFICE_SPECIAL_LANGUAGE_TAG)) {
        return Languages.getLanguageForShortCode(locale.Variant);
      } else {
        return Languages.getLanguageForShortCode(locale.Language + "-" + locale.Country);
      }
    } catch (Throwable e) {
      try {
        return Languages.getLanguageForShortCode(locale.Language);
      } catch (Throwable t) {
        return null;
      }
    }
  }

  /**
   * Get or Create a Number from docID
   * Return -1 if failed
   */
  private int getNumDoc(String docID, PropertyValue[] propertyValues) throws Throwable {
    for (int i = 0; i < documents.size(); i++) {
      if (documents.get(i).getDocID().equals(docID)) {  //  document exist
        if (!testMode && documents.get(i).getXComponent() == null) {
          XComponent xComponent = WtOfficeTools.getCurrentComponent(xContext);
          if (xComponent == null) {
            WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: Error: Document (ID: " + docID + ") has no XComponent -> Internal space will not be deleted when document disposes");
          } else {
            try {
              xComponent.addEventListener(xEventListener);
            } catch (Throwable t) {
              WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: Error: Document (ID: " + docID + ") has no XComponent -> Internal space will not be deleted when document disposes");
              xComponent = null;
            }
            if (xComponent != null) {
              documents.get(i).setXComponent(xContext, xComponent);
              WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: Fixed: XComponent set for Document (ID: " + docID + ")");
            }
          }
        }
        for (String id : disposedIds) {
          int n = removeDoc(id);
          if (n >= 0 && n < i) {
            i--;
          }
        }
        disposedIds.clear();
        return i;
      }
    }
    //  Add new document
    XComponent xComponent = null;
    if (!testMode) {              //  xComponent == null for test cases 
      xComponent = WtOfficeTools.getCurrentComponent(xContext);
      if (xComponent == null) {
        WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: Error: Document (ID: " + docID + ") has no XComponent -> Internal space will not be deleted when document disposes");
      } else {
        for (int i = 0; i < documents.size(); i++) {
          //  work around to compensate a bug at LO
          if (xComponent.equals(documents.get(i).getXComponent())) {
            WtMessageHandler.printToLogFile("Different Doc IDs, but same xComponents!");
            String oldDocId = documents.get(i).getDocID();
            documents.get(i).setDocID(docID);
            documents.get(i).setLanguage(docLanguage);
            WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: Document ID corrected: old: " + oldDocId + ", new: " + docID);
            if (useQueue && textLevelQueue != null) {
              WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: Interrupt text level queue for old document ID: " + oldDocId);
              textLevelQueue.interruptCheck(oldDocId, true);
              WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: Interrupt done");
            }
            if (config.useAiSupport() && config.aiAutoCorrect() && aiQueue != null) {
              WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: Interrupt AI queue for old document ID: " + oldDocId);
              aiQueue.interruptCheck(oldDocId, true);
              WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: AI Interrupt done");
            }
            if (documents.get(i).isDisposed()) {
              documents.get(i).dispose(false);;
            }
            return i;
          }
        }
        try {
          xComponent.addEventListener(xEventListener);
        } catch (Throwable e) {
          WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: Error: Document (ID: " + docID + ") has no XComponent -> Internal space will not be deleted when document disposes");
          xComponent = null;
        }
      }
    }
    WtSingleDocument newDocument = new WtSingleDocument(xContext, config, docID, xComponent, this, docLanguage);
    documents.add(newDocument);
    for (String id : disposedIds) {
      removeDoc(id);
    }
    WtMessageHandler.printToLogFile("MultiDocumentsHandler: getNumDoc: Document " + (documents.size() - 1) + " created; docID = " + docID);
    return documents.size() - 1;
  }

  /**
   * Delete a document number and all internal space
   */
  private int removeDoc(String docID) {
    try {
      for (int i = documents.size() - 1; i >= 0; i--) {
        if (!docID.equals(documents.get(i).getDocID())) {
          if (documents.get(i).isDisposed()) {
            if (useQueue && textLevelQueue != null) {
              WtMessageHandler.printToLogFile("MultiDocumentsHandler: removeDoc: Interrupt text level queue for document " + documents.get(i).getDocID());
              textLevelQueue.interruptCheck(documents.get(i).getDocID(), true);
              WtMessageHandler.printToLogFile("MultiDocumentsHandler: removeDoc: Interrupt done");
            }
            if (aiQueue != null) {
              WtMessageHandler.printToLogFile("MultiDocumentsHandler: removeDoc: Interrupt ai queue for document " + documents.get(i).getDocID());
              aiQueue.interruptCheck(documents.get(i).getDocID(), true);
              WtMessageHandler.printToLogFile("MultiDocumentsHandler: removeDoc: AI Interrupt done");
            }
            WtMessageHandler.printToLogFile("Disposed document " + documents.get(i).getDocID() + " removed");
            documents.remove(i);
            return (i);
          }
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return (-1);
  }
  
  /**
   * Delete the menu listener of a document
   */
  public void removeMenuListener (XComponent xComponent) throws Throwable {
    if (xComponent != null) {
      for (int i = 0; i < documents.size(); i++) {
        XComponent docComponent = documents.get(i).getXComponent();
        if (docComponent != null && xComponent.equals(docComponent)) { //  disposed document found
          documents.get(i).getLtMenu().removeListener();
          if (debugMode) {
            WtMessageHandler.printToLogFile("MultiDocumentsHandler: removeMenuListener: Menu listener of document " + documents.get(i).getDocID() + " removed");
          }
          break;
        }
      }
    }
  }

  /**
   * Initialize LanguageTool
   */
  public WtLanguageTool initLanguageTool() {
    return initLanguageTool(null);
  }

  public WtLanguageTool initLanguageTool(Language currentLanguage) {
    WtLanguageTool lt = null;
    try {
      config = getConfiguration(currentLanguage == null ? docLanguage : currentLanguage);
      if (this.lt == null) {
        WtOfficeTools.setLogLevel(config.getlogLevel());
        debugMode = WtOfficeTools.DEBUG_MODE_MD;
        debugModeTm = WtOfficeTools.DEBUG_MODE_TM;
        if (!noLtSpeller && !WtSpellChecker.isEnoughHeap()) {
          noLtSpeller = true;
          disableLTSpellChecker(xContext, docLanguage);
          WtMessageHandler.showMessage(messages.getString("guiSpellCheckerWarning"));
        }
      }
      long startTime = 0;
      if (debugModeTm) {
        startTime = System.currentTimeMillis();
      }
      noBackgroundCheck = config.noBackgroundCheck();
      if (linguServices == null) {
        linguServices = getLinguisticServices();
      }
      linguServices.setNoSynonymsAsSuggestions(config.noSynonymsAsSuggestions() || testMode);
      if (currentLanguage == null) {
        fixedLanguage = config.getDefaultLanguage();
        if (fixedLanguage != null) {
          docLanguage = fixedLanguage;
        }
        currentLanguage = docLanguage;
      }
      lt = new WtLanguageTool(currentLanguage, config.getMotherTongue(),
          new UserConfig(config.getConfigurableValues(), linguServices), config, extraRemoteRules, 
          noLtSpeller, checkImpressDocument, testMode);
      config.initStyleCategories(lt.getAllRules());
      recheck = false;
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
        if (runTime > WtOfficeTools.TIME_TOLERANCE) {
          WtMessageHandler.printToLogFile("Time to init Language Tool: " + runTime);
        }
      }
      return lt;
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return lt;
  }

  /**
   * Initialize single documents, prepare text level rules and start queue
   */
  public void initDocuments(boolean resetCache) throws Throwable {
    long startTime = 0;
    if (debugModeTm) {
      startTime = System.currentTimeMillis();
    }
    setConfigValues(config, lt);
    if (useQueue && !isBackgroundCheckOff()) {
      if (textLevelQueue == null) {
        textLevelQueue = new WtTextLevelCheckQueue(this);
      } else {
        textLevelQueue.setReset();
      }
    }
    if (config.useAiSupport() && config.aiAutoCorrect() && !isBackgroundCheckOff()) {
      if (aiQueue == null) {
        aiQueue = new WtAiCheckQueue(this);
      } else {
        aiQueue.setReset();
      }
    } else if (aiQueue != null) {
      aiQueue.setStop();
      aiQueue = null;
    }
    if (resetCache) {
      resetResultCaches(true);
    }
    if (debugModeTm) {
      long runTime = System.currentTimeMillis() - startTime;
      if (runTime > WtOfficeTools.TIME_TOLERANCE) {
        WtMessageHandler.printToLogFile("Time to init Documents: " + runTime);
      }
    }
  }
  
  /**
   * Reset ignored matches
   */
  void resetIgnoredMatches() {
    for (WtSingleDocument document : documents) {
      document.resetIgnoreOnce();
    }
  }

  /**
   * Reset result caches
   */
  void resetResultCaches(boolean withSingleParagraph) {
    for (WtSingleDocument document : documents) {
      document.resetResultCache(withSingleParagraph);
    }
  }

  /**
   * Reset document caches
   */
  public void resetDocumentCaches() {
    for (WtSingleDocument document : documents) {
      document.resetDocumentCache();
    }
  }

  /**
   * Is AI used?
   */
  public boolean useAi() {
    if (config == null) {
      config = getConfiguration();
    }
    return config.useAiSupport() && config.aiAutoCorrect();
  }
  
  /**
   * Is AI used?
   */
  public boolean useAiSuggestion() {
    if (config == null) {
      config = getConfiguration();
    }
    return config.aiAutoSuggestion();
  }
  
  /**
   * Get current locale language
   */
  public Locale getLocale() {
    return locale;
  }
  
  /**
   * Get list of single documents
   */
  public List<WtSingleDocument> getDocuments() {
    return documents;
  }

  /**
   * Get text level queue
   */
  public WtTextLevelCheckQueue getTextLevelCheckQueue() {
    return textLevelQueue;
  }
  
  /**
   * Get AI queue
   */
  public WtAiCheckQueue getAiCheckQueue() {
    return aiQueue;
  }
  
  /**
   * Get AI queue
   */
  public void setAiCheckQueue(WtAiCheckQueue queue) {
    aiQueue = queue;
  }
  
  /**
   * true, if LanguageTool is switched off
   */
  public boolean isBackgroundCheckOff() {
    return noBackgroundCheck;
  }

  /**
   * true, if Java look and feel is set
   */
  public static boolean isJavaLookAndFeelSet() {
    return javaLookAndFeelSet >= 0;
  }

  public static int getJavaLookAndFeelSet() {
    return javaLookAndFeelSet;
  }

  /**
   *  Toggle Switch Off / On of LT
   *  return true if toggle was done 
   */
  public boolean toggleNoBackgroundCheck() throws Throwable {
    if (docLanguage == null) {
      docLanguage = getCurrentLanguage();
    }
    if (config == null) {
      config = new WtConfiguration(configDir, configFile, docLanguage, true);
    }
    noBackgroundCheck = !noBackgroundCheck;
    if (!isBackgroundCheckOff()) {
      if (textLevelQueue != null) {
        textLevelQueue.setStop();
        textLevelQueue = null;
      }
      if (aiQueue != null) {
        aiQueue.setStop();
        aiQueue = null;
      }
    }
    setRecheck();
    config.saveNoBackgroundCheck(noBackgroundCheck, docLanguage);
    for (WtSingleDocument document : documents) {
      document.setConfigValues(config);
    }
    if (isBackgroundCheckOff()) {
      resetResultCaches(true);
    }
    return true;
  }

  /**
   * Set docID used within menu
   */
  public void setMenuDocId(String docId) {
    menuDocId = docId;
  }

  /**
   * Set use original spell und grammar dialog (for OO and old LO)
   */
  public void setUseOriginalCheckDialog() {
    useOrginalCheckDialog = true;
  }
  
  /**
   * Set use original spell und grammar dialog (for OO and old LO)
   */
  public boolean useOriginalCheckDialog() {
    return useOrginalCheckDialog;
  }
  
  /**
   * Is true if footnotes exist (tests if OO or very old LO) 
   *//*
  private void testFootnotes(PropertyValue[] propertyValues) {
    for (PropertyValue propertyValue : propertyValues) {
      if ("FootnotePositions".equals(propertyValue.Name)) {
        return;
      }
    }
    //  OO and LO < 4.3 do not support 'FootnotePositions' property and other advanced features
    //  switch back to single paragraph check mode
    //  use OOO configuration file - save existing settings if not already done
    useOrginalCheckDialog = true;
    File ooConfigFile = new File(configDir, OfficeTools.OOO_CONFIG_FILE);
    if (!ooConfigFile.exists()) {
      File loConfigFile = new File(configDir, configFile);
      if (loConfigFile.exists()) {
        try {
          Configuration tmpConfig = new Configuration(configDir, configFile, oldConfigFile, docLanguage, true);
          tmpConfig.setConfigFile(ooConfigFile);
          tmpConfig.setNumParasToCheck(0);
          tmpConfig.setUseTextLevelQueue(false);
          tmpConfig.saveConfiguration(docLanguage);
        } catch (IOException e) {
          MessageHandler.showError(e);
        }
      }
    }
    configFile = OfficeTools.OOO_CONFIG_FILE;
    MessageHandler.printToLogFile("No support of Footnotes: Open Office assumed - Single paragraph check mode set!");
  }
*/
  /**
   * Call method ignoreOnce for concerned document 
   */
  public String ignoreOnce() {
    for (WtSingleDocument document : documents) {
      if (menuDocId.equals(document.getDocID())) {
        return document.ignoreOnce();
      }
    }
    return null;
  }
  
  /**
   * Call method ignoreAll for concerned document 
   */
  public void ignoreAll() {
    for (WtSingleDocument document : documents) {
      if (menuDocId.equals(document.getDocID())) {
        document.ignoreAll();
        return;
      }
    }
  }
  
  /**
   * Call method replace ai spelling error
   */
  public void aiReplaceError(String suggestion) {
    for (WtSingleDocument document : documents) {
      if (menuDocId.equals(document.getDocID())) {
        document.replaceAiError(suggestion);
        return;
      }
    }
  }
  
  /**
   * Call method ignorePermanent for concerned document 
   */
  public String ignorePermanent() {
    for (WtSingleDocument document : documents) {
      if (menuDocId.equals(document.getDocID())) {
        return document.ignorePermanent();
      }
    }
    return null;
  }
  
  /**
   * Call method Ai Support mark errors 
   */
  public void aiAddErrorMarks() {
    for (WtSingleDocument document : documents) {
      if (menuDocId.equals(document.getDocID())) {
        WtAiErrorDetection aiError = new WtAiErrorDetection(document, config, lt);
        aiError.addAiRuleMatchesForParagraph(DetectionType.GRAMMAR);
        return;
      }
    }
  }
  
  /**
   * Call method Ai Support run command to change paragraph
   */
  public void runAiChangeOnParagraph(int commandId) {
    AiCommand command;
    switch (commandId) {
    case 1:
      command = AiCommand.CorrectGrammar;
      break;
    case 2:
      command = AiCommand.ImproveStyle;
      break;
    case 3:
      command = AiCommand.ReformulateText;
      break;
    case 4:
      command = AiCommand.ExpandText;
      break;
    default:
      command = AiCommand.GeneralAi;
    }
    for (WtSingleDocument document : documents) {
      if (menuDocId.equals(document.getDocID())) {
        WtAiParagraphChanging aiChange = new WtAiParagraphChanging(document, config, command);
        aiChange.start();
        return;
      }
    }
  }
  
  public void runAiTranslateText() {
    for (WtSingleDocument document : documents) {
      if (menuDocId.equals(document.getDocID())) {
        WtAiTranslateDocument aiTranslate = new WtAiTranslateDocument(document, messages);
        aiTranslate.start();
        return;
      }
    }
  }
  
  public void runAiTextToSpeech() {
    for (WtSingleDocument document : documents) {
      if (menuDocId.equals(document.getDocID())) {
        WtAiTextToSpeech aiTextToSpeech = new WtAiTextToSpeech(document, messages);
        aiTextToSpeech.start();
        return;
      }
    }
  }
  
  /**
   * Call method resetIgnorePermanent for concerned document 
   */
  public void resetIgnorePermanent() {
    getCurrentDocument().resetIgnorePermanent();
  }
  
  /**
   * Call method renewMarkups for concerned document 
   */
  public void renewMarkups() {
    for (WtSingleDocument document : documents) {
      if (menuDocId != null && menuDocId.equals(document.getDocID())) {
        document.renewMarkups();
      }
    }
  }
  
  /**
   * reset ignoreOnce information in all documents
   */
  public void resetIgnoreOnce() {
    for (WtSingleDocument document : documents) {
      document.resetIgnoreOnce();
    }
  }

  /**
   * change configuration profile 
   */
  private void changeProfile(String profile) {
    try {
      if (profile == null) {
        profile = "";
      }
      WtMessageHandler.printToLogFile("change to profile: " + profile);
      String currentProfile = config.getCurrentProfile();
      if (currentProfile == null) {
        currentProfile = "";
      }
      if (profile.equals(currentProfile)) {
        WtMessageHandler.printToLogFile("profile == null or profile equals current profile: Not changed");
        return;
      }
      List<String> definedProfiles = config.getDefinedProfiles();
      if (!profile.isEmpty() && (definedProfiles == null || !definedProfiles.contains(profile))) {
        WtMessageHandler.showMessage("profile '" + profile + "' not found");
      } else {
        List<String> saveProfiles = new ArrayList<>();
        saveProfiles.addAll(config.getDefinedProfiles());
        config.initOptions();
        config.loadConfiguration(profile == null ? "" : profile);
        config.setCurrentProfile(profile);
        config.addProfiles(saveProfiles);
        config.saveConfiguration(getCurrentDocument().getLanguage());
        resetConfiguration();
      }
    } catch (IOException e) {
      WtMessageHandler.showError(e);
    }
  }
  
  /**
   * Activate a rule by rule iD
   */
  public void activateRule(String ruleId) {
    activateRule(WtOfficeTools.localeToString(locale), ruleId);
  }
  
  public void activateRule(String langcode, String ruleId) {
    if (ruleId != null) {
      removeDisabledRule(langcode, ruleId);
      deactivateRule(ruleId, langcode, true);
      resetDocument();
    }
  }
  
  /**
   * Deactivate a rule as requested by the context menu
   */
  public void deactivateRule() {
    for (WtSingleDocument document : documents) {
      if (menuDocId.equals(document.getDocID())) {
        RuleDesc ruleDesc = document.getCurrentRule();
        if (ruleDesc != null) {
          if (debugMode) {
            WtMessageHandler.printToLogFile("MultiDocumentsHandler: deactivateRule: ruleID = "+ ruleDesc.error.aRuleIdentifier + "langCode = " + ruleDesc.langCode);
          }
          deactivateRule(ruleDesc.error.aRuleIdentifier, ruleDesc.langCode, false);
        }
        return;
      }
    }
  }
  
  /**
   * More Information to current error
   */
  private void moreInfo() {
    try  {
      if (infoDialog != null) {
        infoDialog.close();
        infoDialog = null;
      }
      for (WtSingleDocument document : documents) {
        if (menuDocId.equals(document.getDocID())) {
          RuleDesc ruleDesc = document.getCurrentRule();
          if (ruleDesc != null) {
            try {
              if (debugMode) {
                WtMessageHandler.printToLogFile("MultiDocumentsHandler: moreInfo: ruleID = " 
                              + ruleDesc.error.aRuleIdentifier + ", langCode = " + ruleDesc.langCode);
              }
              WtProofreadingError error = ruleDesc.error;
              //  Try to find exact rule ID
              for (Rule rule : lt.getAllRules()) {
//              WtMessageHandler.printToLogFile("ruleID: " + rule.getId() + "(search: " + error.aRuleIdentifier + ")");
                if (error.aRuleIdentifier.equals(rule.getId())) {
                  showMoreInfo(error, rule);
                  return;
                }
              }
              //  Rule not found: Try to find a rule ID equals the beginning of the searched ID
              for (Rule rule : lt.getAllRules()) {
                if (error.aRuleIdentifier.startsWith(rule.getId())) {
                  showMoreInfo(error, rule);
                  return;
                }
              }
              WtMessageHandler.printToLogFile("MultiDocumentsHandler: moreInfo: ruleID: " 
                  + ruleDesc.error.aRuleIdentifier + " not found (langCode = " + ruleDesc.langCode + ")");
            } catch (MalformedURLException e) {
              WtMessageHandler.showError(e);
            }
          } else {
            WtMessageHandler.printToLogFile("Current Rule not found");
          }
          return;
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }
  
  private void showMoreInfo(WtProofreadingError error, Rule rule) throws Throwable {
    String tmp = error.aShortComment;
    if (StringUtils.isEmpty(tmp)) {
      tmp = error.aFullComment;
    }
    String msg = org.writingtool.tools.WtGeneralTools.shortenComment(tmp);
    String sUrl = null;
    for (WtPropertyValue prop : error.aProperties) {
      if ("FullCommentURL".equals(prop.name)) {
        sUrl = (String) prop.value;
      }
    }
    URL url = sUrl == null? null : new URL(sUrl);
    MoreInfoDialogThread infoThread = new MoreInfoDialogThread(msg, error.aFullComment, rule, url,
        messages, lt.getLanguage().getShortCodeWithCountryAndVariant());
    infoThread.start();
  }

  
  /**
   * Remove a special Proofreading error from all caches
   */
  private void removeRuleError(String ruleId) {
    for (WtSingleDocument document : documents) {
      document.removeRuleError(ruleId);
    }
  }
  
  /**
   * Deactivate a rule by rule iD
   */
  public void deactivateRule(String ruleId, String langcode, boolean reactivate) {
    if (ruleId != null) {
      try {
        WtConfiguration confg = new WtConfiguration(configDir, configFile, docLanguage, true);
        Set<String> ruleIds = new HashSet<>();
        ruleIds.add(ruleId);
        if (reactivate) {
          confg.removeDisabledRuleIds(ruleIds);
          removeDisabledRule(langcode, ruleId);
          confg.saveConfiguration(docLanguage);
          initDocuments(true);
          resetDocument();
        } else {
          confg.addDisabledRuleIds(ruleIds);
          addDisabledRule(langcode, ruleId);
          confg.saveConfiguration(docLanguage);
          lt.disableRule(ruleId);
          initDocuments(false);
          removeRuleError(ruleId);
        }
        if (debugMode) {
          WtMessageHandler.printToLogFile("MultiDocumentsHandler: deactivateRule: Rule " + (reactivate ? "enabled: " : "disabled: ") 
              + (ruleId == null ? "null" : ruleId));
        }
      } catch (Throwable e) {
        WtMessageHandler.printException(e);
      }
    }
  }
  
  /**
   * We leave spell checking to OpenOffice/LibreOffice.
   * @return false
   */
  public final boolean isSpellChecker() {
    return true;
  }
  
  /**
   * Returns extra remote rules
   */
  public List<Rule> getExtraRemoteRules() {
    return extraRemoteRules;
  }

  /**
   * Returns xContext
   */
  public XComponentContext getContext() {
    return xContext;
  }

  /**
   * Runs LT options dialog box.
   */
  public void runOptionsDialog() {
    try {
      WtConfiguration config = getConfiguration();
      Language lang = config.getDefaultLanguage();
      if (lang == null) {
        lang = getCurrentLanguage();
      }
      if (lang == null) {
        return;
      }
      WtLanguageTool lTool = lt;
      if (lTool == null || !lang.equals(docLanguage)) {
        docLanguage = lang;
        lTool = initLanguageTool();
        config = WtDocumentsHandler.config;
      }
      config.initStyleCategories(lTool.getAllRules());
      WtConfigThread configThread = new WtConfigThread(lang, config, lTool, this);
      configThread.start();
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }

  /**
   * @return An array of Locales supported by LT
   */
  public final static Locale[] getLocales() {
    try {
      List<Locale> locales = new ArrayList<>();
      Locale locale = null;
      for (Language lang : Languages.get()) {
        if (lang.getCountries().length == 0) {
          if (lang.getDefaultLanguageVariant() != null) {
            if (lang.getDefaultLanguageVariant().getVariant() != null) {
              locale = new Locale(lang.getDefaultLanguageVariant().getShortCode(),
                  lang.getDefaultLanguageVariant().getCountries()[0], lang.getDefaultLanguageVariant().getVariant());
            } else if (lang.getDefaultLanguageVariant().getCountries().length != 0) {
              locale = new Locale(lang.getDefaultLanguageVariant().getShortCode(),
                  lang.getDefaultLanguageVariant().getCountries()[0], "");
            } else {
              locale = new Locale(lang.getDefaultLanguageVariant().getShortCode(), "", "");
            }
          }
          else if (lang.getVariant() != null) {  // e.g. Esperanto
            locale =new Locale(LIBREOFFICE_SPECIAL_LANGUAGE_TAG, "", lang.getShortCodeWithCountryAndVariant());
          } else {
            locale = new Locale(lang.getShortCode(), "", "");
          }
          if (locales != null && !WtOfficeTools.containsLocale(locales, locale)) {
            locales.add(locale);
          }
        } else {
          for (String country : lang.getCountries()) {
            if (lang.getVariant() != null) {
              locale = new Locale(LIBREOFFICE_SPECIAL_LANGUAGE_TAG, country, lang.getShortCodeWithCountryAndVariant());
            } else {
              locale = new Locale(lang.getShortCode(), country, "");
            }
            if (locales != null && !WtOfficeTools.containsLocale(locales, locale)) {
              locales.add(locale);
            }
          }
        }
      }
      return locales == null ? new Locale[0] : locales.toArray(new Locale[0]);
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      return new Locale[0];
    }
  }

  /**
   * Add a listener that allow re-checking the document after changing the
   * options in the configuration dialog box.
   * 
   * @param eventListener the listener to be added
   * @return true if listener is non-null and has been added, false otherwise
   */
  public final boolean addLinguServiceEventListener(XLinguServiceEventListener eventListener) {
    if (eventListener == null) {
      return false;
    }
    xEventListeners.add(eventListener);
    return true;
  }

  /**
   * Remove a listener from the event listeners list.
   * 
   * @param eventListener the listener to be removed
   * @return true if listener is non-null and has been removed, false otherwise
   */
  public final boolean removeLinguServiceEventListener(XLinguServiceEventListener eventListener) {
    if (eventListener == null) {
      return false;
    }
    if (xEventListeners.contains(eventListener)) {
      xEventListeners.remove(eventListener);
      return true;
    }
    return false;
  }

  /**
   * Inform listener that the document should be rechecked for grammar and style check.
   */
  public boolean resetCheck() {
    return resetCheck(LinguServiceEventFlags.PROOFREAD_AGAIN);
  }

  /**
   * Inform listener that the doc should be rechecked for a special event flag.
   */
  public boolean resetCheck(short eventFlag) {
    if (!xEventListeners.isEmpty()) {
      for (XLinguServiceEventListener xEvLis : xEventListeners) {
        if (xEvLis != null) {
          LinguServiceEvent xEvent = new LinguServiceEvent();
          xEvent.nEvent = eventFlag;
          xEvLis.processLinguServiceEvent(xEvent);
        }
      }
      return true;
    }
    return false;
  }

  /**
   * Configuration has be changed
   */
  public void resetConfiguration() {
    linguServices = null;
    if (config != null) {
      noBackgroundCheck = config.noBackgroundCheck();
      if (sidebarContent != null) {
        sidebarContent.setAiSupport(config.useAiSupport());
      }
    }
    javaLookAndFeelSet = -1;
    resetIgnoredMatches();
    resetResultCaches(true);
    resetDocument();
  }

  /**
   * Inform listener (grammar checking iterator) that options have changed and
   * the doc should be rechecked.
   */
  public void resetDocument() {
    setRecheck();
    resetCheck();
  }

  /**
   * Triggers the events from LT menu
   */
  public void trigger(String sEvent) {
    try {
//      MessageHandler.printToLogFile("Trigger event: " + sEvent);
      if ("noAction".equals(sEvent)) {  //  special dummy action
        return;
      }
      if (getCurrentDocument() == null) {
        return;
      }
      long startTime = 0;
      if (debugModeTm) {
        startTime = System.currentTimeMillis();
      }
      if (("checkDialog".equals(sEvent) || "checkAgainDialog".equals(sEvent)) && !useOrginalCheckDialog && !dialogIsRunning) {
        waitDialog = new WaitDialogThread("Please wait", messages.getString("loWaitMessage"));
        waitDialog.start();
      }
      if (!testDocLanguage(true)) {
        WtMessageHandler.printToLogFile("Test for document language failed: Can't trigger event: " + sEvent);
        return;
      }
      if ("configure".equals(sEvent)) {
        closeDialogs();
        runOptionsDialog();
      } else if ("about".equals(sEvent)) {
        if (aboutDialog != null) {
          aboutDialog.close();
          aboutDialog = null;
        }
        if (!isJavaLookAndFeelSet()) {
          setJavaLookAndFeel();
        }
        AboutDialogThread aboutThread = new AboutDialogThread(messages, xContext);
        aboutThread.start();
      } else if ("moreInfo".equals(sEvent)) {
        moreInfo();
      } else if ("toggleNoBackgroundCheck".equals(sEvent) || "backgroundCheckOn".equals(sEvent) || "backgroundCheckOff".equals(sEvent)) {
        if (toggleNoBackgroundCheck()) {
          resetCheck();
/*  TODO: in LT 6.5 add dynamic toolbar          
          for (SingleDocument document : documents) {
            LtToolbar ltToolbar = document.getLtToolbar();
            if (ltToolbar != null) {
              ltToolbar.makeToolbar(document.getLanguage());
            }
          }
*/
        }
      } else if ("ignoreOnce".equals(sEvent)) {
        ignoreOnce();
      } else if ("ignoreAll".equals(sEvent)) {
        ignoreAll();
      } else if ("ignorePermanent".equals(sEvent)) {
        ignorePermanent();
      } else if ("resetIgnorePermanent".equals(sEvent)) {
        resetIgnorePermanent();
      } else if ("deactivateRule".equals(sEvent)) {
        deactivateRule();
      } else if (sEvent.startsWith("aiReplaceWord_")) {
        String suggestion = sEvent.substring(14);
        this.aiReplaceError(suggestion);
      } else if (sEvent.startsWith("activateRule_")) {
        String ruleId = sEvent.substring(13);
        activateRule(ruleId);
      } else if (sEvent.startsWith("profileChangeTo_")) {
        String profile = sEvent.substring(16);
        changeProfile(profile);
      } else if (sEvent.startsWith("addToDictionary_")) {
        String[] sArray = sEvent.substring(16).split(":");
        WtDictionary.addWordToDictionary(sArray[0], sArray[1], xContext);;
      } else if ("renewMarkups".equals(sEvent)) {
        renewMarkups();
      } else if ("checkDialog".equals(sEvent) || "checkAgainDialog".equals(sEvent)) {
        if (useOrginalCheckDialog) {
          if ("checkDialog".equals(sEvent) ) {
            WtOfficeTools.dispatchCmd(".uno:SpellingAndGrammarDialog", xContext);
          } else {
            WtOfficeTools.dispatchCmd(".uno:RecheckDocument", xContext);
          }
          return;
        }
        closeDialogs();
        if (dialogIsRunning) {
          return;
        }
        if (waitDialog == null || waitDialog.canceled()) {
          return;
        }
        if (noBackgroundCheck && toggleNoBackgroundCheck()) {
          resetCheck();
        }
        setLtDialogIsRunning(true);
        WtCheckDialog checkDialog = new WtCheckDialog(xContext, this, docLanguage, waitDialog);
        if ("checkAgainDialog".equals(sEvent)) {
          WtSingleDocument document = getCurrentDocument();
          if (document != null) {
            XComponent currentComponent = document.getXComponent();
            if (currentComponent != null) {
              if (document.getDocumentType() == DocumentType.WRITER) {
                WtViewCursorTools viewCursor = new WtViewCursorTools(currentComponent);
                WtCheckDialog.setTextViewCursor(0, new TextParagraph (WtDocumentCache.CURSOR_TYPE_TEXT ,0), viewCursor);
              } else if (document.getDocumentType() == DocumentType.IMPRESS){
                WtOfficeDrawTools.setCurrentPage(0, currentComponent);
              } else {
                WtOfficeSpreadsheetTools.setCurrentSheet(0, currentComponent);
              }
            }
          }
          resetIgnoredMatches();
//          resetCheck();
        }
        if (debugMode) {
          WtMessageHandler.printToLogFile("MultiDocumentsHandler: trigger: Start Spell And Grammar Check Dialog");
        }
        checkDialog.start();
      } else if ("nextError".equals(sEvent)) {
        if (this.isBackgroundCheckOff()) {
          WtMessageHandler.showMessage(messages.getString("loExtSwitchOffMessage"));
          return;
        }
        if (useOrginalCheckDialog) {
          WtOfficeTools.dispatchCmd(".uno:SpellingAndGrammarDialog", xContext);
          return;
        }
        WtCheckDialog checkDialog = new WtCheckDialog(xContext, this, docLanguage, null);
        checkDialog.nextError();
      } else if ("refreshCheck".equals(sEvent)) {
        if (ltDialog != null) {
          ltDialog.closeDialog();
        } 
        if (this.isBackgroundCheckOff()) {
          WtMessageHandler.showMessage(messages.getString("loExtSwitchOffMessage"));
          return;
        }
        resetIgnoredMatches();
        resetDocumentCaches();
        resetResultCaches(true);
        WtSpellChecker.resetSpellCache();
        resetDocument();
      } else if ("statisticalAnalyses".equals(sEvent)) {
        if (!statAnDialogRunning) {
          statAnDialogRunning = true;
          WtStatAnDialog statAnDialog = new WtStatAnDialog(getCurrentDocument());
          statAnDialog.start();
        }
      } else if ("offStatisticalAnalyses".equals(sEvent)) {
        //  statistical analysis id not supported for this language --> do nothing
      } else if ("writeAnalyzedParagraphs".equals(sEvent)) {
        new WtAnalyzedParagraphsCache(this); 
      } else if ("aiAddErrorMarks".equals(sEvent)) {
        aiAddErrorMarks();
      } else if ("aiGeneralCommand".equals(sEvent)) {
        runAiChangeOnParagraph(0);
      } else if ("aiCorrectErrors".equals(sEvent)) {
        runAiChangeOnParagraph(1);
      } else if ("aiBetterStyle".equals(sEvent)) {
        runAiChangeOnParagraph(2);
      } else if ("aiReformulateText".equals(sEvent)) {
        runAiChangeOnParagraph(3);
      } else if ("aiAdvanceText".equals(sEvent)) {
        runAiChangeOnParagraph(4);
      } else if ("aiTranslateText".equals(sEvent)) {
        runAiTranslateText();
      } else if ("aiTextToSpeech".equals(sEvent)) {
        runAiTextToSpeech();
      } else if ("remoteHint".equals(sEvent)) {
        if (getConfiguration().useOtherServer()) {
          WtMessageHandler.showMessage(MessageFormat.format(messages.getString("loRemoteInfoOtherServer"), 
              getConfiguration().getServerUrl()));
        } else {
          WtMessageHandler.showMessage(messages.getString("loRemoteInfoDefaultServer"));
        }
      } else {
        WtMessageHandler.printToLogFile("MultiDocumentsHandler: trigger: Sorry, don't know what to do, sEvent = " + sEvent);
      }
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
        if (runTime > WtOfficeTools.TIME_TOLERANCE) {
          WtMessageHandler.printToLogFile("Time to run trigger: " + runTime);
        }
      }
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
  }
  
  /**
   * Test the language of the document
   * switch the check to LT if possible and language is supported
   */
  boolean testDocLanguage(boolean showMessage) throws Throwable {
    try {
      if (docLanguage == null) {
        if (linguServices == null) {
          linguServices = getLinguisticServices();
        }
        if (xContext == null) {
          if (showMessage) { 
            WtMessageHandler.showMessage("There may be a installation problem! \nNo xContext!");
          }
          return false;
        }
        XComponent xComponent = WtOfficeTools.getCurrentComponent(xContext);
        if (xComponent == null) {
          if (showMessage) { 
            WtMessageHandler.showMessage("There may be a installation problem! \nNo xComponent!");
          }
          return false;
        }
        Locale locale;
        DocumentType docType;
        if (WtOfficeDrawTools.isImpressDocument(xComponent)) {
          docType = DocumentType.IMPRESS;
          checkImpressDocument = true;
        } else if (WtOfficeSpreadsheetTools.isSpreadsheetDocument(xComponent)) {
          docType = DocumentType.CALC;
        } else {
          docType = DocumentType.WRITER;
        }
        if (docType == DocumentType.IMPRESS) {
          locale = WtOfficeDrawTools.getDocumentLocale(xComponent);
        } else if (docType == DocumentType.CALC) {
          locale = WtOfficeSpreadsheetTools.getDocumentLocale(xComponent);
        } else {
          locale = WtOfficeTools.getCursorLocale(xContext);
        }
        try {
          int n = 0;
          while (locale == null && n < 100) {
            Thread.sleep(500);
            if (debugMode) {
              WtMessageHandler.printToLogFile("MultiDocumentsHandler: testDocLanguage: Try to get locale: n = " + n);
            }
            if (docType == DocumentType.IMPRESS) {
              locale = WtOfficeDrawTools.getDocumentLocale(xComponent);
            } else if (docType == DocumentType.CALC) {
              locale = WtOfficeSpreadsheetTools.getDocumentLocale(xComponent);
            } else {
              locale = WtOfficeTools.getCursorLocale(xContext);
            }
            n++;
          }
        } catch (InterruptedException e) {
          WtMessageHandler.showError(e);
        }
        if (locale == null) {
          if (showMessage) {
            WtMessageHandler.showMessage("No Locale! LanguageTool can not be started!");
          } else {
            WtMessageHandler.printToLogFile("MultiDocumentsHandler: testDocLanguage: No Locale! LanguageTool can not be started!");
          }
          return false;
        } else if (!hasLocale(locale)) {
          String message = Tools.i18n(messages, "language_not_supported", locale.Language);
          WtMessageHandler.showMessage(message);
          return false;
        }
        if (debugMode) {
          WtMessageHandler.printToLogFile("MultiDocumentsHandler: testDocLanguage: locale: " + locale.Language + "-" + locale.Country);
        }
        if (!linguServices.setLtAsGrammarService(xContext, locale)) {
          if (showMessage) {
            WtMessageHandler.showMessage("Can not set LT as grammar check service! LanguageTool can not be started!");
          } else {
            WtMessageHandler.printToLogFile("MultiDocumentsHandler: testDocLanguage: Can not set LT as grammar check service! LanguageTool can not be started!");
          }
          return false;
        }
        if (docType != DocumentType.WRITER) {
          langForShortName = getLanguage(locale);
          if (langForShortName != null) {
            docLanguage = langForShortName;
            this.locale = locale;
            extraRemoteRules.clear();
            lt = initLanguageTool();
            initDocuments(true);
          }
          return true;
        } else {
          resetCheck();
          if (showMessage) {
            WtMessageHandler.showMessage(messages.getString("loNoGrammarCheckWarning"));
          }
          return false;
        }
      }
      return true;
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return false;
  }

  /**
   * Test if the needed java version is installed
   */
  public boolean javaVersionOkay() {
    String version = System.getProperty("java.version");
    if (version != null
        && (version.startsWith("1.0") || version.startsWith("1.1")
            || version.startsWith("1.2") || version.startsWith("1.3")
            || version.startsWith("1.4") || version.startsWith("1.5")
            || version.startsWith("1.6") || version.startsWith("1.7"))) {
      WtMessageHandler.showMessage("Error: LanguageTool requires Java 8 or later. Current version: " + version);
      return false;
    }
    return true;
  }
  
  /** Set Look and Feel for Java Swing Components
   * 
   */
  public static void setJavaLookAndFeel() {
    if (javaLookAndFeelSet < 0) {
      try {
        WtConfiguration config = WtDocumentsHandler.config;
        if (config == null) {
          Locale locale = new Locale("en","US","");
          config = getConfiguration(getLanguage(locale));
        }
        javaLookAndFeelSet = config == null ? -1 : config.getThemeSelection();
        WtGeneralTools.setJavaLookAndFeel(javaLookAndFeelSet < 0 ? 0 : javaLookAndFeelSet);
      } catch (Throwable t) {
        try {
          UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Throwable e) {
        }
      }
    }
  }
  
  /**
   * heap limit is reached
   */
  public boolean heapLimitIsReached() {
    return heapLimitReached;
  }

  /**
   * Test if enough heap space is left
   * Change to single paragraph mode if not
   * return false if heap space is to small 
   */
  public boolean isEnoughHeapSpace() {
    try {
      double heapRatio = WtOfficeTools.getCurrentHeapRatio();
      if (heapRatio >= 1.0) {
        heapLimitReached = true;
        setConfigValues(config, lt);
        WtMessageHandler.showMessage(messages.getString("loExtHeapMessage"));
        for (WtSingleDocument document : documents) {
          document.resetResultCache(true);
          document.resetDocumentCache();
        }
        return false;
      } else {
        if (heapRatio < 0.5) {
          heapCheckInterval = HEAP_CHECK_INTERVAL;
        } else if (heapRatio > 0.9) {
          heapCheckInterval = (int) (HEAP_CHECK_INTERVAL / (1.0 - heapCheckInterval));
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return true;
  }
  
  /**
   * run heap space test, in intervals
   */
  private void testHeapSpace() {
    if (!heapLimitReached && config.getNumParasToCheck() != 0) {
      if (numSinceHeapTest > heapCheckInterval) {
        isEnoughHeapSpace();
        numSinceHeapTest = 0;
      } else {
        numSinceHeapTest++;
      }
    }
  }
  
  /**
   * class to run the about dialog
   */
  private class AboutDialogThread extends Thread {

    private final ResourceBundle messages;
    private final XComponentContext xContext;

    AboutDialogThread(ResourceBundle messages, XComponentContext xContext) {
      this.messages = messages;
      this.xContext = xContext;
    }

    @Override
    public void run() {
      try {
        aboutDialog = new WtAboutDialog(messages);
        aboutDialog.show(xContext);
      } catch (Throwable t) {
        WtMessageHandler.showError(t);
      }
    }
    
  }

  /**
   * class to run the more info dialog
   */
  private class MoreInfoDialogThread extends Thread {

    private final String title;
    private final String message;
    private final Rule rule;
    private final URL matchUrl;
    private final ResourceBundle messages;
    private final String lang;

    MoreInfoDialogThread(String title, String message, Rule rule, URL matchUrl, ResourceBundle messages, String lang) {
      this.title = title;
      this.message = message;
      this.rule = rule;
      this.matchUrl = matchUrl;
      this.messages = messages;
      this.lang = lang;
    }

    @Override
    public void run() {
      try {
        infoDialog = new WtMoreInfoDialog(title, message, rule, matchUrl, messages, lang);
        infoDialog.show();
      } catch (Throwable t) {
        WtMessageHandler.showError(t);
      }
    }
    
  }

  /**
   * Called when "Ignore" is selected e.g. in the context menu for an error.
   */
  public void ignoreRule(String ruleId, Locale locale) {
    addDisabledRule(WtOfficeTools.localeToString(locale), ruleId);
    setRecheck();
  }

  /**
   * Called on rechecking the document - resets the ignore status for rules that
   * was set in the spelling dialog box or in the context menu.
   * 
   * The rules disabled in the config dialog box are left as intact.
   */
  public void resetIgnoreRules() {
    resetDisabledRules();
    setRecheck();
    resetIgnoreOnce();
    docReset = true;
  }

  /**
   * Get the displayed service name for LT
   */
  public static String getServiceDisplayName(Locale locale) {
    return WtOfficeTools.WT_DISPLAY_SERVICE_NAME;
  }

  /**
   * remove internal stored text if document disposes
   */
  public void disposing(EventObject source) {
    //  the data of document will be removed by next call of getNumDocID
    //  to finish checking thread without crashing
    try {
      XComponent goneComponent = UnoRuntime.queryInterface(XComponent.class, source.Source);
      if (goneComponent == null) {
        WtMessageHandler.printToLogFile("MultiDocumentsHandler: disposing: xComponent of closed document is null");
      } else {
        setContextOfClosedDoc(goneComponent);
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }
  
  /**
   * Start or stop the shape check loop
   */
  public void runShapeCheck (boolean hasShapes, int where) {
    try {
      if (doShapeCheck) {
        if (hasShapes && (shapeChangeCheck == null || !shapeChangeCheck.isRunning())) {
          WtMessageHandler.printToLogFile("MultiDocumentsHandler: runShapeCheck: start");
          shapeChangeCheck = new ShapeChangeCheck();
          shapeChangeCheck.start();
        } else if (!hasShapes && shapeChangeCheck != null) {
          boolean noShapes = true;
          for (int i = 0; i < documents.size() && noShapes; i++) {
            if (documents.get(i).getDocumentCache().textSize(WtDocumentCache.CURSOR_TYPE_SHAPE) > 0) {
              noShapes = false;
            }
          }
          if (noShapes) {
            WtMessageHandler.printToLogFile("MultiDocumentsHandler: runShapeCheck: stop");
            shapeChangeCheck.stopLoop();
            shapeChangeCheck = null;
          }
        }
      }
    } catch (Exception e) {
      WtMessageHandler.showError(e);
    }
  }
  
  /**
   *  start a separate thread to add or remove the internal LT dictionary
   *//*
  private void handleLtDictionary(String text) {
    LtDictionary.setIgnoreWordsForSpelling(text, lt, locale, xContext);
  }
*/
  /**
   *  start a separate thread to add or remove the internal LT dictionary
   *//*
  public void handleLtDictionary(String text, Locale locale) {
    handleDictionary.addTextToCheck(text, locale);
    handleDictionary.wakeup();
  }
*/
  /**
   *  class to start a separate thread to add or remove the internal LT dictionary
   *//*
  private class HandleLtDictionary extends Thread {
    private Object queueWakeup = new Object();
    private final List<String> textToCheck = new ArrayList<>();
    private final List<Locale> localeToCheck = new ArrayList<>();
    private boolean isRunning = true;
    
    void setStop() {
      isRunning = false;
    }
    
    void addTextToCheck (String text, Locale locale) {
      textToCheck.add(text);
      localeToCheck.add(locale);
    }
    
    void wakeup() {
      synchronized(queueWakeup) {
//        if (debugMode) {
//          MessageHandler.printToLogFile("HandleLtDictionary: wakeupQueue: wake queue");
//        }
        queueWakeup.notify();
      }
    }

    @Override
    public void run() {
      isRunning = true;
      while (isRunning) {
        synchronized(queueWakeup) {
          if (textToCheck.size() < 1) {
            try {
//              MessageHandler.printToLogFile("HandleLtDictionary: run: queue waits");
              queueWakeup.wait();
            } catch (InterruptedException e) {
              MessageHandler.printException(e);
            }
          }
        }
        String text = textToCheck.get(0);
        Locale locale = localeToCheck.get(0);
        textToCheck.remove(0);
        localeToCheck.remove(0);
        OfficeTools.waitForLO();
        if (LtDictionary.setIgnoreWordsForSpelling(text, lt, locale, xContext)) {
          resetCheck(LinguServiceEventFlags.SPELL_WRONG_WORDS_AGAIN); 
        }
      }
      isRunning = false;
    }
  }
*/
  /**
   * class to test for text changes in shapes 
   */
  private class ShapeChangeCheck extends Thread {

    boolean runLoop = true;

    @Override
    public void run() {
      try {
        while (runLoop && textLevelQueue != null) {
          try {
            for (int i = 0; i < documents.size(); i++) {
              documents.get(i).addShapeQueueEntries();
            }
            Thread.sleep(WtOfficeTools.CHECK_SHAPES_TIME);
          } catch (Throwable t) {
            WtMessageHandler.printException(t);
          }
        }
      } catch (Throwable e) {
        WtMessageHandler.showError(e);
      }
      runLoop = false;
    }
    
    public void stopLoop() {
      runLoop = false;
    }

    public boolean isRunning() {
      return runLoop;
    }

  }

  /** class to start a separate thread to check for Impress documents
   */
  private class LtHelper extends Thread {
    
    @Override
    public void run() {
      try {
        WtSingleDocument currentDocument = null;
        while (!isHelperDisposed && currentDocument == null) {
          Thread.sleep(250);
          if (isHelperDisposed) {
            return;
          }
          currentDocument = getCurrentDocument();
          if (currentDocument != null && (currentDocument.getDocumentType() == DocumentType.IMPRESS 
              || currentDocument.getDocumentType() == DocumentType.CALC)) {
            if (currentDocument.getDocumentType() == DocumentType.IMPRESS) {
              checkImpressDocument = true;
              locale = WtOfficeDrawTools.getDocumentLocale(currentDocument.getXComponent());
            } else {
              locale = WtOfficeSpreadsheetTools.getDocumentLocale(currentDocument.getXComponent());
            }
            if (locale == null) {
              locale = new Locale("en","US","");
            }
            WtMessageHandler.printToLogFile("MultiDocumentsHandler: LtHelper: local: " + WtOfficeTools.localeToString(locale));
            langForShortName = getLanguage(locale);
            if (langForShortName != null) {
              docLanguage = langForShortName;
              lt = initLanguageTool();
              initDocuments(false);
            }
          }
        }
      } catch (Throwable e) {
        WtMessageHandler.showError(e);
      }
    }
    
  }

  /**
   * class to run a dialog in a separate thread
   * closing if lost focus
   */
  public static class WaitDialogThread extends Thread {
    private final String dialogName;
    private final String text;
    private int max;
    private JDialog dialog = null;
    private boolean isCanceled = false;
    private JProgressBar progressBar;

    public WaitDialogThread(String dialogName, String text) {
      this.dialogName = dialogName;
      this.text = text;
      progressBar = new JProgressBar();
      if (!isJavaLookAndFeelSet()) {
        WtDocumentsHandler.setJavaLookAndFeel();
      }
    }

    @Override
    public void run() {
      try {
        JLabel textLabel = new JLabel(text);
        JButton cancelBottom = new JButton(messages.getString("guiCancelButton"));
        cancelBottom.addActionListener(e -> {
          close_intern();
        });
        progressBar.setIndeterminate(true);
        dialog = new JDialog();
        Container contentPane = dialog.getContentPane();
        dialog.setName("InformationThread");
        dialog.setTitle(dialogName);
        dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
        dialog.addWindowListener(new WindowListener() {
          @Override
          public void windowOpened(WindowEvent e) {
          }
          @Override
          public void windowClosing(WindowEvent e) {
            close_intern();
          }
          @Override
          public void windowClosed(WindowEvent e) {
          }
          @Override
          public void windowIconified(WindowEvent e) {
          }
          @Override
          public void windowDeiconified(WindowEvent e) {
          }
          @Override
          public void windowActivated(WindowEvent e) {
          }
          @Override
          public void windowDeactivated(WindowEvent e) {
          }
        });
        JPanel panel = new JPanel();
        panel.setLayout(new GridBagLayout());
        GridBagConstraints cons = new GridBagConstraints();
        cons.insets = new Insets(16, 24, 16, 24);
        cons.gridx = 0;
        cons.gridy = 0;
        cons.weightx = 1.0f;
        cons.weighty = 10.0f;
        cons.anchor = GridBagConstraints.CENTER;
        cons.fill = GridBagConstraints.BOTH;
        panel.add(textLabel, cons);
        cons.gridy++;
        panel.add(progressBar, cons);
        cons.gridy++;
        cons.fill = GridBagConstraints.NONE;
        panel.add(cancelBottom, cons);
        contentPane.setLayout(new GridBagLayout());
        cons = new GridBagConstraints();
        cons.insets = new Insets(16, 32, 16, 32);
        cons.gridx = 0;
        cons.gridy = 0;
        cons.weightx = 1.0f;
        cons.weighty = 1.0f;
        cons.anchor = GridBagConstraints.NORTHWEST;
        cons.fill = GridBagConstraints.BOTH;
        contentPane.add(panel);
        dialog.pack();
        Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
        Dimension frameSize = dialog.getSize();
        dialog.setLocation(screenSize.width / 2 - frameSize.width / 2,
            screenSize.height / 2 - frameSize.height / 2);
        dialog.setAutoRequestFocus(true);
        dialog.setAlwaysOnTop(true);
        dialog.toFront();
//        if (debugMode) {
          WtMessageHandler.printToLogFile(dialogName + ": run: Dialog is running");
//        }
        dialog.setVisible(true);
        if (isCanceled) {
          dialog.setVisible(false);
          dialog.dispose();
        }
      } catch (Throwable t) {
        WtMessageHandler.showError(t);
      }
    }
    
    public boolean canceled() {
      return isCanceled;
    }
    
    public void close() {
      close_intern();
    }
    
    private void close_intern() {
      try {
//        if (debugMode) {
          WtMessageHandler.printToLogFile("WaitDialogThread: close: Dialog closed");
//        }
        isCanceled = true;
        if (dialog != null) {
          if (dialog.isVisible()) {
            dialog.setVisible(false);
          }
          dialog.dispose();
        }
      } catch (Throwable t) {
        WtMessageHandler.showError(t);
      }
    }
    
    public void initializeProgressBar(int min, int max) {
      if (progressBar != null) {
        this.max = max;
        progressBar.setMinimum(min);
        progressBar.setMaximum(max);
        progressBar.setStringPainted(true);
        progressBar.setIndeterminate(false);
      }
    }
    
    public void setValueForProgressBar(int val, boolean setText) {
      if (progressBar != null) {
        progressBar.setValue(val);
        progressBar.setIndeterminate(false);
        if (setText) {
          int p = (int) (((val * 100) / max) + 0.5);
          progressBar.setString(p + " %  ( " + val + " / " + max + " )");
          progressBar.setStringPainted(true);
        }
      }
    }
    
    public void setIndeterminate() {
      progressBar.setIndeterminate(true);
    }
  }
}


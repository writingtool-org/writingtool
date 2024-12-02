/* WritingTool, a LibreOffice Extension based on LanguageTool
 * Copyright (C) 2024 Fred Kruse (https://fk-es.de)
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

import static java.lang.System.arraycopy;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.languagetool.Language;
import org.writingtool.WtCacheIO.SpellCache;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.WtTextLevelCheckQueue.QueueEntry;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtDocumentCursorTools;
import org.writingtool.tools.WtFlatParagraphTools;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtViewCursorTools;
import org.writingtool.tools.WtOfficeTools.DocumentType;
import org.writingtool.tools.WtOfficeTools.LoErrorType;

import com.sun.star.awt.MouseButton;
import com.sun.star.awt.MouseEvent;
import com.sun.star.awt.XMouseClickHandler;
import com.sun.star.awt.XUserInputInterception;
import com.sun.star.beans.PropertyValue;
import com.sun.star.container.XIndexAccess;
import com.sun.star.document.DocumentEvent;
import com.sun.star.document.XDocumentEventBroadcaster;
import com.sun.star.document.XDocumentEventListener;
import com.sun.star.frame.XController;
import com.sun.star.frame.XModel;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.Locale;
import com.sun.star.lang.XComponent;
import com.sun.star.linguistic2.ProofreadingResult;
import com.sun.star.linguistic2.SingleProofreadingError;
import com.sun.star.text.TextMarkupType;
import com.sun.star.text.XFlatParagraph;
import com.sun.star.text.XTextRange;
import com.sun.star.ui.ContextMenuExecuteEvent;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.view.XSelectionSupplier;

/**
 * Class for checking text of one LO document 
 */
public class WtSingleDocument {
  
  /**
   * Full text Check:
   * numParasToCheck: Paragraphs to be checked for full text rules
   * < 0 check full text (time intensive)
   * == 0 check only one paragraph (works like LT Version <= 3.9)
   * > 0 checks numParasToCheck before and after the processed paragraph
   * 
   * Cache:
   * sentencesCache: only used for doResetCheck == true (LO checks again only changed paragraphs by default)
   * paragraphsCache: used to store LT matches for a fast return to LO (numParasToCheck != 0)
   * singleParaCache: used for one paragraph check by default or for special paragraphs like headers, footers, footnotes, etc.
   *  
   */
  
  private static int debugMode;                   //  should be 0 except for testing; 1 = low level; 2 = advanced level
  private static boolean debugModeTm;             // time measurement should be false except for testing
  
  private WtConfiguration config;

  private int numParasToCheck = 0;                // current number of Paragraphs to be checked

  private XComponentContext xContext;             //  The context of the document
  private String docID;                           //  docID of the document
  private XComponent xComponent;                  //  XComponent of the open document
  private final WtDocumentsHandler mDocHandler;      //  handles the different documents loaded in LO/OO
  private LTDokumentEventListener eventListener = null; //  listens for save of document 
  
  private final WtDocumentCache docCache;           //  cache of paragraphs (only readable by parallel thread)
  private final List<WtResultCache> paragraphsCache;//  Cache for matches of text rules
  private final Map<Integer, String> changedParas;//  Map of last changed paragraphs;
  private final Set<Integer> runningParas;         //  List of running checks for paragraphs;
  private WtDocumentCursorTools docCursor = null;   //  Save document cursor for the single document
//  private ViewCursorTools viewCursor = null;      //  Get the view cursor for desktop
  private WtFlatParagraphTools flatPara = null;     //  Save information for flat paragraphs (including iterator and iterator provider) for the single document
  private Integer numLastVCPara = 0;              //  Save position of ViewCursor for the single documents
  private final List<Integer> numLastFlPara;      //  Save position of FlatParagraph for the single documents
  private WtCacheIO cacheIO = null;
  private int changeFrom = 0;                     //  Change result cache from paragraph
  private int changeTo = 0;                       //  Change result cache to paragraph
  private int paraNum;                            //  Number of current checked paragraph
  private int lastChangedPara;                    //  lastPara which was detected as changed
  private List<Integer> lastChangedParas;         //  lastPara which was detected as changed
  private WtIgnoredMatches ignoredMatches;          //  Map of matches (number of paragraph, number of character) that should be ignored after ignoreOnce was called
  private WtIgnoredMatches permanentIgnoredMatches; //  Map of matches (number of paragraph, number of character) that should be ignored permanent
  private final DocumentType docType;             //  save the type of document
  private boolean disposed = false;               //  true: document with this docId is disposed - SingleDocument shall be removed
  private boolean resetDocCache = false;          //  true: the cache of the document should be reset before the next check
  private boolean hasFootnotes = true;            //  true: Footnotes are supported by LO/OO
  private boolean hasSortedTextId = true;         //  true: Node Index is supported by LO
  private boolean isLastIntern = false;           //  true: last check was intern
  private boolean isRightButtonPressed = false;   //  true: right mouse Button was pressed
  private boolean isOnUnload = false;             //  Document will be closed
  private String lastSinglePara = null;           //  stores the last paragraph which is checked as single paragraph
  private Language docLanguage;                   //  docLanguage (usually the Language of the first paragraph)
//  private Locale docLocale;                       //  docLanguage as Locale
  private final Language fixedLanguage;           //  fixed language (by configuration); if null: use language of document (given by LO/OO)
  private WtMenus ltMenus = null;                 //  LT menus (tools menu and context menu)
//  TODO: add in 6.5   private LtToolbar ltToolbar = null;             //  LT dynamic toolbar
  private WtResultCache statAnCache = null;         //  Cache for results of statistical analysis
  private String statAnRuleId = null;             //  RuleId of current statistical rule tested

  WtSingleDocument(XComponentContext xContext, WtConfiguration config, String docID, 
      XComponent xComp, WtDocumentsHandler mDH, Language lang) {
    numLastFlPara = new ArrayList<>();
    for (int i = 0; i < WtDocumentCache.NUMBER_CURSOR_TYPES + 1; i++) {
      numLastFlPara.add(-1);
    }
    debugMode = WtOfficeTools.DEBUG_MODE_SD;
    debugModeTm = WtOfficeTools.DEBUG_MODE_TM;
    if (WtOfficeTools.DEVELOP_MODE_ST) {
      hasSortedTextId = false;
    }
    this.xContext = xContext;
    this.config = config;
    this.docID = docID;
    docLanguage = lang;
//    docLocale = WtLinguisticServices.getLocale(lang);
    if (docID.charAt(0) == 'I') {
      docType = DocumentType.IMPRESS;
    } else if (docID.charAt(0) == 'C') {
      docType = DocumentType.CALC;
    } else {
      docType = DocumentType.WRITER;
    }
    xComponent = xComp;
    mDocHandler = mDH;
    fixedLanguage = config == null ? null : config.getDefaultLanguage();
    changedParas = new HashMap<Integer, String>();
    runningParas = new HashSet<>();
    setDokumentListener(xComponent);
    List<WtResultCache> paraCache = new ArrayList<>();
    for (int i = 0; i < WtOfficeTools.NUMBER_CACHE; i++) {
      paraCache.add(new WtResultCache());
    }
    paragraphsCache = Collections.unmodifiableList(paraCache);
    if (config != null) {
      setConfigValues(config);
    }
    resetResultCache(true);
    ignoredMatches = new WtIgnoredMatches();
    permanentIgnoredMatches = new WtIgnoredMatches();
    docCache = new WtDocumentCache(docType);
    if (config != null && config.saveLoCache() && !config.noBackgroundCheck() && xComponent != null && !mDocHandler.isTestMode()) {
      readCaches();
    }
    if (xComponent != null) {
      setFlatParagraphTools();
    }
    if (!mDocHandler.isOpenOffice && (docType == DocumentType.IMPRESS 
        || (mDH.isBackgroundCheckOff() && docType == DocumentType.WRITER)) && ltMenus == null) {
      ltMenus = new WtMenus(xContext, this, config);
    }
/*  TODO: in LT 6.5 add dynamic toolbar          
    if (!mDocHandler.isOpenOffice && docType == DocumentType.WRITER) {
      ltToolbar = new LtToolbar(xContext, this, docLanguage);
    }
*/
  }
  
  /**  get the result for a check of a single document 
   * 
   * @param paraText          paragraph text
   * @param paRes             proof reading result
   * @return                  proof reading result
   */
  ProofreadingResult getCheckResults(String paraText, Locale locale, ProofreadingResult paRes, 
      PropertyValue[] propertyValues, boolean docReset, WtLanguageTool lt, LoErrorType errType) {
    return getCheckResults(paraText, locale, paRes, propertyValues, docReset, lt, -1, errType);
  }
    
  public ProofreadingResult getCheckResults(String paraText, Locale locale, ProofreadingResult paRes, 
      PropertyValue[] propertyValues, boolean docReset, WtLanguageTool lt, int nPara, LoErrorType errType) {
    try {
      boolean isIntern = (nPara >= 0);
      boolean isMouseRequest = false;
      if (isRightButtonPressed) {
        isMouseRequest = true;
        isRightButtonPressed = false;
      }
      int [] footnotePositions = null;  // e.g. for LO/OO < 4.3 and the 'FootnotePositions' property
      int proofInfo = WtOfficeTools.PROOFINFO_UNKNOWN;  //  OO and LO < 6.5 do not support ProofInfo
      int sortedTextId = -1;
      int documentElementsCount = -1;
      for (PropertyValue propertyValue : propertyValues) {
        if ("FootnotePositions".equals(propertyValue.Name)) {
          if (propertyValue.Value instanceof int[]) {
            footnotePositions = (int[]) propertyValue.Value;
          } else {
            WtMessageHandler.printToLogFile("SingleDocument: getCheckResults: Not of expected type int[]: " + propertyValue.Name + ": " + propertyValue.Value.getClass());
          }
        }
        if ("ProofInfo".equals(propertyValue.Name)) {
          if (propertyValue.Value instanceof Integer) {
            proofInfo = (int) propertyValue.Value;
          } else {
            WtMessageHandler.printToLogFile("SingleDocument: getCheckResults: Not of expected type int: " + propertyValue.Name + ": " + propertyValue.Value.getClass());
          }
        }
        if (!isIntern && hasSortedTextId) {
          if ("SortedTextId".equals(propertyValue.Name)) {
            if (propertyValue.Value instanceof Integer) {
              sortedTextId = (int) propertyValue.Value;
            } else {
              WtMessageHandler.printToLogFile("SingleDocument: getCheckResults: Not of expected type int: " + propertyValue.Name + ": " + propertyValue.Value.getClass());
            }
          }
          if ("DocumentElementsCount".equals(propertyValue.Name)) {
            if (propertyValue.Value instanceof Integer) {
              documentElementsCount = (int) propertyValue.Value;
            } else {
              WtMessageHandler.printToLogFile("SingleDocument: getCheckResults: Not of expected type int: " + propertyValue.Name + ": " + propertyValue.Value.getClass());
            }
          }
        }
      }
      if (!isIntern && hasSortedTextId && sortedTextId < 0) {
        hasSortedTextId = false;
        WtMessageHandler.printToLogFile("SingleDocument: getCheckResults: SortedTextId and DocumentElementsCount are not supported by LO!");
      }
      if (debugMode > 0 && hasSortedTextId) {
        WtMessageHandler.printToLogFile("SingleDocument: getCheckResults: sortedTextId: " + sortedTextId);
        WtMessageHandler.printToLogFile("SingleDocument: getCheckResults: documentElementsCount: " + documentElementsCount);
      }
      hasFootnotes = footnotePositions != null;
      if (!hasFootnotes) {
        //  OO and LO < 4.3 do not support 'FootnotePositions' property and other advanced features
        //  switch back to single paragraph check mode - save settings in configuration
        if (numParasToCheck != 0) {
          if (config.useTextLevelQueue()) {
            mDocHandler.getTextLevelCheckQueue().setStop();
          }
          numParasToCheck = 0;
          config.setNumParasToCheck(numParasToCheck);
          config.setUseTextLevelQueue(false);
          try {
            config.saveConfiguration(docLanguage);
          } catch (IOException e) {
            WtMessageHandler.showError(e);
          }
          WtMessageHandler.printToLogFile("Single paragraph check mode set!");
        }
        mDocHandler.setUseOriginalCheckDialog();
      }
      
      if (proofInfo == WtOfficeTools.PROOFINFO_GET_PROOFRESULT) {
        int nFPara = -1;
        if (hasSortedTextId) {
          nFPara = docCache.getFlatparagraphFromSortedTextId(sortedTextId);
          if (debugMode > 0) {
            WtMessageHandler.printToLogFile("SingleDocument: getCheckResults: get errors direct from cache, nFPara: " + nFPara);
          }
          if (nFPara >= 0) {
            return getErrorsFromCache(nFPara, paRes, paraText, locale, lt);
          }
        }
        if ((WtDocumentCursorTools.isBusy() || WtViewCursorTools.isBusy() || WtFlatParagraphTools.isBusy() || docCache.isResetRunning())) {
          //  NOTE: LO blocks the read of information by document or view cursor tools till a PROOFINFO_GET_PROOFRESULT request is done
          //        This causes a hanging of LO when the request isn't answered immediately by a 0 matches result
          WtSingleCheck singleCheck = new WtSingleCheck(this, paragraphsCache, fixedLanguage,
              docLanguage, numParasToCheck, true, isMouseRequest, false);
          paRes.aErrors = WtOfficeTools.wtErrorsToProofreading(singleCheck.checkParaRules(paraText, locale, 
                            footnotePositions, -1, paRes.nStartOfSentencePosition, lt, 0, 0, false, false, errType));
          closeDocumentCursor();
          return paRes;
        }
      }
      if (debugMode > 0 && proofInfo == WtOfficeTools.PROOFINFO_GET_PROOFRESULT) {
        WtMessageHandler.printToLogFile("SingleDocument: getCheckResults: start PROOFRESULT");
      }
      if (resetDocCache) {
        if (debugMode > 0 && proofInfo == WtOfficeTools.PROOFINFO_GET_PROOFRESULT) {
          WtMessageHandler.printToLogFile("SingleDocument: getCheckResults: is resetDocCache");
        }
        if (docCursor == null) {
          if (debugMode > 0 && proofInfo == WtOfficeTools.PROOFINFO_GET_PROOFRESULT) {
            WtMessageHandler.printToLogFile("SingleDocument: getCheckResults: get docCursor");
          }
          docCursor = getDocumentCursorTools();
        }
        if (debugMode > 0 && proofInfo == WtOfficeTools.PROOFINFO_GET_PROOFRESULT) {
          WtMessageHandler.printToLogFile("SingleDocument: getCheckResults: refresh docCache");
        }
        docCache.refresh(this, WtLinguisticServices.getLocale(fixedLanguage), 
            WtLinguisticServices.getLocale(docLanguage),xComponent, 6);
        resetDocCache = false;
      }
      if (docLanguage == null) {
        docLanguage = lt.getLanguage();
      }
      if (disposed) {
        closeDocumentCursor();
  //      viewCursor = null;
        return paRes;
      }
      if (docReset) {
        numLastVCPara = 0;
        ignoredMatches = new WtIgnoredMatches();
      }
      boolean isDialogRequest = (nPara >= 0 || (proofInfo == WtOfficeTools.PROOFINFO_GET_PROOFRESULT));
      
      WtCheckRequestAnalysis requestAnalysis = null;
      long startTime = 0;
      if (debugModeTm) {
        startTime = System.currentTimeMillis();
      }
      int paraNum;
//      MessageHandler.printToLogFile("Single document: Check Paragraph: " + paraText);
      if (hasSortedTextId) {
        if (isIntern) {
          paraNum = nPara;
        } else {
          requestAnalysis = new WtCheckRequestAnalysis(numLastVCPara, numLastFlPara,
            proofInfo, numParasToCheck, fixedLanguage, docLanguage, this, paragraphsCache, changedParas, runningParas);
          paraNum = requestAnalysis.getNumberOfParagraphFromSortedTextId(sortedTextId, documentElementsCount, paraText, locale, footnotePositions);
        }
      } else {
        requestAnalysis = new WtCheckRequestAnalysis(numLastVCPara, numLastFlPara,
            proofInfo, numParasToCheck, fixedLanguage, docLanguage, this, paragraphsCache, changedParas, runningParas);
        paraNum = requestAnalysis.getNumberOfParagraph(nPara, paraText, locale, paRes.nStartOfSentencePosition, footnotePositions);
      }
      if (docCache.isAutomaticGenerated(paraNum, true)) {
        return paRes;
      }
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
        if (runTime > WtOfficeTools.TIME_TOLERANCE) {
          WtMessageHandler.printToLogFile("Single document: Time to run request analyses: " + runTime);
        }
      }
      if (debugMode > 1) {
        WtMessageHandler.printToLogFile("Single document: getCheckResults: paraNum = " + paraNum + ", nPara = " + nPara);
      }
      if (paraNum == -2) {
        paraNum = isLastIntern ? this.paraNum : -1;
        if (debugMode > 1) {
          WtMessageHandler.printToLogFile("Single document: getCheckResults: paraNum set to: " + paraNum + ", isLastIntern = " + isLastIntern);
        }
      }
      this.paraNum = paraNum;
      runningParas.add(paraNum);
      isLastIntern = isIntern;
      boolean textIsChanged = false;
      if (!isIntern && requestAnalysis != null) {
        changeFrom = requestAnalysis.getFirstParagraphToChange();
        changeTo = requestAnalysis.getLastParagraphToChange();
        numLastVCPara = requestAnalysis.getLastParaNumFromViewCursor();
        textIsChanged = requestAnalysis.textIsChanged();
      }
      
      if (disposed) {
        closeDocumentCursor();
//        viewCursor = null;
        return paRes;
      }
      if (debugModeTm) {
        startTime = System.currentTimeMillis();
      }
//      MessageHandler.printToLogFile("Single document: Check Paragraph: " + paraNum);
      WtSingleCheck singleCheck = new WtSingleCheck(this, paragraphsCache, fixedLanguage,
          docLanguage, numParasToCheck, isDialogRequest, isMouseRequest, isIntern);
      paRes.aErrors = WtOfficeTools.wtErrorsToProofreading(singleCheck.getCheckResults(paraText, footnotePositions, locale, lt, paraNum, 
          paRes.nStartOfSentencePosition, textIsChanged, changeFrom, changeTo, lastSinglePara, lastChangedPara, errType));
//    MessageHandler.printToLogFile("Single document: Check Paragraph: " + paraNum + " done");
//      MessageHandler.printToLogFile("Single document: Check Paragraph: resultCache for Para " + paraNum + ": "
//          + (paragraphsCache.get(0).getCacheEntry(paraNum) == null 
//          ? "null" : paragraphsCache.get(0).getCacheEntry(paraNum).errorArray.length));
      lastSinglePara = singleCheck.getLastSingleParagraph();
      paRes.nStartOfSentencePosition = paragraphsCache.get(0).getStartSentencePosition(paraNum, paRes.nStartOfSentencePosition);
      paRes.nStartOfNextSentencePosition = paragraphsCache.get(0).getNextSentencePosition(paraNum, paRes.nStartOfSentencePosition);
      if (paRes.nStartOfNextSentencePosition == 0) {
        paRes.nStartOfNextSentencePosition = paraText.length();
      }
      paRes.nBehindEndOfSentencePosition = paRes.nStartOfNextSentencePosition;
      lastChangedPara = (textIsChanged && numParasToCheck != 0) ? paraNum : -1;
      if (debugModeTm) {
        long runTime = System.currentTimeMillis() - startTime;
        if (runTime > WtOfficeTools.TIME_TOLERANCE) {
          WtMessageHandler.printToLogFile("Single document: Time to run single check: " + runTime);
        }
      }
      if (proofInfo == WtOfficeTools.PROOFINFO_GET_PROOFRESULT || isIntern) {
        addStatAnalysisErrors (paRes, paraNum);
        if (debugModeTm) {
          startTime = System.currentTimeMillis();
        }
        addSynonyms(paRes, paraText, locale, lt);
        if (debugModeTm) {
          long runTime = System.currentTimeMillis() - startTime;
          if (runTime > WtOfficeTools.TIME_TOLERANCE) {
            WtMessageHandler.printToLogFile("Single document: Time to addSynonyms: " + runTime);
          }
        }
      }
      if (textIsChanged && numParasToCheck != 0 && config.useTextLevelQueue() && !isDialogRequest
          && mDocHandler.getTextLevelCheckQueue() != null && !mDocHandler.isTestMode()) {
        mDocHandler.getTextLevelCheckQueue().wakeupQueue(docID);
      }
      if (ltMenus == null && !mDocHandler.isOpenOffice && docType == DocumentType.WRITER && paraText.length() > 0) {
        ltMenus = new WtMenus(xContext, this, config);
      }
/*  TODO: in LT 6.5 add dynamic toolbar          
      if (!mDocHandler.isOpenOffice && docType == DocumentType.WRITER && docCache != null && docCache.getDocumentLocale() != null
          && docLocale != null && !OfficeTools.isEqualLocale(docLocale, docCache.getDocumentLocale())) {
        docLocale = docCache.getDocumentLocale();
        ltToolbar.makeToolbar(getLanguage());
      }
*/
      if (proofInfo == WtOfficeTools.PROOFINFO_MARK_PARAGRAPH) {
        paRes.aErrors = WtOfficeTools.wtErrorsToProofreading(
            filterOverlappingErrors(WtOfficeTools.proofreadingToWtErrors(paRes.aErrors), config.filterOverlappingMatches()));
      }

    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    } finally {
      runningParas.remove(paraNum);
    }
    closeDocumentCursor();
 //   viewCursor = null;
    return paRes;
  }
  
  /**
   * set values set by configuration dialog
   */
  void setConfigValues(WtConfiguration config) {
    this.config = config;
    numParasToCheck = (mDocHandler.isTestMode() || mDocHandler.heapLimitIsReached()) ? 0 : config.getNumParasToCheck();
    if (ltMenus != null) {
      ltMenus.setConfigValues(config);
    }
    if (config.noBackgroundCheck() || numParasToCheck == 0) {
      setFlatParagraphTools();
    }
  }

  /**
   * set the document cache - use only for tests
   * @since 5.3
   */
  void setDocumentCacheForTests(List<String> paragraphs, List<List<String>> textParagraphs, List<int[]> footnotes, List<List<Integer>> chapterBegins, Locale locale) {
    try {
      docCache.setForTest(paragraphs, textParagraphs, footnotes, chapterBegins, locale);
      numParasToCheck = -1;
      mDocHandler.getLanguageTool().resetSortedTextRules(mDocHandler.isCheckImpressDocument());
    } catch (Throwable t) {
      //  For tests no messages
    }
  }
  
  /** Get LanguageTool menu
   */
  WtMenus getLtMenu() {
    return ltMenus;
  }
  
  /** Get LanguageTool toolbar
   */
  /*  TODO: in LT 6.5 add dynamic toolbar          
  LtToolbar getLtToolbar() {
    return ltToolbar;
  }
*/  
  /**
   * set menu ID to MultiDocumentsHandler
   */
  void dispose(boolean disposed) {
    this.disposed = disposed;
    if (disposed) {
      if (docCursor != null) {
        docCursor.setDisposed();
      }
//      if (viewCursor != null) {
//        viewCursor.setDisposed();
//      }
      if (flatPara != null) {
        flatPara.setDisposed();
      }
      if (ltMenus != null) {
        ltMenus.removeListener();
        ltMenus = null;
      }
    }
  }
  
  /**
   * get type of document
   */
  public DocumentType getDocumentType() {
    return docType;
  }
  
  /**
   * get number of current paragraph
   */
  public boolean isDisposed() {
    return disposed;
  }
  
  /**
   * A check of the specific paragraph is running
   */
  boolean isRunning(int nFPara) {
    return runningParas.contains(nFPara);
  }
  
  /**
   * set menu ID to MultiDocumentsHandler
   */
  void setMenuDocId() {
    mDocHandler.setMenuDocId(getDocID());
  }
  
  /**
   * get number of current paragraph
   */
  int getCurrentNumberOfParagraph() {
    return paraNum;
  }
  
  /**
   * get language of the document
   */
  public Language getLanguage() {
    Locale locale = docCache.getDocumentLocale();
    if (locale == null) {
      return docLanguage;
    }
    Language lang = WtDocumentsHandler.getLanguage(locale);
    if (lang != null && !lang.equals(docLanguage)) {
      docLanguage = lang;
    }
    return docLanguage;
  }
  
  /**
   * set language of the document
   */
  void setLanguage(Language language) {
    docLanguage = language;
//    docLocale = WtLinguisticServices.getLocale(language);
/*  TODO: in LT 6.5 add dynamic toolbar          
    if (ltToolbar != null) {
      ltToolbar.makeToolbar(language);
    }
*/
  }
  
  /** 
   * Set XComponentContext and XComponent of the document
   */
  void setXComponent(XComponentContext xContext, XComponent xComponent) {
    this.xContext = xContext;
    this.xComponent = xComponent;
    if (xComponent == null) {
      closeDocumentCursor();
//      viewCursor = null;
//      flatPara = null;
    } else {
      setDokumentListener(xComponent);
    }
  }
  
  /**
   *  Get xComponent of the document
   */
  public XComponent getXComponent() {
    return xComponent;
  }
  
  /**
   *  Get MultiDocumentsHandler
   */
  public WtDocumentsHandler getMultiDocumentsHandler() {
    return mDocHandler;
  }
  
  /**
   *  Get ID of the document
   */
  public String getDocID() {
    return docID;
  }
  
  /**
   *  Set ID of the document
   */
  void setDocID(String docId) {
    docID = docId;
  }
  
  /**
   *  Set cache for results of statistical analysis
   */
  public void setStatAnCache(WtResultCache cache) {
    statAnCache = cache;
  }
  
  /**
   *  Set current ruleId for statistical analysis
   */
  public void setStatAnRuleId(String ruleId) {
    this.statAnRuleId = ruleId;
  }
  
  /**
   *  Get flat paragraph tools of the document
   */
  public WtFlatParagraphTools getFlatParagraphTools() {
    if (flatPara == null) {
      setFlatParagraphTools();
    }
    return flatPara;
  }
  
  /**
   *  Get document cursor tools
   */
  public WtDocumentCursorTools getDocumentCursorTools() {
    WtOfficeTools.waitForLO();
    if (docCursor == null) {
      docCursor = new WtDocumentCursorTools(xComponent);
    }
    return docCursor;
  }

  /**
   *  Get document cache of the document
   */
  public List<WtResultCache> getParagraphsCache() {
    return paragraphsCache;
  }
  
  /**
   *  Get document cache of the document
   */
  public WtDocumentCache getDocumentCache() {
    return docCache;
  }
  
  /**
   *  reset document cache of the document
   */
  void resetDocumentCache() {
    resetDocCache = true;
  }
  
  /**
   *  set last changed paragraphs
   */
  void setLastChangedParas(List<Integer> lastChangedParas) {
    this.lastChangedParas = lastChangedParas;
  }
  
  /**
   *  get last changed paragraphs
   */
  List<Integer> getLastChangedParas() {
    return lastChangedParas;
  }
  
  /**
   *  get changed paragraphs map
   */
  Map<Integer, String> getChangedParasMap() {
    return changedParas;
  }
  
  /**
   * reset the Document
   */
  void resetDocument() {
    mDocHandler.resetDocument();
  }
  
   /**
   * read caches from file
   */
  void readCaches() {
    try {
      if (numParasToCheck != 0 && docType != DocumentType.CALC) {
        cacheIO = new WtCacheIO(xComponent);
        boolean cacheExist = cacheIO.readAllCaches(config, mDocHandler);
        if (cacheExist) {
          docCache.put(cacheIO.getDocumentCache());
          for (int i = 0; i < cacheIO.getParagraphsCache().size(); i++) {
  //        if (debugMode > 0) {
            WtMessageHandler.printToLogFile("SingleDocument: readCaches: Copy ResultCache " + i + ": Size: " + cacheIO.getParagraphsCache().get(i).size());
  //        }
            paragraphsCache.get(i).replace(cacheIO.getParagraphsCache().get(i));
          }
          permanentIgnoredMatches = new WtIgnoredMatches(cacheIO.getIgnoredMatches());
          if (docType == DocumentType.WRITER && mDocHandler != null) {
            mDocHandler.runShapeCheck(docCache.hasUnsupportedText(), 9);
          }
        }
        cacheIO.resetAllCache();
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    } finally {
      if (cacheIO != null) {
        cacheIO.resetAllCache();
      }
    }
  }
  
  /**
   * write caches to file
   */
  void writeCaches() {
    try {
      if (numParasToCheck != 0 && !config.noBackgroundCheck() && docType != DocumentType.CALC) {
        if (debugMode > 0) {
          WtMessageHandler.printToLogFile("SingleDocument: writeCaches: Copy DocumentCache");
        }
        WtDocumentCache docCache = new WtDocumentCache(this.docCache);
        List<WtResultCache> paragraphsCache = new ArrayList<WtResultCache>();
        for (int i = 0; i < this.paragraphsCache.size(); i++) {
//          if (debugMode > 0) {
            WtMessageHandler.printToLogFile("SingleDocument: writeCaches: Copy ResultCache " + i + ": Size: " + this.paragraphsCache.get(i).size());
//          }
          paragraphsCache.add(new WtResultCache(this.paragraphsCache.get(i)));
        }
        if (cacheIO != null) {
          WtMessageHandler.printToLogFile("SingleDocument: writeCaches: Save Caches ...");
          cacheIO.saveCaches(docCache, paragraphsCache, permanentIgnoredMatches, config, mDocHandler);
          SpellCache sc = cacheIO.new SpellCache();
          sc.write(WtSpellChecker.getWrongWords(), WtSpellChecker.getSuggestions());
        } else {
          WtMessageHandler.printToLogFile("SingleDocument: writeCaches: cacheIO == null: Can't save cache");
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
  }
  
  /** 
   * Reset all caches of the document
   */
  public void resetResultCache(boolean withSingleParagraph) {
    for (int i = withSingleParagraph ? 0 : 1; i < WtOfficeTools.NUMBER_CACHE; i++) {
      paragraphsCache.get(i).removeAll();
    }
  }
  
  /**
   * remove all cached matches for one paragraph
   */
  public void removeResultCache(int nPara, boolean alsoParaLevel) {
    if (!isDisposed()) {
      if (alsoParaLevel) {
        paragraphsCache.get(0).remove(nPara);
      }
      if (!docCache.setSingleParagraphsCacheToNull(nPara, paragraphsCache)) {
        //  NOTE: Don't remove paragraph cache 0. It is needed to set correct markups
        for (int i = 1; i < paragraphsCache.size(); i++) {
          paragraphsCache.get(i).remove(nPara);
        }
      }
    }
  }
  
  /**
   * Remove a special Proofreading error from all caches of document
   */
  public void removeRuleError(String ruleId) {
    List<Integer> allChanged = new ArrayList<>();
    for (WtResultCache cache : paragraphsCache) {
      List<Integer> changed = cache.removeRuleError(ruleId);
      if (changed.size() > 0) {
        for (int n : changed) {
          if (!allChanged.contains(n)) {
            allChanged.add(n);
          }
        }
      }
    }
    if (allChanged.size() > 0) {
      allChanged.sort(null);
      remarkChangedParagraphs(allChanged, allChanged, true);
    }
  }
  
  /** 
   * Open new flat paragraph tools or initialize them again
   */
  public WtFlatParagraphTools setFlatParagraphTools() {
    try  {
  	  if (disposed) {
        flatPara = null;
        return flatPara;
  	  }
      WtOfficeTools.waitForLO();
  	  if (flatPara == null) {
        flatPara = new WtFlatParagraphTools(xComponent);
        if (!flatPara.isValid()) {
          flatPara = null;
        }
      } else {
        flatPara.init();
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return flatPara;
  }
  
  private void closeDocumentCursor() {
    if (docCursor != null) {
//      docCursor.close();
      docCursor = null;
    }
  }

  /**
   * Add an new entry to text level queue
   * nFPara is number of flat paragraph
   * @throws Throwable 
   */
  public void addQueueEntry(int nFPara, int nCache, int nCheck, String docId, boolean overrideRunning) throws Throwable {
    if (!disposed && mDocHandler.getTextLevelCheckQueue() != null && mDocHandler.getLanguageTool().isSortedRuleForIndex(nCache) && 
        docCache != null && (nCache == 0 || !docCache.isSingleParagraph(nFPara))) {
      boolean checkOnlyParagraph = docCache.isSingleParagraph(nFPara);
      if (nCache > 0 && checkOnlyParagraph) {
        return;
      }
      TextParagraph nTPara = docCache.getNumberOfTextParagraph(nFPara);
      if (nTPara != null && nTPara.type != WtDocumentCache.CURSOR_TYPE_UNKNOWN) {
        int nStart;
        int nEnd;
        if (checkOnlyParagraph) {
          nStart = nTPara.number;
          nEnd = nTPara.number + 1;
        } else {
          nStart = docCache.getStartOfParaCheck(nTPara, nCheck, checkOnlyParagraph, true, false);
          nEnd = docCache.getEndOfParaCheck(nTPara, nCheck, checkOnlyParagraph, true, false);
        }
        mDocHandler.getTextLevelCheckQueue().addQueueEntry(docCache.createTextParagraph(nTPara.type, nStart), 
            docCache.createTextParagraph(nTPara.type, nEnd), nCache, nCheck, docId, overrideRunning);
      }
    }
  }
  
  /**
   * create a queue entry 
   * used by getNextQueueEntry
   */
  QueueEntry createQueueEntry(TextParagraph nPara, int nCache) {
    int nCheck = mDocHandler.getLanguageTool().getNumMinToCheckParas().get(nCache);
    int nStart = docCache.getStartOfParaCheck(nPara, nCheck, false, true, false);
    int nEnd = docCache.getEndOfParaCheck(nPara, nCheck, false, true, false);
    if (nCheck > 0 && nStart + 1 < nEnd) {
      if ((nStart == nPara.number || (nPara.number == 0
              || paragraphsCache.get(nCache).getCacheEntry(docCache.getFlatParagraphNumber(new TextParagraph(nPara.type, nPara.number - 1))) != null)) 
          && (nEnd == nPara.number || nPara.number == docCache.textSize(nPara) - 1
              || paragraphsCache.get(nCache).getCacheEntry(docCache.getFlatParagraphNumber(new TextParagraph(nPara.type, nPara.number + 1))) != null)) {
        nStart = nPara.number;
        nEnd = nStart + 1;
      }
    }
    return mDocHandler.getTextLevelCheckQueue().createQueueEntry(docCache.createTextParagraph(nPara.type, nStart), 
        docCache.createTextParagraph(nPara.type, nEnd), nCache, nCheck, docID, false);
  }

  /**
   * get the next queue entry which is the next empty cache entry
   */
  public QueueEntry getNextQueueEntry(TextParagraph nPara) {
    if (!disposed && docCache != null) {
      if (nPara != null && nPara.type != WtDocumentCache.CURSOR_TYPE_UNKNOWN && nPara.number < docCache.textSize(nPara)
          && !docCache.isSingleParagraph(docCache.getFlatParagraphNumber(nPara))) {
        for (int nCache = 1; nCache < paragraphsCache.size(); nCache++) {
          if (mDocHandler.getLanguageTool().isSortedRuleForIndex(nCache) && docCache.isFinished() 
              && (paragraphsCache.get(nCache).getCacheEntry(docCache.getFlatParagraphNumber(nPara)) == null && 
                  !docCache.isSingleParagraph(docCache.getFlatParagraphNumber(nPara)))) {
            return createQueueEntry(nPara, nCache);
          }
        }
      }
      int nStart = (nPara == null || nPara.type == WtDocumentCache.CURSOR_TYPE_UNKNOWN || nPara.number < docCache.textSize(nPara)) ? 
          0 : docCache.getFlatParagraphNumber(nPara);
      for (int i = nStart; i < docCache.size(); i++) {
        if (docCache.getNumberOfTextParagraph(i).type != WtDocumentCache.CURSOR_TYPE_UNKNOWN && !docCache.isSingleParagraph(i)) {
          for (int nCache = 1; nCache < paragraphsCache.size(); nCache++) {
            if (mDocHandler.getLanguageTool().isSortedRuleForIndex(nCache) && docCache.isFinished() && 
                (paragraphsCache.get(nCache).getCacheEntry(i) == null  && !docCache.isSingleParagraph(i))) {
              return createQueueEntry(docCache.getNumberOfTextParagraph(i), nCache);
            }
          }
        }
      }
      for (int i = 0; i < nStart && i < docCache.size(); i++) {
        if (docCache.getNumberOfTextParagraph(i).type != WtDocumentCache.CURSOR_TYPE_UNKNOWN && !docCache.isSingleParagraph(i)) {
          for (int nCache = 1; nCache < paragraphsCache.size(); nCache++) {
            if (mDocHandler.getLanguageTool().isSortedRuleForIndex(nCache) && docCache.isFinished() && 
                (paragraphsCache.get(nCache).getCacheEntry(i) == null  && !docCache.isSingleParagraph(i))) {
              return createQueueEntry(docCache.getNumberOfTextParagraph(i), nCache);
            }
          }
        }
      }
    }
    return null;
  }

  /**
   * create a queue entry for AI queue
   */
  QueueEntry createAiQueueEntry(int nFPara) {
    TextParagraph nPara = docCache.getNumberOfTextParagraph(nFPara);
    return mDocHandler.getTextLevelCheckQueue().createQueueEntry(nPara, nPara, WtOfficeTools.CACHE_AI, 0, docID, false);
  }
  /**
   * get the next queue entry which is the next empty cache entry
   */
  public QueueEntry getNextAiQueueEntry(TextParagraph nPara) {
    if (!disposed && docCache != null) {
      int nFPara = nPara == null ? 0 : docCache.getFlatParagraphNumber(nPara);
      for (int i = nFPara; i < docCache.size(); i++) {
        if (docCache.isFinished() && i >= 0 && paragraphsCache.get(WtOfficeTools.CACHE_AI).getCacheEntry(i) == null) {
          return createAiQueueEntry(i);
        }
      }
      for (int i = 0; i < nFPara; i++) {
        if (docCache.isFinished() && paragraphsCache.get(WtOfficeTools.CACHE_AI).getCacheEntry(i) == null) {
          return createAiQueueEntry(i);
        }
      }
    }
    return null;
  }

  /**
   * Add an new AI entry to queue
   */
  public void addAiQueueEntry(int nFPara, boolean next) {
    if (!disposed && mDocHandler.getAiCheckQueue() != null && docCache != null) {
      TextParagraph nTPara = docCache.getNumberOfTextParagraph(nFPara);
      if (nTPara != null) {
        mDocHandler.getAiCheckQueue().addQueueEntry(nTPara, docID, next); 
      }
    }
  }
  
  /**
   * get the queue entry for the first changed paragraph in document cache
   */
  public QueueEntry getQueueEntryForChangedParagraph() {
    try {
      if (!disposed && docCache != null && flatPara != null && !changedParas.isEmpty()) {
        Set<Integer> nParas = new HashSet<Integer>(changedParas.keySet());
        if (!nParas.isEmpty()) {
          try {
            Thread.sleep(50);
          } catch (InterruptedException e) {
            WtMessageHandler.printException(e);
          }
        }
        for (int nPara : nParas) {
          WtOfficeTools.waitForLO();
          XFlatParagraph xFlatParagraph = flatPara.getFlatParagraphAt(nPara);
          if (xFlatParagraph != null) {
            String sPara = xFlatParagraph.getText();
            if (sPara != null) {
              if (!isRunning(nPara)) {
                String sChangedPara = changedParas.get(nPara);
                if (sChangedPara != null && (!sChangedPara.equals(sPara)
                    || (mDocHandler.useAnalyzedSentencesCache() && !docCache.isCorrectAnalyzedParagraphLength(nPara, sPara)))) {
                  docCache.setFlatParagraph(nPara, sPara);
  //                removeResultCache(nPara, false);
                  for (int i = 1; i < mDocHandler.getLanguageTool().getNumMinToCheckParas().size(); i++) {
                    addQueueEntry(nPara, i, mDocHandler.getLanguageTool().getNumMinToCheckParas().get(i), docID, false);
                  }
                  if (mDocHandler.useAi()) {
                    addAiQueueEntry(nPara, true);
                  }
                  if (!changedParas.isEmpty()) {
                    addQueueEntry(nPara, 0, 0, docID, false);
                  } else {
                    return createQueueEntry(docCache.getNumberOfTextParagraph(nPara), 0);
                  }
                } else {
                  changedParas.remove(nPara);  // test it as long there is no change
                  List<Integer> changedParas = new ArrayList<>();
                  if (nPara > 0) {
                    changedParas.add(nPara - 1);                                                          
                  }
  //                changedParas.add(nPara);                                                          
                  if (nPara < docCache.size() - 1) {
                    changedParas.add(nPara + 1);                                                          
                  }
                  remarkChangedParagraphs(changedParas, changedParas, false);
                }
              } else {
                try {
                  Thread.sleep(50);
                } catch (InterruptedException e) {
                  WtMessageHandler.printException(e);
                }
              }
            }
          }
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return null;
  }
  
  public void addShapeQueueEntries() throws Throwable {
    int shapeTextSize = docCache.textSize(WtDocumentCache.CURSOR_TYPE_SHAPE) + docCache.textSize(WtDocumentCache.CURSOR_TYPE_TABLE);
    if (shapeTextSize > 0) {
      if (docCursor == null) {
        docCursor = getDocumentCursorTools();
      }
      List<Integer> changedParas = docCache.getChangedUnsupportedParagraphs(docCursor, paragraphsCache.get(0));
      if (changedParas != null) { 
        for (int i = 0; i < changedParas.size(); i++) {
          for (int nCache = 0; nCache < paragraphsCache.size(); nCache++) {
            int nCheck = mDocHandler.getLanguageTool().getNumMinToCheckParas().get(nCache);
            addQueueEntry(changedParas.get(i), nCache, nCheck, docID, true);
          }
        }
      }
    }
  }

  /**
   * run a text level check from a queue entry (initiated by the queue)
   */
  public void runQueueEntry(TextParagraph nStart, TextParagraph nEnd, int cacheNum, int nCheck, boolean override, WtLanguageTool lt) throws Throwable {
    if (!disposed && flatPara != null && docCache.isFinished() && nStart.number < docCache.textSize(nStart)) {
      WtSingleCheck singleCheck = new WtSingleCheck(this, paragraphsCache,
          fixedLanguage, docLanguage, numParasToCheck, false, false, false);
      singleCheck.addParaErrorsToCache(docCache.getFlatParagraphNumber(nStart), lt, cacheNum, nCheck, 
          nEnd.number == nStart.number + 1, override, false, hasFootnotes);
      closeDocumentCursor();
    }
  }
  
  /**
   * set marks for a given list of changed paragraphs
   * before: remove all marks of all paragraphs of a list of paragraphs to remark
   */
  public void remarkChangedParagraphs(List<Integer> changedParas, List<Integer> toRemarkParas, boolean isIntern) {
    try {
      if (!disposed) {
        WtSingleCheck singleCheck = new WtSingleCheck(this, paragraphsCache, fixedLanguage, docLanguage, 
            numParasToCheck, false, false, isIntern);
        singleCheck.remarkChangedParagraphs(changedParas, toRemarkParas, mDocHandler.getLanguageTool());
        closeDocumentCursor();
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }

/**
 * Renew text markups for paragraphs under view cursor
 */
  public void renewMarkups() {
    if (disposed) {
      return;
    }
    WtViewCursorTools viewCursor = new WtViewCursorTools(xComponent);
    int y = docCache.getFlatParagraphNumber(viewCursor.getViewCursorParagraph());
    if (debugMode > 0) {
      WtMessageHandler.printToLogFile("SingleDocument: renewMarkups: Number of Flat Paragraph = " + y);
    }
    List<Integer> changedParas = new ArrayList<Integer>();
    changedParas.add(y);
    remarkChangedParagraphs(changedParas, changedParas, false);
  }

  /**
   * Get an error at position of view cursor
   * test if the range is correct and change it if necessary
   * return null if there is no error or the range is correct
   */
  public WtProofreadingError getErrorAndChangeRange(ContextMenuExecuteEvent aEvent, boolean onlyNonRange) {
    try {
      if (disposed) {
        return null;
      }
      WtViewCursorTools viewCursor = new WtViewCursorTools(xComponent);
      int y = docCache.getFlatParagraphNumber(viewCursor.getViewCursorParagraph());
      int x = viewCursor.getViewCursorCharacter();
      WtProofreadingError error = getErrorFromCache(y, x);
      if (error == null) {
        return null;
      }
//      MessageHandler.printToLogFile("SingleDocument: getErrorAndChangeRange: error at " + x + ": " + error.aRuleIdentifier
//          + ", nStart: " + error.nErrorStart + ", nLength: " + error.nErrorLength);
      XSelectionSupplier xSelectionSupplier = aEvent.Selection;
      Object selection = xSelectionSupplier.getSelection();
      XIndexAccess xIndexAccess = UnoRuntime.queryInterface(XIndexAccess.class, selection);
      if (xIndexAccess == null) {
        WtMessageHandler.printToLogFile("SingleDocument: getErrorAndChangeRange: xIndexAccess == null");
        return null;
      }
      XTextRange xTextRange = UnoRuntime.queryInterface(XTextRange.class, xIndexAccess.getByIndex(0));
      if (xTextRange == null) {
        WtMessageHandler.printToLogFile("SingleDocument: getErrorAndChangeRange: xTextRange == null");
        return null;
      }
      if (onlyNonRange && xTextRange.getString().length() != 0) {
        return null;
      }
      if (xTextRange.getString().length() == error.nErrorLength) {
        return null;
      }
      viewCursor.setViewCursorSelection((short)error.nErrorStart, (short)error.nErrorLength);
      return error;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
    return null;
  }
  
  /**
   * is a ignore once entry in cache
   */
  public boolean isIgnoreOnce(int xFrom, int xTo, int y, String ruleId) {
    return ignoredMatches.isIgnored(xFrom, xTo, y, ruleId);
  }
  
  /**
   * reset the ignore once cache
   */
  public void resetIgnoreOnce() {
    ignoredMatches = new WtIgnoredMatches();
  }
  
  /**
   * add a ignore once entry and remove the mark
   */
  public String ignoreOnce() {
    if (disposed) {
      return null;
    }
    WtViewCursorTools viewCursor = new WtViewCursorTools(xComponent);
    int y = docCache.getFlatParagraphNumber(viewCursor.getViewCursorParagraph());
    int x = viewCursor.getViewCursorCharacter();
    WtProofreadingError error = getErrorFromCache(y, x);
    setIgnoredMatch (x, y, error.aRuleIdentifier, false);
    return docID;
  }
  
  /**
   * add a ignore once entry for point x, y and remove the mark
   */
  public void setIgnoredMatch(int x, int y, String ruleId, boolean isIntern) {
    ignoredMatches.setIgnoredMatch(x, y, 0, ruleId, null, null);
    if (debugMode > 1) {
      WtMessageHandler.printToLogFile("SingleDocument: setIgnoredMatch: DocumentType = " + docType + "; numParasToCheck = " + numParasToCheck);
    }
    if (docType == DocumentType.WRITER && numParasToCheck != 0) {
      List<Integer> changedParas = new ArrayList<>();
      changedParas.add(y);
      remarkChangedParagraphs(changedParas, changedParas, isIntern);
    }
    if (debugMode > 0) {
      WtMessageHandler.printToLogFile("SingleDocument: setIgnoredMatch: Ignore Match added at: paragraph: " + y + "; character: " + x + "; ruleId: " + ruleId);
    }
  }
  
  /**
   * ignore all - call ignoreRule for specified rule
   */
  public void ignoreAll() {
    if (disposed) {
      return;
    }
    WtViewCursorTools viewCursor = new WtViewCursorTools(xComponent);
    int y = docCache.getFlatParagraphNumber(viewCursor.getViewCursorParagraph());
    int x = viewCursor.getViewCursorCharacter();
    WtProofreadingError error = getErrorFromCache(y, x);
    Locale locale = docCache.getFlatParagraphLocale(y);
    mDocHandler.ignoreRule(error.aRuleIdentifier, locale);
  }
  
  /**
   * reset the permanent ignore cache
   */
  public void resetIgnorePermanent() {
    permanentIgnoredMatches.resetAllLocale(getFlatParagraphTools());
    List<Integer> changedParas = permanentIgnoredMatches.getAllParagraphs();
    permanentIgnoredMatches = new WtIgnoredMatches();
    remarkChangedParagraphs(changedParas, changedParas, false);
  }
  
  /**
   * add a ignore once entry to queue and remove the mark
   */
  public String ignorePermanent() {
    if (disposed) {
      return null;
    }
    WtViewCursorTools viewCursor = new WtViewCursorTools(xComponent);
    int y = docCache.getFlatParagraphNumber(viewCursor.getViewCursorParagraph());
    int x = viewCursor.getViewCursorCharacter();
    WtProofreadingError error = getErrorFromCache(y, x);
    Locale locale = error.nErrorType == TextMarkupType.SPELLCHECK ? docCache.getFlatParagraphLocale(y) : null;
    int len = error.nErrorType == TextMarkupType.SPELLCHECK ? error.nErrorLength : 0;
    setPermanentIgnoredMatch(error.nErrorStart, y, len, error.aRuleIdentifier, locale, false);
    return docID;
  }
  
  /**
   * add a ignore once entry for point x, y to queue and remove the mark
   */
  public void setPermanentIgnoredMatch(int x, int y, int len, String ruleId, Locale locale, boolean isIntern) {
    permanentIgnoredMatches.setIgnoredMatch(x, y, len, ruleId, locale, getFlatParagraphTools());
    if (debugMode > 1) {
      WtMessageHandler.printToLogFile("SingleDocument: setPermanentIgnoredMatch: DocumentType = " + docType + "; numParasToCheck = " + numParasToCheck);
    }
    if (docType == DocumentType.WRITER && numParasToCheck != 0) {
      List<Integer> changedParas = new ArrayList<>();
      changedParas.add(y);
      remarkChangedParagraphs(changedParas, changedParas, isIntern);
    }
    if (debugMode > 0) {
      WtMessageHandler.printToLogFile("SingleDocument: setPermanentIgnoredMatch: Ignore Match added at: paragraph: " + y + "; character: " + x + "; ruleId: " + ruleId);
    }
  }
  
  public void setPermanentIgnoredMatches(WtIgnoredMatches ignoredMatches) {
    List<Integer> changedParas = permanentIgnoredMatches.getAllParagraphs();
    permanentIgnoredMatches = ignoredMatches;
    remarkChangedParagraphs(changedParas, changedParas, false);
    changedParas = permanentIgnoredMatches.getAllParagraphs();
    remarkChangedParagraphs(changedParas, changedParas, false);
  }
  
  public WtIgnoredMatches getPermanentIgnoredMatches() {
    return permanentIgnoredMatches;
  }
  
  /**
   * remove all ignore once entries for paragraph y from queue and set the mark
   */
  public void removeAndShiftIgnoredMatch(int from, int to, int oldSize, int newSize) {
    if (!ignoredMatches.isEmpty()) {
      WtIgnoredMatches tmpIgnoredMatches = new WtIgnoredMatches();
      for (int i = 0; i < from; i++) {
        if (ignoredMatches.containsParagraph(i)) {
          tmpIgnoredMatches.put(i, ignoredMatches.get(i));
        }
      }
      for (int i = to + 1; i < oldSize; i++) {
        int n = i + newSize - oldSize;
        if (ignoredMatches.containsParagraph(i)) {
          tmpIgnoredMatches.put(n, ignoredMatches.get(i));
        }
      }
      ignoredMatches = tmpIgnoredMatches;
    }
    if (!permanentIgnoredMatches.isEmpty()) {
      WtIgnoredMatches tmpIgnoredMatches = new WtIgnoredMatches();
      for (int i = 0; i < from; i++) {
        if (permanentIgnoredMatches.containsParagraph(i)) {
          tmpIgnoredMatches.put(i, permanentIgnoredMatches.get(i));
        }
      }
      for (int i = to + 1; i < oldSize; i++) {
        int n = i + newSize - oldSize;
        if (permanentIgnoredMatches.containsParagraph(i)) {
          tmpIgnoredMatches.put(n, permanentIgnoredMatches.get(i));
        }
      }
      permanentIgnoredMatches = tmpIgnoredMatches;
    }
  }
  
  /**
   * remove all ignore once entries for paragraph y from queue and set the mark
   */
  public void removeIgnoredMatch(int y, boolean isIntern) {
    ignoredMatches.removeIgnoredMatches(y, null);
    if (numParasToCheck != 0 && flatPara != null) {
      List<Integer> changedParas = new ArrayList<>();
      changedParas.add(y);
      remarkChangedParagraphs(changedParas, changedParas, isIntern);
    }
    if (debugMode > 0) {
      WtMessageHandler.printToLogFile("SingleDocument: removeIgnoredMatch: All Ignored Matches removed at: paragraph: " + y);
    }
  }
  
  /**
   * remove a ignore once entry for point x, y from queue and set the mark
   * if x &lt; 0 remove all ignore once entries for paragraph y
   */
  public void removeIgnoredMatch(int x, int y, String ruleId, boolean isIntern) {
    ignoredMatches.removeIgnoredMatch(x, y, ruleId, null);
    if (numParasToCheck != 0) {
      List<Integer> changedParas = new ArrayList<>();
      changedParas.add(y);
      remarkChangedParagraphs(changedParas, changedParas, isIntern);
    }
    if (debugMode > 0) {
      WtMessageHandler.printToLogFile("SingleDocument: removeIgnoredMatch: Ignore Match removed at: paragraph: " + y + "; character: " + x);
    }
  }
  
  /**
   * remove a ignore Permanent entry for point x, y from queue and set the mark
   * if x &lt; 0 remove all ignore once entries for paragraph y
   */
  public void removePermanentIgnoredMatch(int x, int y, String ruleId, boolean isIntern) {
    permanentIgnoredMatches.removeIgnoredMatch(x, y, ruleId, getFlatParagraphTools());
    if (numParasToCheck != 0) {
      List<Integer> changedParas = new ArrayList<>();
      changedParas.add(y);
      remarkChangedParagraphs(changedParas, changedParas, isIntern);
    }
    if (debugMode > 0) {
      WtMessageHandler.printToLogFile("SingleDocument: removePermanentIgnoredMatch: Ignore Match removed at: paragraph: " + y + "; character: " + x);
    }
  }
  
  /**
   * get an error out of the cache 
   * by the position of the error (flat paragraph number and number of character)
   */
  private WtProofreadingError getErrorFromCache(int nPara, int nChar) {
    List<WtProofreadingError> tmpErrors = new ArrayList<WtProofreadingError>();
    if (nPara < 0 || nPara >= docCache.size()) {
      WtMessageHandler.printToLogFile("SingleDocument: getRuleIdFromCache(nPara = " + nPara + ", docCache.size() = " + docCache.size() + "): nPara out of range!");
      return null;
    }
    for (WtResultCache paraCache : paragraphsCache) {
      List<WtProofreadingError> tErrors = paraCache.getErrorsAtPosition(nPara, nChar);
      if (tErrors != null) {
        tmpErrors.addAll(tErrors);
      }
    }
    if (tmpErrors.size() > 0) {
      WtProofreadingError[] errors = new WtProofreadingError[tmpErrors.size()];
      for (int i = 0; i < tmpErrors.size(); i++) {
        errors[i] = tmpErrors.get(i);
      }
      Arrays.sort(errors, new WtErrorPositionComparator());
      if (debugMode > 0) {
        for (int i = 0; i < errors.length; i++) {
          WtMessageHandler.printToLogFile("SingleDocument: getRuleIdFromCache: Error[" + i + "]: ruleID: " + errors[i].aRuleIdentifier + ", Start = " + errors[i].nErrorStart + ", Length = " + errors[i].nErrorLength);
        }
      }
      errors = filterOverlappingErrors(errors, config.filterOverlappingMatches());
      for (int i = 0; i < errors.length; i++) {
        if (nChar >= errors[i].nErrorStart && nChar < errors[i].nErrorStart + errors[i].nErrorLength) {
          return errors[i];
        }
      }
      return errors[0];
    } else {
      WtMessageHandler.printToLogFile("SingleDocument: getRuleIdFromCache(nPara = " + nPara + ", nChar = " + nChar + "): No ruleId found!");
      return null;
    }
  }
  
  /**
   * Merge errors from different checks (paragraphs and sentences)
   */
  public WtProofreadingError[] mergeErrors(List<WtProofreadingError[]> pErrors, int nPara) {
    int errorCount = 0;
    if (pErrors != null) {
      for (WtProofreadingError[] pError : pErrors) {
        if (pError != null) {
          errorCount += pError.length;
        }
      }
    }
    if (errorCount == 0 || pErrors == null) {
      return new WtProofreadingError[0];
    }
    WtProofreadingError[] errorArray = new WtProofreadingError[errorCount];
    if (pErrors != null) {
      errorCount = 0;
      for (WtProofreadingError[] pError : pErrors) {
        if (pError != null) {
          arraycopy(pError, 0, errorArray, errorCount, pError.length);
          errorCount += pError.length;
        }
      }
    }
    Arrays.sort(errorArray, new WtErrorPositionComparator());
    return filterIgnoredMatches(errorArray, nPara);
  }
  
  /**
   * Proofs if an error is equivalent
   */
  private boolean isEquivalentError(WtProofreadingError filteredError, WtProofreadingError error) {
    if (filteredError == null || error == null 
        || error.nErrorStart != filteredError.nErrorStart || error.nErrorLength != filteredError.nErrorLength) {
      return false;
    }
    if ((error.aSuggestions == null && filteredError.aSuggestions != null) 
        || (error.aSuggestions != null && filteredError.aSuggestions == null)) {
      return false;
    }
    if (error.aSuggestions == null && filteredError.aSuggestions == null) {
      return true;
    }
    if (error.aSuggestions.length != filteredError.aSuggestions.length) {
      return false;
    }
    for (String suggestion : error.aSuggestions) {
      boolean contents = false;
      for (String fSuggestion : filteredError.aSuggestions) {
        if (suggestion.equals(fSuggestion)) {
          contents = true;
          break;
        }
      }
      if (!contents) {
        return false;
      }
    }
    return true;
  }
  
  /**
   * Filter ignored errors (from ignore once and spell errors)
   */
  private WtProofreadingError[] filterIgnoredMatches (WtProofreadingError[] unFilteredErrors, int nPara) {
    if ((!ignoredMatches.isEmpty() && ignoredMatches.containsParagraph(nPara)) || 
        (!permanentIgnoredMatches.isEmpty() && permanentIgnoredMatches.containsParagraph(nPara))){
      WtProofreadingError lastFilteredError = null;
      List<WtProofreadingError> filteredErrors = new ArrayList<>();
      for (WtProofreadingError error : unFilteredErrors) {
        if (ignoredMatches.isIgnored(error.nErrorStart, error.nErrorStart + error.nErrorLength, nPara, error.aRuleIdentifier) ||
            permanentIgnoredMatches.isIgnored(error.nErrorStart, error.nErrorStart + error.nErrorLength, nPara, error.aRuleIdentifier)) {
          lastFilteredError = error;
        } else if (!isEquivalentError(lastFilteredError, error)) {
          filteredErrors.add(error);
        }
      }
      if (debugMode > 2) {
        WtMessageHandler.printToLogFile("SingleCheck: filterIgnoredMatches: unFilteredErrors.length: " + unFilteredErrors.length);
        WtMessageHandler.printToLogFile("SingleCheck: filterIgnoredMatches: filteredErrors.length: " + filteredErrors.size());
      }
      return filteredErrors.toArray(new WtProofreadingError[0]);
    }
    return unFilteredErrors;
  }

  /**
   * Is an overlapping error
   */
  private boolean isOverlappingError(WtProofreadingError error1, WtProofreadingError error2) {
    return ((error1.nErrorStart >= error2.nErrorStart && error1.nErrorStart < error2.nErrorStart + error2.nErrorLength)
        || (error2.nErrorStart >= error1.nErrorStart && error2.nErrorStart < error1.nErrorStart + error1.nErrorLength));
  }
  
  /**
   * change size of overlapping error1
   */
  private boolean isAiRule(WtProofreadingError error) {
    return error.aRuleIdentifier.equals(WtOfficeTools.AI_GRAMMAR_RULE_ID);
  }
  
  /**
   * Filter overlapping errors
   * Splits overlapping errors
   */
  public WtProofreadingError[] filterOverlappingErrors (WtProofreadingError[] errors, boolean filterOverlap) {
    if (errors == null || errors.length < 2) {
      return errors;
    }
    List<Integer> overlaps = new ArrayList<>();
    for(int i = 0; i < errors.length; i++) {
      WtProofreadingError error1 = errors[i];
      for(int j = 0; j < errors.length; j++) {
        WtProofreadingError error2 = errors[j];
        if (i != j && isOverlappingError(error1, error2)) {
          if (!overlaps.contains(i)) {
            overlaps.add(i);
          }
        }
      }
    }
    if (overlaps.isEmpty()) {
      return errors;
    }
    if (debugMode > 0) {
      WtMessageHandler.printToLogFile("overlaps: " + overlaps.size() + ", filterOverlap: " + filterOverlap);
    }
    List<WtProofreadingError> filteredErrors = new ArrayList<>();
    if (!filterOverlap) {
      for (int i = 0; i < overlaps.size(); i++) {
        int k = overlaps.get(i);
        WtProofreadingError error1 = new WtProofreadingError(errors[k]);
        for(int j = 0; j < errors.length; j++) {
          WtProofreadingError error2 = errors[j];
          if (k != j) {
            if (error2.nErrorStart == error1.nErrorStart && error2.nErrorLength < error1.nErrorLength) {
              int diff = error2.nErrorStart + error2.nErrorLength + 1 - error1.nErrorStart;
              error1.nErrorStart += diff;
              error1.nErrorLength -= diff;
            } else if (error2.nErrorStart > error1.nErrorStart && error2.nErrorStart < error1.nErrorStart + error1.nErrorLength) {
              if (error2.nErrorStart + error2.nErrorLength < error1.nErrorStart + error1.nErrorLength) {
                WtProofreadingError tmpError = new WtProofreadingError(error1);
                error1.nErrorLength = error2.nErrorStart - error1.nErrorStart - 1;
                filteredErrors.add(error1);
                error1 = tmpError;
                int diff = error2.nErrorStart + error2.nErrorLength - error1.nErrorStart;
                error1.nErrorStart += (diff + 1);
                error1.nErrorLength -= diff;
              } else {
                error1.nErrorLength = error2.nErrorStart - error1.nErrorStart - 1;
              }
            }
          }
        }
        filteredErrors.add(error1);
      }
    } else {
      List<Integer> filtered = new ArrayList<>();
      for (int i = 0; i < overlaps.size(); i++) {
        int k = overlaps.get(i);
        if (!filtered.contains(k)) {
          WtProofreadingError error1 = new WtProofreadingError(errors[k]);
          for(int j = 0; j < overlaps.size(); j++) {
            int l = overlaps.get(j);
            WtProofreadingError error2 = errors[l];
            if (k != l && isOverlappingError(error1, error2)) {
              if (!error1.bDefaultRule && error2.bDefaultRule) {
                filtered.add(k);
                error1 = error2;
              } else if (isAiRule(error1) && error2.bDefaultRule && error2.aSuggestions.length > 0) { 
                filtered.add(k);
                error1 = error2;
              } else if (error1.aSuggestions.length < error2.aSuggestions.length) { 
                filtered.add(k);
                error1 = error2;
              } else {
                filtered.add(l);
              }
            }
          }
          filteredErrors.add(error1);
        }
      }
      if (debugMode > 0) {
        for (int i : filtered) {
          WtMessageHandler.printToLogFile("Filtered Rule: " + errors[i].aRuleIdentifier);
        }
      }
    }
    for(int i = 0; i < errors.length; i++) {
      if (!overlaps.contains(i)) {
        filteredErrors.add(errors[i]);
      }
    }
    return filteredErrors.toArray(new WtProofreadingError[0]);
  }

  /**
   * get a rule ID of an error from a check 
   * by the position of the error (number of character)
   */
  private RuleDesc getRuleIdFromCheck(int nChar, WtViewCursorTools viewCursor) {
    String text = viewCursor.getViewCursorParagraphText();
    if (text == null) {
      return null;
    }
    PropertyValue[] propertyValues = new PropertyValue[0];
    ProofreadingResult paRes = new ProofreadingResult();
    paRes.nStartOfSentencePosition = 0;
    paRes.nStartOfNextSentencePosition = 0;
    paRes.nBehindEndOfSentencePosition = paRes.nStartOfNextSentencePosition;
    paRes.xProofreader = null;
    paRes.aLocale = mDocHandler.getLocale();
    paRes.aDocumentIdentifier = docID;
    paRes.aText = text;
    paRes.aProperties = propertyValues;
    paRes.aErrors = null;
    while (nChar > paRes.nStartOfNextSentencePosition && paRes.nStartOfNextSentencePosition < text.length()) {
      paRes.nStartOfSentencePosition = paRes.nStartOfNextSentencePosition;
      paRes.nStartOfNextSentencePosition = text.length();
      paRes.nBehindEndOfSentencePosition = paRes.nStartOfNextSentencePosition;
      paRes = getCheckResults(text, paRes.aLocale, paRes, propertyValues, false, mDocHandler.getLanguageTool(), -1, LoErrorType.GRAMMAR);
      if (paRes.nStartOfNextSentencePosition > nChar) {
        if (paRes.aErrors == null) {
          return null;
        }
        for (SingleProofreadingError error : paRes.aErrors) {
          if (error.nErrorStart <= nChar && nChar < error.nErrorStart + error.nErrorLength) {
            return new RuleDesc(paRes.aLocale, -1, new WtProofreadingError(error));
          }
        }
      }
    }
    WtMessageHandler.printToLogFile("SingleDocument: getRuleIdFromCache: No ruleId found");
    return null;
  }
  
  /**
   * get back the rule ID to deactivate a rule
   */
  public RuleDesc getCurrentRule() {
    if (disposed) {
      return null;
    }
    WtViewCursorTools viewCursor = new WtViewCursorTools(xComponent);
    int x = viewCursor.getViewCursorCharacter();
    if (numParasToCheck == 0) {
      return getRuleIdFromCheck(x, viewCursor);
    }
    int y = docCache.getFlatParagraphNumber(viewCursor.getViewCursorParagraph());
    WtProofreadingError error = getErrorFromCache(y, x);
    if (error != null) {
      return new RuleDesc(docCache.getFlatParagraphLocale(y), y, error);
    }
    return null;
  }

  /**
   * get all synonyms as array
   */
  public String[] getSynonymArray(SingleProofreadingError error, String para, Locale locale, WtLanguageTool lt, boolean setLimit) {
    Map<String, List<String>> synonymMap = getSynonymMap(error, para, locale, lt);
    if (synonymMap.isEmpty()) {
      return new String[0];
    }
    List<String> suggestions = new ArrayList<>();
    int n = 0;
    for (String lemma : synonymMap.keySet()) {
      for (String suggestion : synonymMap.get(lemma)) {
        suggestions.add(suggestion);
        n++;
        if (setLimit && n >= WtOfficeTools.MAX_SUGGESTIONS) {
          break;
        }
      }
      if (setLimit && n >= WtOfficeTools.MAX_SUGGESTIONS) {
        break;
      }
    }
    return suggestions.toArray(new String[suggestions.size()]);
  }
  
  /**
   * get all synonyms as map
   */
  public Map<String, List<String>> getSynonymMap(SingleProofreadingError error, String para, Locale locale, WtLanguageTool lt) {
    Map<String, List<String>> suggestionMap = new HashMap<>();
    try {
      String word = para.substring(error.nErrorStart, error.nErrorStart + error.nErrorLength);
      boolean startUpperCase = Character.isUpperCase(word.charAt(0));
      if (debugMode > 0) {
        WtMessageHandler.printToLogFile("SingleDocument: getSynonymMap: Find Synonyms for word:" + word);
      }
//      List<String> lemmas = lt.getLemmasOfWord(word);
      List<String> lemmas = lt.isRemote() ? lt.getLemmasOfWord(word) : lt.getLemmasOfParagraph(para, error.nErrorStart);
      for (String lemma : lemmas) {
        if (debugMode > 1) {
          WtMessageHandler.printToLogFile("SingleDocument: getSynonymMap: Find Synonyms for lemma:" + lemma);
        }
        List<String> suggestions = new ArrayList<>();
        List<String> synonyms = mDocHandler.getLinguisticServices().getSynonyms(lemma, locale);
        for (String synonym : synonyms) {
          synonym = synonym.replaceAll("\\(.*\\)", "").trim();
          if (debugMode > 1) {
            WtMessageHandler.printToLogFile("SingleDocument: getSynonymMap: Synonym:" + synonym);
          }
          if (!synonym.isEmpty() && !suggestions.contains(synonym)
              && ( (startUpperCase && Character.isUpperCase(synonym.charAt(0))) 
                  || (!startUpperCase && Character.isLowerCase(synonym.charAt(0))))) {
            suggestions.add(synonym);
          }
        }
        if (!suggestions.isEmpty()) {
          suggestionMap.put(lemma, suggestions);
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
    return suggestionMap;
  }

  private void addSynonyms(ProofreadingResult paRes, String para, Locale locale, WtLanguageTool lt) throws IOException {
    WtLinguisticServices linguServices = mDocHandler.getLinguisticServices();
    if (linguServices != null) {
      for (SingleProofreadingError error : paRes.aErrors) {
        if ((error.aSuggestions == null || error.aSuggestions.length == 0) 
            && linguServices.isThesaurusRelevantRule(error.aRuleIdentifier)) {
          error.aSuggestions = getSynonymArray(error, para, locale, lt, true);
        }
      }
    }
  }
  
  /**
   * Get the proofreading result from cache (only if sortedTextId exist)
   * @throws IOException 
   *//*
  ProofreadingResult getErrorsFromCache(int sortedTextId, ProofreadingResult paRes, 
                      String para, Locale locale, SwJLanguageTool lt) throws IOException {
    int nFPara = docCache.getFlatparagraphFromSortedTextId(sortedTextId);
    List<SingleProofreadingError[]> errors = new ArrayList<>();
    paRes.nStartOfSentencePosition = paragraphsCache.get(0).getStartSentencePosition(paraNum, paRes.nStartOfSentencePosition);
    paRes.nStartOfNextSentencePosition = paragraphsCache.get(0).getNextSentencePosition(paraNum, paRes.nStartOfSentencePosition);
    if (paRes.nStartOfNextSentencePosition == 0) {
      paRes.nStartOfNextSentencePosition = para.length();
    }
    paRes.nBehindEndOfSentencePosition = paRes.nStartOfNextSentencePosition;
    for (int cacheNum = 0; cacheNum < mDocHandler.getNumMinToCheckParas().size(); cacheNum++) {
      errors.add(paragraphsCache.get(cacheNum).getFromPara(nFPara, 
              paRes.nStartOfSentencePosition, paRes.nBehindEndOfSentencePosition, LoErrorType.GRAMMAR));
    }
    paRes.aErrors = mergeErrors(errors, nFPara);
//    if (debugMode > 1) {
      MessageHandler.printToLogFile("SingleDocument: getErrorsFromCache: Sentence: start: " + paRes.nStartOfSentencePosition
          + ", end: " + paRes.nBehindEndOfSentencePosition + ", next: " + paRes.nStartOfNextSentencePosition 
          + ", num errors: " + paRes.aErrors.length);
//    }
    addStatAnalysisErrors (paRes, nFPara);
    addSynonyms(paRes, para, locale, lt);
    return paRes;
  }
*/  
  /**
   * Get the proofreading result from cache (only if sortedTextId exist)
   * @throws IOException 
   */
  ProofreadingResult getErrorsFromCache(int nFPara, ProofreadingResult paRes, 
                      String para, Locale locale, WtLanguageTool lt) throws IOException {
    List<WtProofreadingError[]> errors = new ArrayList<>();
    paRes.nStartOfSentencePosition = 0;
    paRes.nStartOfNextSentencePosition = para.length();
    paRes.nBehindEndOfSentencePosition = paRes.nStartOfNextSentencePosition;
    for (int cacheNum = 0; cacheNum < WtOfficeTools.NUMBER_CACHE; cacheNum++) {
      errors.add(paragraphsCache.get(cacheNum).getMatches(nFPara, LoErrorType.GRAMMAR));
    }
    paRes.aErrors = WtOfficeTools.wtErrorsToProofreading(mergeErrors(errors, nFPara));
    if (debugMode > 1) {
      WtMessageHandler.printToLogFile("SingleDocument: getErrorsFromCache: Sentence: start: " + paRes.nStartOfSentencePosition
          + ", end: " + paRes.nBehindEndOfSentencePosition + ", next: " + paRes.nStartOfNextSentencePosition 
          + ", num errors: " + paRes.aErrors.length);
    }
    addStatAnalysisErrors(paRes, nFPara);
    addSynonyms(paRes, para, locale, lt);
    return paRes;
  }
  
/*    !!!  remove after tests   !!!
  private void addSynonyms(ProofreadingResult paRes, String para, Locale locale, SwJLanguageTool lt) throws IOException {
    LinguisticServices linguServices = mDocHandler.getLinguisticServices();
    if (linguServices != null) {
      for (SingleProofreadingError error : paRes.aErrors) {
        if ((error.aSuggestions == null || error.aSuggestions.length == 0) 
            && linguServices.isThesaurusRelevantRule(error.aRuleIdentifier)) {
          String word = para.substring(error.nErrorStart, error.nErrorStart + error.nErrorLength);
          List<String> suggestions = new ArrayList<>();
          List<String> lemmas = lt.getLemmasOfWord(word);
          int num = 0;
          for (String lemma : lemmas) {
            if (debugMode > 0) {
              MessageHandler.printToLogFile("SingleDocument: addSynonyms: Find Synonyms for lemma:" + lemma);
            }
            List<String> synonyms = linguServices.getSynonyms(lemma, locale);
            for (String synonym : synonyms) {
              synonym = synonym.replaceAll("\\(.*\\)", "").trim();
              if (!synonym.isEmpty() && !suggestions.contains(synonym)) {
                suggestions.add(synonym);
                num++;
              }
              if (num >= OfficeTools.MAX_SUGGESTIONS) {
                break;
              }
            }
            if (num >= OfficeTools.MAX_SUGGESTIONS) {
              break;
            }
          }
          if (!suggestions.isEmpty()) {
            error.aSuggestions = suggestions.toArray(new String[suggestions.size()]);
          }
        }
      }
    }
  }
*/
  private void setDokumentListener(XComponent xComponent) {
    try {
      if (!disposed && xComponent != null && eventListener == null) {
        eventListener = new LTDokumentEventListener();
        XDocumentEventBroadcaster broadcaster = UnoRuntime.queryInterface(XDocumentEventBroadcaster.class, xComponent);
        if (!disposed && broadcaster != null) {
          broadcaster.addDocumentEventListener(eventListener);
        } else {
          WtMessageHandler.printToLogFile("SingleDocument: setDokumentListener: Could not add document event listener!");
        }
        XModel xModel = UnoRuntime.queryInterface(XModel.class, xComponent);
        if (disposed || xModel == null) {
          WtMessageHandler.printToLogFile("SingleDocument: setDokumentListener: XModel not found!");
          return;
        }
        XController xController = xModel.getCurrentController();
        if (disposed || xController == null) {
          WtMessageHandler.printToLogFile("SingleDocument: setDokumentListener: XController not found!");
          return;
        }
        XUserInputInterception xUserInputInterception = UnoRuntime.queryInterface(XUserInputInterception.class, xController);
        if (disposed || xUserInputInterception == null) {
          WtMessageHandler.printToLogFile("SingleDocument: setDokumentListener: XUserInputInterception not found!");
          return;
        }
        xUserInputInterception.addMouseClickHandler(eventListener);
//        xUserInputInterception.addKeyHandler(eventListener);
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
  }
  
  public void removeDokumentListener(XComponent xComponent) {
    if (eventListener != null) {
      XDocumentEventBroadcaster broadcaster = UnoRuntime.queryInterface(XDocumentEventBroadcaster.class, xComponent);
      if (broadcaster != null) {
        broadcaster.removeDocumentEventListener(eventListener);
      }
    }
  }
  
  public static class RuleDesc {
    String langCode;
    int nFPara;
    WtProofreadingError error;
    
    RuleDesc(Locale locale, int nFPara, WtProofreadingError error) {
      langCode = WtOfficeTools.localeToString(locale);
      this.nFPara = nFPara;
      this.error = error;
    }
  }

  /**
   * Add statistical analysis errors
   */
  public static WtProofreadingError[] addStatAnalysisErrors (WtProofreadingError[] errors, 
      WtProofreadingError[] statAnErrors, String statAnRuleId) {
    
    List<WtProofreadingError> errorList = new  ArrayList<>();
    for (WtProofreadingError error : statAnErrors) {
      errorList.add(error);
    }
    for (WtProofreadingError error : errors) {
      if (!error.aRuleIdentifier.equals(statAnRuleId)) {
        errorList.add(error);
      }
    }
    return errorList.toArray(new WtProofreadingError[errorList.size()]);
  }
  
  private void addStatAnalysisErrors(ProofreadingResult paRes, int nFPara) {
    if (statAnCache != null && statAnRuleId != null) {
      WtProofreadingError[] statAnErrors = statAnCache.getSafeMatches(nFPara);
      if (statAnErrors != null && statAnErrors.length > 0) {
        paRes.aErrors = WtOfficeTools.wtErrorsToProofreading(
            addStatAnalysisErrors (WtOfficeTools.proofreadingToWtErrors(paRes.aErrors), statAnErrors, statAnRuleId));
      }
    }
  }
  
  
//  private class LTDokumentEventListener implements XDocumentEventListener, XMouseClickHandler, XKeyHandler {
  private class LTDokumentEventListener implements XDocumentEventListener, XMouseClickHandler {

    @Override
    public void disposing(EventObject event) {
    }

    @Override
    public void documentEventOccured(DocumentEvent event) {
//      if(event.EventName.equals("OnUnload")) {
      if(event.EventName.equals("OnPrepareUnload")) {
        try {
          isOnUnload = true;
          writeCaches();
        } catch (Throwable t) {
          WtMessageHandler.printException(t);;
        }
      } else if(event.EventName.equals("OnUnfocus") && !isOnUnload) {
        mDocHandler.getCurrentDocument();
      } else if(event.EventName.equals("OnSave") && config.saveLoCache()) {
          //  save cache before document is saved (xComponent may be null after saving)
        try {
          if (cacheIO != null && xComponent != null) {
            writeCaches();
          }
        } catch (Throwable t) {
          WtMessageHandler.printException(t);;
        }
      } else if(event.EventName.equals("OnSaveAs") && config.saveLoCache()) {
        try {
          if (cacheIO != null && xComponent != null) {
            cacheIO.setDocumentPath(xComponent);
            writeCaches();
          }
        } catch (Throwable t) {
          WtMessageHandler.showError(t);
        }
      }
    }

    @Override
    public boolean mousePressed(MouseEvent event) {
      if (event.Buttons == MouseButton.RIGHT) {
        isRightButtonPressed = true;
      }
      return false;
    }

    @Override
    public boolean mouseReleased(MouseEvent event) {
      return false;
    }
/*
    @Override
    public boolean keyPressed(KeyEvent arg0) {
      return false;
    }

    @Override
    public boolean keyReleased(KeyEvent arg0) {
      MessageHandler.printToLogFile("SingleDocument: setDokumentListener: Set Timestamp");
      OfficeTools.setKeyReleaseTime(System.currentTimeMillis());
      return false;
    }
*/
  }

}

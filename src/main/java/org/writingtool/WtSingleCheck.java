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

import java.awt.Color;
import java.net.URL;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.jetbrains.annotations.Nullable;
import org.languagetool.JLanguageTool;
import org.languagetool.Language;
import org.languagetool.rules.RuleMatch;
import org.languagetool.rules.RuleMatch.Type;
import org.languagetool.tools.StringTools;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.WtResultCache.CacheEntry;
import org.writingtool.aisupport.WtAiDetectionRule;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtDocumentCursorTools;
import org.writingtool.tools.WtFlatParagraphTools;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtOfficeTools.DocumentType;
import org.writingtool.tools.WtOfficeTools.LoErrorType;

import com.sun.star.lang.Locale;
import com.sun.star.lang.XComponent;
import com.sun.star.text.TextMarkupType;
import com.sun.star.text.XParagraphCursor;

/**
 * Class for processing one LO/OO check request
 * Note: There can be some parallel requests from background iteration, dialog, right mouse click or LT text level iteration
 * Gives back the matches found by LT
 * Adds matches to result cache
 * @since 1.0
 * @author Fred Kruse, (including some methods developed by Marcin Miłkowski)
 */
public class WtSingleCheck {
  
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
  
  private static int debugMode;                     //  should be 0 except for testing; 1 = low level; 2 = advanced level

  private final WtSingleDocument singleDocument;      //  handles one document
  private final WtDocumentsHandler mDocHandler;  //  handles the different documents loaded in LO/OO
  private final XComponent xComponent;              //  XComponent of the open document
  private final WtConfiguration config;
  private final WtDocumentCache docCache;             //  cache of paragraphs (only readable by parallel thread)
  private final List<Integer> minToCheckPara;       //  List of minimal to check paragraphs for different classes of text level rules
  private final List<WtResultCache> paragraphsCache;  //  Cache for matches of text rules
  private final int numParasToCheck;                //  current number of Paragraphs to be checked
  private final DocumentType docType;               //  save the type of document
  private final boolean isDialogRequest;            //  true: check was initiated by proofreading dialog
  private final boolean isMouseRequest;             //  true: check was initiated by right mouse click
  private final boolean isIntern;                   //  true: check was initiated by intern check dialog
  private final boolean useQueue;                   //  true: use queue to check text level rules (will be overridden by config)
  private final Language docLanguage;               //  docLanguage (usually the Language of the first paragraph)
  private final Language fixedLanguage;             //  fixed language (by configuration); if null: use language of document (given by LO/OO)

  private int changeFrom = 0;                       //  Change result cache from paragraph
  private int changeTo = 0;                         //  Change result cache to paragraph
  private String lastSinglePara = null;             //  stores the last paragraph which is checked as single paragraph

  private List<Integer> changedParas;               //  List of changed paragraphs after editing the document
  
  WtSingleCheck(WtSingleDocument singleDocument, List<WtResultCache> paragraphsCache,
      Language fixedLanguage, Language docLanguage, 
      int numParasToCheck, boolean isDialogRequest, boolean isMouseRequest, boolean isIntern) {
    debugMode = WtOfficeTools.DEBUG_MODE_SC;
    this.singleDocument = singleDocument;
    this.paragraphsCache = paragraphsCache;
    this.numParasToCheck = numParasToCheck;
    this.isDialogRequest = isDialogRequest;
    this.isMouseRequest = isMouseRequest;
    this.isIntern = isIntern;
    this.docLanguage = docLanguage;
    this.fixedLanguage = fixedLanguage;
    mDocHandler = singleDocument.getMultiDocumentsHandler();
    xComponent = singleDocument.getXComponent();
    docCache = singleDocument.getDocumentCache();
    docType = singleDocument.getDocumentType();
    config = mDocHandler.getConfiguration();
    useQueue = numParasToCheck != 0 && !isDialogRequest && !mDocHandler.isTestMode() && config.useTextLevelQueue();
    minToCheckPara = mDocHandler.getLanguageTool().getNumMinToCheckParas();
    changedParas = new ArrayList<>();
  }
  
  /**
   *   get the result for a check of a single document 
   */
  public WtProofreadingError[] getCheckResults(String paraText, int[] footnotePositions, Locale locale, WtLanguageTool lt, 
      int paraNum, int startOfSentence, boolean textIsChanged, int changeFrom, int changeTo, String lastSinglePara, 
      int lastChangedPara, LoErrorType errType) throws Throwable {
    if (isDisposed()) {
      return new WtProofreadingError[0];
    }
    if (docType == DocumentType.WRITER && !isIntern && lastChangedPara >= 0 && !useQueue) {
//    only for expert mode (not text or paragraph mode)
      List<Integer> changedParas = singleDocument.getLastChangedParas();
      if (changedParas == null) {
        changedParas = new ArrayList<Integer>();
      } else {
        singleDocument.setLastChangedParas(null);
      }
      if (changedParas.contains(lastChangedPara) )
      changedParas.add(lastChangedPara);
      remarkChangedParagraphs(changedParas, changedParas, lt);
    }
    this.lastSinglePara = lastSinglePara;

    List<WtProofreadingError[]> pErrors = checkTextRules(paraText, locale, footnotePositions, paraNum, 
                                                                      startOfSentence, lt, textIsChanged, isIntern, errType);
    if (config.useAiSupport() && config.aiAutoCorrect() && paraNum >= 0) {
      int startSentencePos = paragraphsCache.get(0).getStartSentencePosition(paraNum, startOfSentence);
      int endSentencePos = paragraphsCache.get(0).getNextSentencePosition(paraNum, startOfSentence);
      WtProofreadingError[] aiErrors = paragraphsCache.get(WtOfficeTools.CACHE_AI).getFromPara(paraNum, startSentencePos, endSentencePos, errType);
      if (aiErrors == null) {
        singleDocument.addAiQueueEntry(paraNum, textIsChanged);
      }
      pErrors.add(aiErrors);
    }
    WtProofreadingError[] errors = singleDocument.mergeErrors(pErrors, paraNum, false);
    if (debugMode > 1) {
      WtMessageHandler.printToLogFile("SingleCheck: getCheckResults: paRes.aErrors.length: " + errors.length 
          + "; docID: " + singleDocument.getDocID());
    }
    if (!isDisposed() && docType == DocumentType.WRITER && numParasToCheck != 0 && paraNum >= 0 && !useQueue && (textIsChanged || isDialogRequest)) {
      if (!isIntern && isDialogRequest && !textIsChanged) {
        List<Integer> changedParas = new ArrayList<Integer>();
        changedParas.add(paraNum);
        remarkChangedParagraphs(changedParas, changedParas, lt);
      } else if (textIsChanged && (!useQueue || isDialogRequest)) {
        remarkChangedParagraphs(changedParas, changedParas, lt);
      }
    }
    return errors;
  }
  
  /**
   *   check for number of Paragraphs &gt; 0, chapter wide or full text
   *   is also called by text level queue
   */
  public void addParaErrorsToCache(int nFPara, WtLanguageTool lt, int cacheNum, int parasToCheck, 
        boolean checkOnlyParagraph, boolean override, boolean isIntern, boolean hasFootnotes) throws Throwable {
    //  make the method thread save
    WtDocumentsHandler mDH = mDocHandler;
    if (isDisposed() || docCache == null || nFPara < 0 || nFPara >= docCache.size()) {
      WtMessageHandler.printToLogFile("SingleCheck: addParaErrorsToCache: return: isDisposed = " + isDisposed() + ", nFPara = " + nFPara 
          + ", docCache(Size) = " + (docCache == null ? "null" : docCache.size()) );
      return;
    }
    WtDocumentCache docCache = new WtDocumentCache(this.docCache);
    if (debugMode > 0 && lt == null && !docCache.isAutomaticGenerated(nFPara, true)) {
      WtMessageHandler.printToLogFile("SingleCheck: addParaErrorsToCache: return: lt is null");
    }
    try {

      WtResultCache oldCache = null;
      if (useQueue && !isDialogRequest) {
        oldCache = new WtResultCache(paragraphsCache.get(cacheNum));
      }
      
      TextParagraph tPara = docCache.getNumberOfTextParagraph(nFPara);
      if (tPara.type < 0 || tPara.number < 0) {
        WtMessageHandler.printToLogFile("WARNING: doc cache corrupted (at SingleCheck: addParaErrorsToCache) : refresh doc cache!");
        this.docCache.refresh(singleDocument, WtLinguisticServices.getLocale(fixedLanguage), 
            WtLinguisticServices.getLocale(docLanguage), xComponent, true, false, 7);
        docCache = new WtDocumentCache(this.docCache);
        tPara = docCache.getNumberOfTextParagraph(nFPara);
        if (tPara.type < 0 || tPara.number < 0) {
          WtMessageHandler.printToLogFile("Error: doc cache problem: error cache(" + cacheNum 
              + ") set empty for nFpara = " + nFPara + "!");
          paragraphsCache.get(cacheNum).put(nFPara, null, new WtProofreadingError[0]);
          oldCache = null;
          return;
        }
      }
      int cursorType = tPara.type;
      
      int startPara = docCache.getStartOfParaCheck(tPara, parasToCheck, checkOnlyParagraph, useQueue, false);
      int endPara = docCache.getEndOfParaCheck(tPara, parasToCheck, checkOnlyParagraph, useQueue, false);
      String textToCheck = docCache.getDocAsString(tPara, parasToCheck, checkOnlyParagraph, useQueue, hasFootnotes);
      List<RuleMatch> paragraphMatches = null;
      List<Integer> nextSentencePositions = null;
      //  NOTE: lt == null if language is not supported by LT
      //        but empty proof reading errors have added to cache to satisfy text level queue
      if (lt != null && mDocHandler.getLanguageTool().isSortedRuleForIndex(cacheNum)) {
        if (!docCache.isAutomaticGenerated(nFPara, true)) {
          int startText = docCache.getStartOfParaCheck(tPara, parasToCheck, checkOnlyParagraph, useQueue, true);
          int endText = docCache.getEndOfParaCheck(tPara, parasToCheck, checkOnlyParagraph, useQueue, true);
          paragraphMatches = lt.check(new TextParagraph(tPara.type, startText), new TextParagraph(tPara.type, endText), textToCheck,
              cacheNum == 0 ? JLanguageTool.ParagraphHandling.NORMAL : JLanguageTool.ParagraphHandling.ONLYPARA, singleDocument);
        }
        if (cacheNum == 0) {
          nextSentencePositions = getNextSentencePositions(textToCheck, lt);
        }
      }
      
      int startPos = docCache.getStartOfParagraph(startPara, tPara, parasToCheck, checkOnlyParagraph, useQueue, hasFootnotes);
      int endPos;
      if (debugMode > 1) {
        WtMessageHandler.printToLogFile("SingleCheck: addParaErrorsToCache: tPara.type = " + tPara.type + "; tPara.number = " + tPara.number);
        WtMessageHandler.printToLogFile("SingleCheck: addParaErrorsToCache: nFPara = " + nFPara + "; startPara = " + startPara + "; endPara = " + endPara);
      }
      for (int i = startPara; i < endPara; i++) {
        if (isDisposed() || (useQueue && !isDialogRequest && (mDH.getTextLevelCheckQueue() == null || mDH.getTextLevelCheckQueue().isInterrupted()))) {
          WtMessageHandler.printToLogFile("SingleCheck: addParaErrorsToCache: return: isDisposed = " + isDisposed() + ", useQueue = " + useQueue
              + ", isDialogRequest = " + isDialogRequest + ", TextLevelCheckQueue(isInterrupted) = " 
              + (mDH.getTextLevelCheckQueue() == null ? "null" : mDH.getTextLevelCheckQueue().isInterrupted()));
          oldCache = null;
          return;
        }
        TextParagraph textPara = docCache.createTextParagraph(cursorType, i);
        int[] footnotePos = docCache.getTextParagraphFootnotes(textPara);
        if (i < endPara - 1) {
          endPos = docCache.getStartOfParagraph(i + 1, tPara, parasToCheck, checkOnlyParagraph, useQueue, hasFootnotes);
        } else {
          endPos = textToCheck.length();
        }
        if (paragraphMatches == null || paragraphMatches.isEmpty() || lt == null) {
          paragraphsCache.get(cacheNum).put(docCache.getFlatParagraphNumber(textPara), nextSentencePositions, new WtProofreadingError[0]);
          if (debugMode > 1) {
            WtMessageHandler.printToLogFile("SingleCheck: addParaErrorsToCache: Enter to para cache(" + cacheNum + "): Paragraph(" 
                + docCache.getFlatParagraphNumber(textPara) + "): " + docCache.getTextParagraph(textPara) + "; Error number: 0");
          }
        } else {
          List<WtProofreadingError> errorList = new ArrayList<>();
          int textPos = startPos;
          if (textPos < 0) textPos = 0;
          for (RuleMatch myRuleMatch : paragraphMatches) {
            if (isCorrectRuleMatch(myRuleMatch, textToCheck, lt.getLanguage())) {
              int startErrPos = myRuleMatch.getFromPos();
              if (debugMode > 2) {
                WtMessageHandler.printToLogFile("SingleCheck: addParaErrorsToCache: Cache = " + cacheNum 
                    + ", startPos = " + startPos + ", endPos = " + endPos + ", startErrPos = " + startErrPos);
              }
              if (startErrPos >= startPos && startErrPos < endPos) {
                int toPos = docCache.getTextParagraph(textPara).length();
                if (toPos > 0) {
                  WtProofreadingError error = correctRuleMatchWithFootnotes(
                      createOOoError(myRuleMatch, -textPos, footnotePos, docLanguage, config),
                        footnotePos, docCache.getTextParagraphDeletedCharacters(textPara),
                      docCache.getFlatParagraphNumber(textPara), docCache.getHiddenCharactersMap());
                  errorList.add(error);
                  if (mDocHandler.getAiCheckQueue() != null && mDocHandler.getLinguisticServices() != null 
                      && mDocHandler.getLinguisticServices().isThesaurusRelevantRule(error.aRuleIdentifier)) {
                    if (debugMode > 1) {
                      WtMessageHandler.printToLogFile("WtSingleCheck: addParaErrorsToCache: cacheNum = " + cacheNum
                          + ", nFPara = " + docCache.getFlatParagraphNumber(textPara) + ", error.aRuleIdentifier: " + error.aRuleIdentifier);
                    }
                    singleDocument.addAiQueueEntry(docCache.getFlatParagraphNumber(textPara), cacheNum, error);
                  }
                }
              }
            }
          }
          if (!errorList.isEmpty()) {
            paragraphsCache.get(cacheNum).put(docCache.getFlatParagraphNumber(textPara), nextSentencePositions, errorList.toArray(new WtProofreadingError[0]));
            if (debugMode > 1) {
              WtMessageHandler.printToLogFile("SingleCheck: addParaErrorsToCache: Enter to para cache(" + cacheNum + "): Paragraph(" 
                  + docCache.getFlatParagraphNumber(textPara) + "): " + docCache.getTextParagraph(textPara) 
                  + "; Error number: " + errorList.size());
            }
          } else {
            paragraphsCache.get(cacheNum).put(docCache.getFlatParagraphNumber(textPara), nextSentencePositions, new WtProofreadingError[0]);
            if (debugMode > 1) {
              WtMessageHandler.printToLogFile("SingleCheck: addParaErrorsToCache: Enter to para cache(" + cacheNum + "): Paragraph(" 
                  + docCache.getFlatParagraphNumber(textPara) + "): " + docCache.getTextParagraph(textPara) + "; Error number: 0");
            }
          }
        }
        startPos = endPos;
      }
      if (!isDisposed() && docType == DocumentType.WRITER && useQueue && !isDialogRequest) {
        if (mDH.getTextLevelCheckQueue() == null || mDH.getTextLevelCheckQueue().isInterrupted()) {
          oldCache = null;
          return;
        }
//        flatPara = singleDocument.getFlatParagraphTools();
        
        List<Integer> changedParas = new ArrayList<>();
        List<Integer> toRemarkParas = new ArrayList<>();
        if (cacheNum == 0) {
          changedParas.add(nFPara);
          remarkChangedParagraphs(changedParas, changedParas, lt);
        } else if (oldCache != null) {
          for (int nText = startPara; nText < endPara; nText++) {
            int nFlat = docCache.getFlatParagraphNumber(docCache.createTextParagraph(cursorType, nText));
            if (paragraphsCache.get(0).getCacheEntry(nFlat) != null) {
              if (WtResultCache.areDifferentEntries(paragraphsCache.get(cacheNum).getSerialCacheEntry(nFlat), oldCache.getSerialCacheEntry(nFlat))) {
                changedParas.add(nFlat);
//                if(!ResultCache.isEmptyEntry(oldCache.getSerialCacheEntry(nFlat))) {
                  toRemarkParas.add(nFlat);
//                }
              }
            }
          }
          if (!changedParas.isEmpty()) {
            if (debugMode > 1) {
              WtMessageHandler.printToLogFile("SingleCheck: addParaErrorsToCache: Cache(" + cacheNum + "): Mark paragraphs from " 
                  + startPara + " to " + endPara + ": " + changedParas.size() 
                  + " changes, tPara.type: " + tPara.type + ", tPara.number: " + tPara.number + ", nFPara: " + nFPara);
              String tmpText = "Changed Paras: ";
              for (int n : changedParas) {
                tmpText += n + " ";
              }
              WtMessageHandler.printToLogFile(tmpText);
            }
            singleDocument.setLastChangedParas(changedParas);
            remarkChangedParagraphs(changedParas, toRemarkParas, lt);
          } else if (debugMode > 1) {
            WtMessageHandler.printToLogFile("SingleCheck: addParaErrorsToCache: Cache(" + cacheNum + ") Mark paragraphs from " + startPara 
                + " to " + endPara + ": No Paras to Mark, tPara.type: " + tPara.type + ", tPara.number: " + tPara.number + ", nFPara: " + nFPara);
          }
        }
      }
      oldCache = null;
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }
  
  /** 
   * filter one changed paragraph if sentence or error length exceeds paragraph length
   */
  private void filterOneChangedParasMap (int nPara, int paraLen,
      Map <Integer, List<SentenceErrors>> changedParasMap) throws Throwable {
    for (SentenceErrors errors : changedParasMap.get(nPara)) {
      if (errors.sentenceEnd > paraLen) {
        changedParasMap.remove(nPara);
        return;
      }
    }
    for (WtResultCache cache : paragraphsCache) {
      if (cache.errorLenthExceedsParagraphLength(nPara, paraLen)) {
        changedParasMap.remove(nPara);
        return;
      }
    }
  }
  
  /** 
   * filter marks out of range
   */
  private Map <Integer, List<SentenceErrors>> filterChangedParasMap (Map <Integer, List<SentenceErrors>> changedParasMap, 
      WtDocumentCursorTools docCursor) throws Throwable {
    if (!isDisposed() && docCursor != null) {
      Set<Integer> paras = new HashSet<>(changedParasMap.keySet());
      for (int para : paras) {
        TextParagraph tPara = docCache.getNumberOfTextParagraph(para);
        XParagraphCursor xPara = tPara == null ? null : docCursor.getParagraphCursor(tPara);
        if (xPara != null) {
          int paraLen = xPara.getString().length();
          filterOneChangedParasMap (para, paraLen, changedParasMap);
        }
      }
    }
    return changedParasMap;
  }
  
  /**
   * remark changed paragraphs
   * override existing marks
   */
  public void remarkChangedParagraphs(List<Integer> changedParas, List<Integer> toRemarkParas, 
                                                  WtLanguageTool lt) throws Throwable {
    if (!isDisposed() && !mDocHandler.isBackgroundCheckOff() && (!isDialogRequest || isIntern)) {
      try {
        Map <Integer, List<SentenceErrors>> changedParasMap = new HashMap<>();
        List <TextParagraph> toRemarkTextParas = new ArrayList<>();
        for (int i = 0; i < changedParas.size(); i++) {
          if (!singleDocument.isRunning(i)) {
            List<SentenceErrors> sentencesErrors = getSentencesErrosAsList(changedParas.get(i), 
                lt, LoErrorType.GRAMMAR, config.filterOverlappingMatches());
            changedParasMap.put(changedParas.get(i), sentencesErrors);
            if (debugMode > 1) {
              String message = "SingleCheck: remarkChangedParagraphs: Mark errors: Paragraph: " + changedParas.get(i) 
              + "; Number of sentences: " + sentencesErrors.size();
              for (int j = 0; j < sentencesErrors.size(); j++) {
                message += "; Sentence " + j + ": Start = " + sentencesErrors.get(j).sentenceStart + "; End = " + sentencesErrors.get(j).sentenceEnd 
                            + ", Number of Errors = " + sentencesErrors.get(j).sentenceErrors.length;
              }
              WtMessageHandler.printToLogFile(message);
              message = "SingleCheck: remarkChangedParagraphs: Errors of Sentence 0: ";
              for (int j = 0; j < sentencesErrors.get(0).sentenceErrors.length; j++) {
                message += "Error " + j + ": Start = " + sentencesErrors.get(0).sentenceErrors[j].nErrorStart 
                    + ", Length = " + sentencesErrors.get(0).sentenceErrors[j].nErrorLength 
                    + "; ErrorID = " + sentencesErrors.get(0).sentenceErrors[j].aRuleIdentifier + "; ";
              }
              WtMessageHandler.printToLogFile(message);
              for (int j = 0; j < paragraphsCache.size(); j++) {
                WtMessageHandler.printToLogFile("SingleCheck: remarkChangedParagraphs: Paragraph " + changedParas.get(i) + ": Cache " + j 
                        + ": Number of Errors = " 
                        + (paragraphsCache.get(j).getMatches(changedParas.get(i), LoErrorType.GRAMMAR) == null ? "null" 
                            : paragraphsCache.get(j).getMatches(changedParas.get(i), LoErrorType.GRAMMAR).length));
              }
            }
          }
        }
        for (int i = 0; i < toRemarkParas.size(); i++) {
          if (!(singleDocument).isRunning(i)) {
            toRemarkTextParas.add(docCache.getNumberOfTextParagraph(toRemarkParas.get(i)));
            if (debugMode > 1) {
              String message = "SingleCheck: remarkChangedParagraphs: Remark: Paragraph: " + toRemarkParas.get(i);
              WtMessageHandler.printToLogFile(message);
            }
          }
        }
        WtDocumentCursorTools docCursor = singleDocument.getDocumentCursorTools();
        if (!isDisposed() && !toRemarkTextParas.isEmpty()) {
          if (docCursor != null) {
            docCursor.removeMarks(toRemarkTextParas);
          }
        }
        if (!isDisposed()) {
          WtFlatParagraphTools flatPara = singleDocument.getFlatParagraphTools();
          if (flatPara != null && docCursor != null) {
            flatPara.markParagraphs(filterChangedParasMap(changedParasMap, docCursor));
          }
        }
      } catch (Throwable t) {
        WtMessageHandler.showError(t);
      }
    }
  }
  
  /**
   *  return last single paragraph (not text paragraph)
   */
  public String getLastSingleParagraph () {
    return lastSinglePara;
  }

  /**
   *  Is document disposed?
   */
  private boolean isDisposed() {
    return singleDocument.isDisposed();
  }

    /**
   * check text rules 
   * different caches are supported for check of different number of paragraphs at once 
   * (for different kinds of text level rules)
   */
  private List<WtProofreadingError[]> checkTextRules( String paraText, Locale locale, int[] footnotePos, int paraNum, 
      int startSentencePos, WtLanguageTool lt, boolean textIsChanged, boolean isIntern, LoErrorType errType) throws Throwable {
    List<WtProofreadingError[]> pErrors = new ArrayList<>();
    if (isDisposed()) {
      return pErrors;
    }
    TextParagraph nTParas = paraNum < 0 ? null : docCache.getNumberOfTextParagraph(paraNum);
    if (nTParas == null || nTParas.type == WtDocumentCache.CURSOR_TYPE_UNKNOWN || docCache.isSingleParagraph(paraNum)) {
      pErrors.add(checkParaRules(paraText, locale, footnotePos, paraNum, startSentencePos, lt, 0, 0, textIsChanged, isIntern, errType));
    } else {
      //  Real full text check / numParas < 0
      WtResultCache oldCache = null;
      List<Integer> tmpChangedParas;
      for (int i = 0; i < minToCheckPara.size(); i++) {
        int parasToCheck = minToCheckPara.get(i);
        if (i == 0 || lt.isSortedRuleForIndex(i)) {
          lt.activateTextRulesByIndex(i);
          if (debugMode > 0) {
            WtMessageHandler.printToLogFile("SingleCheck: checkTextRules: Index: " + i + "/" + minToCheckPara.size() 
            + "; paraNum: " + paraNum + "; numParasToCheck: " + parasToCheck + "; useQueue: " + useQueue);
          }
          if (textIsChanged && !useQueue && parasToCheck != 0 ) {
            if (debugMode > 1) {
              WtMessageHandler.printToLogFile("SingleCheck: checkTextRules: Copy old cache!");
            }
            oldCache = new WtResultCache(paragraphsCache.get(i));
            if (debugMode > 1) {
              WtMessageHandler.printToLogFile("SingleCheck: checkTextRules: Old cache copied!");
            }
          }
          pErrors.add(checkParaRules(paraText, locale, footnotePos, paraNum, startSentencePos, lt, i, parasToCheck, textIsChanged, isIntern, errType));
          if (debugMode > 1) {
            WtMessageHandler.printToLogFile("SingleCheck: checkTextRules: Error Cache added!");
          }
          if (!isDisposed() && textIsChanged && !useQueue) {
            if (parasToCheck != 0) {
              tmpChangedParas = paragraphsCache.get(i).differenceInCaches(oldCache);
              for (int chPara : tmpChangedParas) {
                if (!changedParas.contains(chPara)) {
                  changedParas.add(chPara);
                }
              }
              if (!changedParas.contains(paraNum)) {
                changedParas.add(paraNum);
              }
              oldCache = null;
            } else {
              addChangedParas();
            }
          } 
        } else {
          pErrors.add(new WtProofreadingError[0]);
        }
      }
      lt.reactivateTextRules();
      if (debugMode > 1) {
        WtMessageHandler.printToLogFile("SingleCheck: checkTextRules: Text rules reactivated");
      }
    }
    return pErrors;
  }
  
  /**
   * add the numbers of changed paragraphs to list
   */
  private void addChangedParas() {
    int firstPara = changeFrom;
    if (firstPara < 0) {
      firstPara = 0;
    }
    int lastPara = changeTo;
    if (lastPara > docCache.size()) {
      lastPara = docCache.size();
    }
    for (int n = firstPara; n < lastPara; n++) {
      if (!changedParas.contains(n)) {
        changedParas.add(n);
      }
    }
  }

  /**
   * check the text level rules associated with a given cache (cacheNum)
   */
  @Nullable
  public WtProofreadingError[] checkParaRules(String paraText, Locale locale, int[] footnotePos, int nFPara, int sentencePos, 
          WtLanguageTool lt, int cacheNum, int parasToCheck, boolean textIsChanged, 
          boolean isIntern, LoErrorType errType)  throws Throwable {
    List<RuleMatch> paragraphMatches;
    WtProofreadingError[] pErrors = null;
    int startSentencePos = 0;
    int endSentencePos = 0;
    try {
      if (isDisposed()) {
        return pErrors;
      }
      boolean isMultiLingual = nFPara >= 0 ? docCache.isMultilingualFlatParagraph(nFPara) : false;
      // use Cache for check in single paragraph mode only after the first call of paragraph
      if (nFPara >= 0 || (sentencePos > 0 && lastSinglePara != null && lastSinglePara.equals(paraText))) {
        if (paragraphsCache.get(0).getCacheEntry(nFPara) != null) {
          startSentencePos = paragraphsCache.get(0).getStartSentencePosition(nFPara, sentencePos);
          endSentencePos = paragraphsCache.get(0).getNextSentencePosition(nFPara, sentencePos);
          pErrors = paragraphsCache.get(cacheNum).getFromPara(nFPara, startSentencePos, endSentencePos, errType);
          if (debugMode > 0 && pErrors != null) {
            String eInfo = ", ";
            for (WtProofreadingError error : pErrors) {
              eInfo += "(" + error.nErrorStart + "/" + error.nErrorLength + "), ";
            }
            WtMessageHandler.printToLogFile("SingleCheck: checkParaRules: Para: " + nFPara + "; pErrors from cache(" + cacheNum + "): " + pErrors.length
                + ", start = " + startSentencePos + ", end = " + endSentencePos + eInfo);
          }
        }
      } else if (sentencePos == 0) {
        lastSinglePara = paraText;
      }
      // return Cache result if available / for right mouse click or Dialog only use cache
      boolean isTextParagraph = nFPara >= 0 && docCache != null && docCache.getNumberOfTextParagraph(nFPara).type != WtDocumentCache.CURSOR_TYPE_UNKNOWN
          && !docCache.isSingleParagraph(nFPara);
      if (nFPara >= 0 && (pErrors != null || isMouseRequest || (useQueue && !isDialogRequest && parasToCheck != 0))) {
        if (useQueue && pErrors == null && parasToCheck != 0 && isTextParagraph) {
          singleDocument.addQueueEntry(nFPara, cacheNum, parasToCheck, singleDocument.getDocID(), textIsChanged);
        }
        return pErrors;
      }
//      if (runAiQueue && parasToCheck == 0 && nFPara >= 0) {
//        singleDocument.addAiQueueEntry(nFPara);
//      }
      
      //  One paragraph check
      if (!isTextParagraph || parasToCheck == 0) {
        Locale primaryLocale = isMultiLingual ? docCache.getFlatParagraphLocale(nFPara) : locale;
        WtLanguageTool mLt = null;
        if (WtOfficeTools.isEqualLocale(primaryLocale, locale) || !WtDocumentsHandler.hasLocale(primaryLocale)) {
          mLt = lt;
        } else {
          Language mLang = WtDocumentsHandler.getLanguage(primaryLocale);
          if (mLang != null) {
            mLt = mDocHandler.initLanguageTool(mLang);
          }
        }
        List<Integer> nextSentencePositions = getNextSentencePositions(paraText, mLt);
        List<Integer> deletedChars = isTextParagraph ? docCache.getFlatParagraphDeletedCharacters(nFPara): null;
        if (mLt == null || (nFPara >= 0 && docCache != null && docCache.isAutomaticGenerated(nFPara, true))) {
          paragraphMatches = null;
        } else {
          paraText = removeFootnotes(paraText, footnotePos, deletedChars, nFPara, docCache.getHiddenCharactersMap());
          paragraphMatches = mLt.check(paraText, JLanguageTool.ParagraphHandling.NORMAL, nFPara, singleDocument);
        }
        if (isDisposed()) {
          return null;
        }
        if (paragraphMatches == null || paragraphMatches.isEmpty()) {
          paragraphsCache.get(cacheNum).put(nFPara, nextSentencePositions, new WtProofreadingError[0]);
          if (debugMode > 1) {
            WtMessageHandler.printToLogFile("SingleCheck: checkParaRules: Enter " + (isMultiLingual ? "only para " : " ") + "errors to cache(" 
                + cacheNum + "): Paragraph(" + nFPara + "): " + paraText + "; Error number: " + 0);
          }
        } else {
          List<WtProofreadingError> errorList = new ArrayList<>();
          for (RuleMatch myRuleMatch : paragraphMatches) {
            if (isCorrectRuleMatch(myRuleMatch, paraText, lt.getLanguage())) {
              int toPos = myRuleMatch.getToPos();
              if (toPos > paraText.length()) {
                toPos = paraText.length();
              }
              WtProofreadingError error = correctRuleMatchWithFootnotes(
                  createOOoError(myRuleMatch, 0, footnotePos, docLanguage, config), footnotePos, deletedChars,
                  nFPara, docCache.getHiddenCharactersMap());
              errorList.add(error);
              if (mDocHandler.getAiCheckQueue() != null && mDocHandler.getLinguisticServices() != null 
                  && mDocHandler.getLinguisticServices().isThesaurusRelevantRule(error.aRuleIdentifier)) {
                if (debugMode > 1) {
                  WtMessageHandler.printToLogFile("WtSingleCheck: checkParaRules: cacheNum = " + cacheNum
                      + ", nFPara = " + nFPara + ", error.aRuleIdentifier: " + error.aRuleIdentifier);
                }
                singleDocument.addAiQueueEntry(nFPara, cacheNum, error);
              }
            }
          }
          if (!errorList.isEmpty()) {
            if (debugMode > 1) {
              WtMessageHandler.printToLogFile("SingleCheck: checkParaRules: Enter " + (isMultiLingual ? "only para " : "") + "errors to cache(" 
                  + cacheNum + "): Paragraph(" + nFPara + "): " + paraText + "; Error number: " + errorList.size());
            }
            paragraphsCache.get(cacheNum).put(nFPara, nextSentencePositions, errorList.toArray(new WtProofreadingError[0]));
          } else {
            if (debugMode > 1) {
              WtMessageHandler.printToLogFile("SingleCheck: checkParaRules: Enter " + (isMultiLingual ? "only para " : " ") + "errors to cache(" 
                  + cacheNum + "): Paragraph(" + nFPara + "): " + nFPara + "): " + paraText + "; Error number: " + 0);
            }
            paragraphsCache.get(cacheNum).put(nFPara, nextSentencePositions, new WtProofreadingError[0]);
          }
        }
        startSentencePos = paragraphsCache.get(cacheNum).getStartSentencePosition(nFPara, sentencePos);
        endSentencePos = paragraphsCache.get(cacheNum).getNextSentencePosition(nFPara, sentencePos);
        return paragraphsCache.get(cacheNum).getFromPara(nFPara, startSentencePos, endSentencePos, errType);
      }

      //  check of numParasToCheck or full text 
      if (isDisposed()) {
        return null;
      }
      addParaErrorsToCache(nFPara, lt, cacheNum, parasToCheck, (cacheNum == 0), textIsChanged, isIntern, (footnotePos != null));
      return paragraphsCache.get(cacheNum).getFromPara(nFPara, startSentencePos, endSentencePos, errType);

    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return null;
  }
  
  /**
   * is a grammar error or a correct spell error
   */
  private boolean isCorrectRuleMatch(RuleMatch ruleMatch, String text, Language lang) throws Throwable {
    if (!ruleMatch.getRule().isDictionaryBasedSpellingRule()) {
      return true;
    }
    String word;
    if (ruleMatch.getToPos() < text.length()) {
      word = text.substring(ruleMatch.getFromPos(), ruleMatch.getToPos());
    } else {
      word = text.substring(ruleMatch.getFromPos());
    }
    if (!mDocHandler.getLinguisticServices().isCorrectSpell(word, lang)) {
      if (debugMode > 0) {
        WtMessageHandler.printToLogFile("SingleCheck: checkParaRules: not correct spelled word: " + word + "; lang: " + lang.toString());
      }
      return true;
    }
    return false;
  }
  
  /**
   * Creates a SingleGrammarError object for use in LO/OO.
   */
  public static WtProofreadingError createOOoError(RuleMatch ruleMatch, int startIndex, int[] footnotes,
      Language docLanguage, WtConfiguration config) throws Throwable {
    WtProofreadingError aError = new WtProofreadingError();
    if (ruleMatch.getRule().isDictionaryBasedSpellingRule()) {
      aError.nErrorType = TextMarkupType.SPELLCHECK;
    } else {
      aError.nErrorType = TextMarkupType.PROOFREADING;
    }
    if (ruleMatch.getRule().isDefaultOff()) {
      aError.bDefaultRule = false;
    } else {
      aError.bDefaultRule = true;
    }
    
    // the API currently has no support for formatting text in comments
    String msg = ruleMatch.getMessage();
    if (docLanguage != null) {
      msg = docLanguage.toAdvancedTypography(msg);
    }
    msg = msg.replaceAll("<suggestion>", docLanguage == null ? "\"" : docLanguage.getOpeningDoubleQuote())
        .replaceAll("</suggestion>", docLanguage == null ? "\"" : docLanguage.getClosingDoubleQuote())
        .replaceAll("([\r]*\n)", " "); 
    aError.aFullComment = msg;
    // not all rules have short comments
    if (!config.useLongMessages() && !StringTools.isEmpty(ruleMatch.getShortMessage())) {
      aError.aShortComment = ruleMatch.getShortMessage();
    } else {
      aError.aShortComment = aError.aFullComment;
    }
    aError.aShortComment = org.writingtool.tools.WtGeneralTools.shortenComment(aError.aShortComment);
    //  Filter: provide user to delete footnotes by suggestion
    boolean noSuggestions = false;
    if (footnotes != null && footnotes.length > 0 && !ruleMatch.getSuggestedReplacements().isEmpty()) {
      int cor = 0;
      for (int n : footnotes) {
        if (n + cor <= ruleMatch.getFromPos() + startIndex) {
          cor++;
        } else if (n + cor > ruleMatch.getFromPos() + startIndex && n + cor <= ruleMatch.getToPos() + startIndex) {
          noSuggestions = true;
          break;
        }
      }
    }
    int numSuggestions;
    String[] allSuggestions;
    if (noSuggestions) {
      numSuggestions = 0;
      allSuggestions = new String[0];
    } else {
      numSuggestions = ruleMatch.getSuggestedReplacements().size();
      allSuggestions = ruleMatch.getSuggestedReplacements().toArray(new String[numSuggestions]);
    }
    //  Filter: remove suggestions for override dot at the end of sentences
    //  needed because of error in dialog
    /*  since LT 5.2: Filter is commented out because of default use of LT dialog
    if (lastChar == '.' && (ruleMatch.getToPos() + startIndex) == sentencesLength) {
      int i = 0;
      while (i < numSuggestions && i < OfficeTools.MAX_SUGGESTIONS
          && allSuggestions[i].length() > 0 && allSuggestions[i].charAt(allSuggestions[i].length()-1) == '.') {
        i++;
      }
      if (i < numSuggestions && i < OfficeTools.MAX_SUGGESTIONS) {
      numSuggestions = 0;
      allSuggestions = new String[0];
      }
    }
    */
    //  End of Filter
    if (numSuggestions > WtOfficeTools.MAX_SUGGESTIONS) {
      aError.aSuggestions = Arrays.copyOfRange(allSuggestions, 0, WtOfficeTools.MAX_SUGGESTIONS);
    } else {
      aError.aSuggestions = allSuggestions;
    }
    aError.nErrorStart = ruleMatch.getFromPos() + startIndex;
    aError.nErrorLength = ruleMatch.getToPos() - ruleMatch.getFromPos();
    aError.aRuleIdentifier = ruleMatch.getRule().getId();
    // LibreOffice since version 3.5 supports an URL that provides more information about the error,
    // LibreOffice since version 6.2 supports the change of underline color (key: "LineColor", value: int (RGB))
    // LibreOffice since version 6.2 supports the change of underline style (key: "LineType", value: short (DASHED = 5))
    // older version will simply ignore the properties
    String category = ruleMatch.getRule().getCategory().getName();
    aError.bStyleRule = config.isStyleCategory(category) ? true : false;
    aError.bPunctuationRule = ruleMatch.getRule().getCategory().getId().toString().equals("PUNCTUATION") ? true : false;
    String ruleId = ruleMatch.getRule().getId();
    if (ruleMatch.getRule() instanceof WtAiDetectionRule) {
      if (ruleMatch.getType() == Type.UnknownWord) {
        aError.nErrorType = TextMarkupType.SPELLCHECK;
        ruleId = WtOfficeTools.AI_UNKNOWN_WORD_RULE_ID;
      } else if (ruleMatch.getType() == Type.Hint) {
        category = WtOfficeTools.AI_GRAMMAR_CATEGORY;
        ruleId = WtOfficeTools.AI_GRAMMAR_HINT_RULE_ID;
      } else {
        category = WtOfficeTools.AI_STYLE_CATEGORY;
        ruleId = WtOfficeTools.AI_GRAMMAR_OTHER_RULE_ID;
      }
    }
    Color underlineColor = aError.nErrorType != TextMarkupType.SPELLCHECK 
                            ? config.getUnderlineColor(category, ruleId) : Color.red;
    short underlineType = aError.nErrorType != TextMarkupType.SPELLCHECK 
                            ? config.getUnderlineType(category, ruleId) : WtConfiguration.UNDERLINE_WAVE;
    URL url = ruleMatch.getUrl();
    if (url == null) {                      // match URL overrides rule URL 
      url = ruleMatch.getRule().getUrl();
    }
    int nDim = 0;
    if (url != null) {
      nDim++;
    }
    if (underlineColor != Color.blue || aError.nErrorType == TextMarkupType.SPELLCHECK) {
      nDim++;
    }
    if (underlineType != WtConfiguration.UNDERLINE_WAVE || (config.markSingleCharBold() && aError.nErrorLength == 1)) {
      nDim++;
    }
    if (nDim > 0) {
      //  HINT: Because of result cache handling:
      //  handle should always be -1
      //  property state should always be PropertyState.DIRECT_VALUE
      //  otherwise result cache handling has to be adapted
      WtPropertyValue[] propertyValues = new WtPropertyValue[nDim];
      int n = 0;
      if (url != null) {
        propertyValues[n] = new WtPropertyValue("FullCommentURL", url.toString());
        n++;
      }
      if (aError.nErrorType == TextMarkupType.SPELLCHECK) {
        int ucolor = Color.red.getRGB() & 0xFFFFFF;
        propertyValues[n] = new WtPropertyValue("LineColor", ucolor);
        n++;
/*        
      } else if (ruleMatch.getRule() instanceof WtAiDetectionRule) {
        int ucolor;
        if (ruleMatch.getType() == Type.Hint) {
          ucolor = WtAiDetectionRule.RULE_HINT_COLOR.getRGB() & 0xFFFFFF;
        } else {
          ucolor = WtAiDetectionRule.RULE_OTHER_COLOR.getRGB() & 0xFFFFFF;
        }
        propertyValues[n] = new PropertyValue("LineColor", -1, ucolor, PropertyState.DIRECT_VALUE);
        n++;
*/
      } else {
        if (underlineColor != Color.blue) {
          int ucolor = underlineColor.getRGB() & 0xFFFFFF;
          propertyValues[n] = new WtPropertyValue("LineColor", ucolor);
          n++;
        }
      }
      if (underlineType != WtConfiguration.UNDERLINE_WAVE) {
        propertyValues[n] = new WtPropertyValue("LineType", underlineType);
      } else if (config.markSingleCharBold() && aError.nErrorLength == 1) {
        propertyValues[n] = new WtPropertyValue("LineType", WtConfiguration.UNDERLINE_BOLDWAVE);
      }
      aError.aProperties = propertyValues;
    } else {
        aError.aProperties = new WtPropertyValue[0];
    }
    return aError;
  }
  
  /**
   * get beginning of next sentence using LanguageTool tokenization
   */
  public static List<Integer> getNextSentencePositions (String paraText, WtLanguageTool lt) {
    List<Integer> nextSentencePositions = new ArrayList<Integer>();
    if (paraText == null || paraText.isEmpty()) {
      nextSentencePositions.add(0);
      return nextSentencePositions;
    }
//    if (lt == null || lt.isRemote()) {
    if (lt == null) {
      nextSentencePositions.add(paraText.length());
    } else {
      List<String> tokenizedSentences = lt.sentenceTokenize(cleanFootnotes(paraText));
      int position = 0;
      for (String sentence : tokenizedSentences) {
        position += sentence.length();
        nextSentencePositions.add(position);
      }
      if (nextSentencePositions.get(nextSentencePositions.size() - 1) != paraText.length()) {
        nextSentencePositions.set(nextSentencePositions.size() - 1, paraText.length());
      }
    }
    return nextSentencePositions;
  }
  
  /**
   * Fix numbers that are (probably) foot notes.
   * See https://bugs.freedesktop.org/show_bug.cgi?id=69416
   * public for test reasons
   */
  static String cleanFootnotes(String paraText) {
    return paraText.replaceAll("([^\\d][.!?])\\d ", "$1¹ ");
  }
  
  /**
   *  get hidden characters as list
   */
  public static List<Integer> getHiddenCharactersAsList(String paraText) {
    int i = paraText.indexOf(WtOfficeTools.SOFT_HYPHEN);
    if (i < 0) {
      return null;
    }
    List<Integer> hiddenCharacter = new ArrayList<>();
    hiddenCharacter.add(i);
    while (i > 0) {
      i = paraText.indexOf(WtOfficeTools.SOFT_HYPHEN, i + 1);
      if (i >= 0) {
        hiddenCharacter.add(i);
      }
    }
    return hiddenCharacter;
  }
  
  /**
   * merge Footnotes, deleted and hidden characters to one array
   */
  public static int[] mergeFootnotesEtc(int[] footnotes, List<Integer> deletedChars, List<Integer> hiddenCharacters) throws Throwable {
    if ((deletedChars == null || deletedChars.isEmpty()) && (hiddenCharacters == null || hiddenCharacters.isEmpty())) {
      return footnotes;
    }
    int n = 0;
    if (footnotes != null) {
      n += footnotes.length;
    }
    if (deletedChars != null) {
      n += deletedChars.size();
    }
    if (hiddenCharacters != null) {
      n += hiddenCharacters.size();
    }
    int[] allChars = new int[n];
    n = 0;
    if (footnotes != null) {
      for (int i = 0; i < footnotes.length; i++, n++) {
        allChars[n] = footnotes[i];
      }
    }
    if (deletedChars != null) {
      for (int i = 0; i < deletedChars.size(); i++, n++) {
        allChars[n] = deletedChars.get(i);
      }
    }
    if (hiddenCharacters != null) {
      for (int i = 0; i < hiddenCharacters.size(); i++, n++) {
        allChars[n] = hiddenCharacters.get(i);
      }
    }
    Arrays.sort(allChars);
    return allChars;
  }

  
  /**
   * Remove footnotes, deleted characters and hidden characters from paraText
   * run cleanFootnotes if information about footnotes are not supported
   */
  public static String removeFootnotes(String paraText, int[] footnotes, List<Integer> deletedChars,
      int nPara, Map<Integer, List<Integer>> hiddenCharacters) throws Throwable {
    if (paraText == null) {
      return null;
    }
    if (footnotes == null) {
      return cleanFootnotes(paraText);
    }
    List<Integer> hiddenCharacterList = getHiddenCharactersAsList(paraText);
    if (hiddenCharacterList != null) {
      hiddenCharacters.put(nPara, hiddenCharacterList);
    } else {
      hiddenCharacters.remove(nPara);
    }
//    WtMessageHandler.printToLogFile("removeFootnotes: paraText.length: " + paraText.length() + ", footnotes.length: " + footnotes.length);
    int[] allChars = mergeFootnotesEtc(footnotes, deletedChars, hiddenCharacterList);
    for (int i = allChars.length - 1; i >= 0; i--) {
      if (allChars[i] < paraText.length()) {
        paraText = paraText.substring(0, allChars[i]) + paraText.substring(allChars[i] + 1);
      }
    }
//    WtMessageHandler.printToLogFile("removeFootnotes (new): paraText.length: " + paraText.length() + ", allChars.length: " + allChars.length);
    return paraText;
  }

/*
  public static String removeFootnotes(String paraText, int[] footnotes, List<Integer> deletedChars,
      int nPara, Map<Integer, List<Integer>> hiddenCharacters) throws Throwable {
    if (paraText == null) {
      return null;
    }
    if (deletedChars == null || deletedChars.isEmpty()) {
      if (footnotes == null) {
        return cleanFootnotes(paraText);
      }
      for (int i = footnotes.length - 1; i >= 0; i--) {
        if (footnotes[i] < paraText.length()) {
          paraText = paraText.substring(0, footnotes[i]) + paraText.substring(footnotes[i] + 1);
        }
      }
    } else {
      if (footnotes == null || footnotes.length == 0) {
        if (footnotes == null) {
          paraText = cleanFootnotes(paraText);
        }
        for (int i = deletedChars.size() - 1; i >= 0; i--) {
          if (deletedChars.get(i) < paraText.length()) {
            paraText = paraText.substring(0, deletedChars.get(i)) + paraText.substring(deletedChars.get(i) + 1);
          }
        }
      } else {
        int idc = deletedChars.size() - 1;
        int ifn = footnotes.length - 1;
        while (idc >= 0 || ifn >= 0) {
          if (idc >= 0 && (ifn < 0 || deletedChars.get(idc) >= footnotes[ifn])) {
            if (deletedChars.get(idc) < paraText.length()) {
              paraText = paraText.substring(0, deletedChars.get(idc)) + paraText.substring(deletedChars.get(idc) + 1);
            }
            if (ifn >= 0 && deletedChars.get(idc) == footnotes[ifn]) {
              ifn--;
            }
            idc--;
          } else {
            if (footnotes[ifn] < paraText.length()) {
              paraText = paraText.substring(0, footnotes[ifn]) + paraText.substring(footnotes[ifn] + 1);
            }
            ifn--;
          }
        }
      }
    }
    List<Integer> hiddenCharacterList = getHiddenCharactersAsList(paraText);
    if (hiddenCharacterList != null) {
      hiddenCharacters.put(nPara, hiddenCharacterList);
      paraText = paraText.replace(WtOfficeTools.SOFT_HYPHEN, "");
    } else {
      hiddenCharacters.remove(nPara);
    }
    return paraText;
  }
*/  
  /**
   * Correct WtProofreadingError by footnote positions
   * footnotes before is the sum of all footnotes before the checked paragraph
   */
  public static WtProofreadingError correctRuleMatchWithFootnotes(WtProofreadingError pError, int[] footnotes, List<Integer> deletedChars,
      int nPara, Map<Integer, List<Integer>> hiddenCharacters) throws Throwable {
    if (footnotes == null) {
      return pError;
    }
    List<Integer> hiddenCharacterList = hiddenCharacters.get(nPara);
    int[] allChars = mergeFootnotesEtc(footnotes, deletedChars, hiddenCharacterList);
    for (int i :allChars) {
      if (i <= pError.nErrorStart) {
        pError.nErrorStart++;
      } else if (i < pError.nErrorStart + pError.nErrorLength) {
        pError.nErrorLength++;
      }
    }
    return pError;
  }
/*
  public static WtProofreadingError correctRuleMatchWithFootnotes(WtProofreadingError pError, int[] footnotes, List<Integer> deletedChars,
      int nPara, Map<Integer, List<Integer>> hiddenCharacters) throws Throwable {
    List<Integer> hiddenCharactersList = hiddenCharacters.get(nPara);
    if (hiddenCharactersList != null) {
      for (int i : hiddenCharactersList) {
        if (i <= pError.nErrorStart) {
          pError.nErrorStart++;
        } else if (i < pError.nErrorStart + pError.nErrorLength) {
          pError.nErrorLength++;
        }
      }
    }
    if (deletedChars == null || deletedChars.isEmpty()) {
      if (footnotes == null || footnotes.length == 0) {
        return pError;
      }
      for (int i :footnotes) {
        if (i <= pError.nErrorStart) {
          pError.nErrorStart++;
        } else if (i < pError.nErrorStart + pError.nErrorLength) {
          pError.nErrorLength++;
        }
      }
    } else {
      if (footnotes == null || footnotes.length == 0) {
        for (int i : deletedChars) {
          if (i <= pError.nErrorStart) {
            pError.nErrorStart++;
          } else if (i < pError.nErrorStart + pError.nErrorLength) {
            pError.nErrorLength++;
          }
        }
      } else {
        int ifn = 0;
        int idc = 0;
        while (ifn < footnotes.length || idc < deletedChars.size()) {
          if (idc < deletedChars.size() && (ifn >= footnotes.length || deletedChars.get(idc) < footnotes[ifn])) {
            if (deletedChars.get(idc) <= pError.nErrorStart) {
              pError.nErrorStart++;
            } else if (deletedChars.get(idc) < pError.nErrorStart + pError.nErrorLength) {
              pError.nErrorLength++;
            }
            if (ifn < footnotes.length && deletedChars.get(idc) == footnotes[ifn]) {
              ifn++;
            }
            idc++;
          } else {
            if (footnotes[ifn] <= pError.nErrorStart) {
              pError.nErrorStart++;
            } else if (footnotes[ifn] < pError.nErrorStart + pError.nErrorLength) {
              pError.nErrorLength++;
            }
            ifn++;
          }
        }
      }
    }
    return pError;
  }
*/  
  /**
   * get all errors of a Paragraph as list
   */
  private List<SentenceErrors> getSentencesErrosAsList(int numberOfParagraph, WtLanguageTool lt, 
            LoErrorType errType, boolean filterOverlap) {
    List<SentenceErrors> sentenceErrors = new ArrayList<SentenceErrors>();
    if (!isDisposed()) {
      CacheEntry entry = paragraphsCache.get(0).getCacheEntry(numberOfParagraph);
      List<Integer> nextSentencePositions = null;
      if (entry != null) {
        nextSentencePositions = entry.nextSentencePositions;
      }
      if (nextSentencePositions == null) {
        nextSentencePositions = new ArrayList<Integer>();
      }
      if (nextSentencePositions.size() == 0 && docCache != null 
          && numberOfParagraph >= 0 && numberOfParagraph < docCache.size()) {
        nextSentencePositions = getNextSentencePositions(docCache.getFlatParagraph(numberOfParagraph), lt);
      }
      int startPosition = 0;
      if (nextSentencePositions.size() == 1) {
        List<WtProofreadingError[]> errorList = new ArrayList<WtProofreadingError[]>();
        for (WtResultCache cache : paragraphsCache) {
          errorList.add(cache.getMatches(numberOfParagraph, errType));
        }
        sentenceErrors.add(new SentenceErrors(startPosition, nextSentencePositions.get(0), singleDocument.mergeErrors(errorList, numberOfParagraph, false)));
      } else {
        for (int nextPosition : nextSentencePositions) {
          List<WtProofreadingError[]> errorList = new ArrayList<WtProofreadingError[]>();
          for (WtResultCache cache : paragraphsCache) {
            errorList.add(cache.getFromPara(numberOfParagraph, startPosition, nextPosition, errType));
          }
          sentenceErrors.add(new SentenceErrors(startPosition, nextPosition, singleDocument.mergeErrors(errorList, numberOfParagraph, false)));
          startPosition = nextPosition;
        }
      }
    }
    return sentenceErrors;
  }

  /**
   * Class of proofreading errors of one sentence
   */
  public static class SentenceErrors {
    public final int sentenceStart;
    public final int sentenceEnd;
    public final WtProofreadingError[] sentenceErrors;
    
    public SentenceErrors(int start, int end, WtProofreadingError[] errors) {
      sentenceStart = start;
      sentenceEnd = end;
      sentenceErrors = errors;
    }
  }

  
}

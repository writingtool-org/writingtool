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
import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.locks.ReentrantReadWriteLock;

import org.languagetool.AnalyzedSentence;
import org.languagetool.AnalyzedToken;
import org.languagetool.AnalyzedTokenReadings;
import org.languagetool.JLanguageTool;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtDocumentCursorTools;
import org.writingtool.tools.WtFlatParagraphTools;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeDrawTools;
import org.writingtool.tools.WtOfficeSpreadsheetTools;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtDocumentCursorTools.DocumentText;
import org.writingtool.tools.WtFlatParagraphTools.FlatParagraphContainer;
import org.writingtool.tools.WtOfficeDrawTools.ParagraphContainer;
import org.writingtool.tools.WtOfficeTools.DocumentType;
import org.writingtool.tools.WtVersionInfo;

import com.sun.star.lang.Locale;
import com.sun.star.lang.XComponent;

/**
 * Class to store the Text WtProofreadingErrorent cache)
 * 
 * @since 1.0
 * @author Fred Kruse
 */
public class WtDocumentCache implements Serializable {

  private static final long serialVersionUID = 13L;

  public final static int CURSOR_TYPE_UNKNOWN = -1;
  public final static int CURSOR_TYPE_ENDNOTE = 0;
  public final static int CURSOR_TYPE_FOOTNOTE = 1;
  public final static int CURSOR_TYPE_HEADER_FOOTER = 2;
  public final static int CURSOR_TYPE_SHAPE = 3;
  public final static int CURSOR_TYPE_TEXT = 4;
  public final static int CURSOR_TYPE_TABLE = 5;

  public static final int NUMBER_CURSOR_TYPES = 6;
  
  private static final int MAX_NOTE_CHAR = 7;       //  supports Roman numbers to 87
  private static final int MAX_PRINTED_PARAS = 3;   //  maximal printed paragraphs to log file
  private static final int WAIT_TIME = 20;          //  time to wait before update of paragraph cache

  private static boolean debugMode;     // should be false except for testing
  private static boolean debugModeTm;   // time measurement should be false except for testing


  private final List<String> paragraphs = new ArrayList<String>(); // stores the flat paratoTextMappinggraphs of
                                                                   // document

  private final List<List<Integer>> chapterBegins = new ArrayList<List<Integer>>(); // stores the paragraphs formatted as
                                                                                    // headings; is used to subdivide
                                                                                    // the document in chapters
  private final List<Integer> automaticParagraphs = new ArrayList<Integer>(); // stores the paragraphs automatic generated (will not be checked)
  private final List<SerialLocale> locales = new ArrayList<SerialLocale>(); // stores the language of the paragraphs;
  private final List<int[]> footnotes = new ArrayList<int[]>();             // stores the footnotes of the paragraphs;
  private final List<List<Integer>> deletedCharacters = new ArrayList<List<Integer>>(); // stores the deleted characters (report changes) of the paragraphs;
  private final List<List<Integer>> openingQuotes = new ArrayList<List<Integer>>(); // stores the opening quotes in a paragraph;
  private final List<List<Integer>> closingQuotes = new ArrayList<List<Integer>>(); // stores the closing quotes in a paragraph;
  private final List<TextParagraph> toTextMapping = new ArrayList<>(); // Mapping from FlatParagraph to DocumentCursor
  protected final List<List<Integer>> toParaMapping = new ArrayList<>(); // Mapping from DocumentCursor to FlatParagraph
  private final DocumentType docType;                 // stores the document type (Writer, Impress, Calc)
  private final Map<Integer, List<AnalyzedSentence>> analyzedParagraphs = new HashMap<>();  //  stores analyzed paragraphs
  private List<Integer> sortedTextIds = null;           // stores the node index of the paragraphs (since LO 7.5 / else null)
  private Map<Integer, Integer> headingMap;
  private boolean isReset = false;
  private boolean isDirty = false;
  private int documentElementsCount = -1;
  private int nEndnote = 0;
  private int nFootnote = 0;
  private int nHeaderFooter = 0;
  private int nShape = 0;
  private int nText = 0;
  private int nTable = 0;
  private SerialLocale docLocale; 
  
  private ReentrantReadWriteLock rwLock = new ReentrantReadWriteLock();

  WtDocumentCache(DocumentType docType) {
    debugMode = WtOfficeTools.DEBUG_MODE_DC;
    debugModeTm = WtOfficeTools.DEBUG_MODE_TM;
    this.docType = docType;
  }

  WtDocumentCache(WtSingleDocument document, Locale fixedLocale, Locale docLocale,
      XComponent xComponent, DocumentType docType) {
    debugMode = WtOfficeTools.DEBUG_MODE_DC;
    debugModeTm = WtOfficeTools.DEBUG_MODE_TM;
    this.docType = docType;
    refresh(document, fixedLocale, docLocale, xComponent, false, 0);
  }

  public WtDocumentCache(WtDocumentCache in) {
    rwLock.writeLock().lock();
    in.rwLock.readLock().lock();
    try {
      isReset = true;
      debugMode = WtOfficeTools.DEBUG_MODE_DC;
      debugModeTm = WtOfficeTools.DEBUG_MODE_TM;
      if (in.paragraphs != null && in.paragraphs.size() > 0) {
        add(in);
      }
      docType = in.docType;
      docLocale = getMostUsedLanguage(locales);
    } finally {
      isReset = false;
      in.rwLock.readLock().unlock();
      rwLock.writeLock().unlock();
    }
  }

  /**
   * set the cache only for test use
   */
  public void setForTest(List<String> paragraphs, List<List<String>> textParagraphs, List<int[]> footnotes,
      List<List<Integer>> chapterBegins, Locale locale) {
    rwLock.writeLock().lock();
    try {
      isReset = true;
      debugMode = WtOfficeTools.DEBUG_MODE_DC;
      debugModeTm = WtOfficeTools.DEBUG_MODE_TM;
      clearAnalyzedParagraphs();
      this.paragraphs.addAll(paragraphs);
      this.footnotes.addAll(footnotes);
      this.chapterBegins.addAll(chapterBegins);
      for (int i = 0; i < paragraphs.size(); i++) {
        locales.add(new SerialLocale(locale));
      }
      for (int i = 0; i < NUMBER_CURSOR_TYPES; i++) {
        toParaMapping.add(new ArrayList<Integer>());
      }
      nText = textParagraphs.get(CURSOR_TYPE_TEXT).size();
      nTable = textParagraphs.get(CURSOR_TYPE_TABLE).size();
      nShape = textParagraphs.get(CURSOR_TYPE_SHAPE).size();
      nFootnote = textParagraphs.get(CURSOR_TYPE_FOOTNOTE).size();
      nEndnote = textParagraphs.get(CURSOR_TYPE_ENDNOTE).size();
      nHeaderFooter = textParagraphs.get(CURSOR_TYPE_HEADER_FOOTER).size();
      docLocale = new SerialLocale(locale);
      mapParagraphs(this.paragraphs, toTextMapping, toParaMapping, this.chapterBegins, locales, footnotes, textParagraphs, deletedCharacters, null);
    } finally {
      isReset = false;
      rwLock.writeLock().unlock();
    }
  }
  
  /**
   * Refresh the cache
   */
  public void refresh(WtSingleDocument document, Locale fixedLocale, Locale docLocale, XComponent xComponent, boolean wait, int fromWhere) {
    if (isReset) {
      WtMessageHandler.printToLogFile("DocumentCache:refresh: isReset == true: return");
      return;
    }
    isReset = true;
    try {
      if (debugMode) {
        WtMessageHandler.printToLogFile("DocumentCache: refresh: Called from: " + fromWhere);
      }
      clearAnalyzedParagraphs();
      if (docType != DocumentType.WRITER) {
        refreshImpressCalcCache(xComponent);
      } else {
        if (wait) {
          WtOfficeTools.sleep(WAIT_TIME);
        }
        refreshWriterCache(document, fixedLocale, docLocale, fromWhere);
      }
      setSingleParagraphsCacheToNull(document.getParagraphsCache());
    } finally {
      isReset = false;
    }
  }

  /**
   * reset the document cache load the actual state of the document into the cache
   * is only used for writer documents
   */
  private void refreshWriterCache(WtSingleDocument document, Locale fixedLocale, Locale docLocale, int fromWhere) {
    try {
      long startTime = System.currentTimeMillis();
      WtFlatParagraphTools flatPara = document.getFlatParagraphTools();
      List<String> paragraphs = new ArrayList<String>();
      List<List<Integer>> chapterBegins = new ArrayList<List<Integer>>();
      List<List<Integer>> deletedCharacters = new ArrayList<List<Integer>>();
      List<SerialLocale> locales = new ArrayList<SerialLocale>();
      List<int[]> footnotes = new ArrayList<int[]>();
      List<TextParagraph> toTextMapping = new ArrayList<>();
      List<List<Integer>> toParaMapping = new ArrayList<>();
      List<Integer> sortedTextIds;
      clear();
      boolean withDeleted = document.getMultiDocumentsHandler().getConfiguration().includeTrackedChanges();
      for (int i = 0; i < NUMBER_CURSOR_TYPES; i++) {
        toParaMapping.add(new ArrayList<Integer>());
      }
      List<DocumentText> documentTexts = new ArrayList<>();
      for (int i = 0; i < NUMBER_CURSOR_TYPES; i++) {
        documentTexts.add(null);
      }
      WtDocumentCursorTools docCursor = document.getDocumentCursorTools();
      if (docCursor == null) {
        WtMessageHandler.printToLogFile("DocumentCache: refreshWriterCache: docCursor == null: return");
        return;
      }
      documentTexts.set(CURSOR_TYPE_TEXT, docCursor.getAllTextParagraphs(withDeleted));
      documentTexts.set(CURSOR_TYPE_TABLE, docCursor.getTextOfAllTables(withDeleted));
      documentTexts.set(CURSOR_TYPE_SHAPE, docCursor.getTextOfAllShapes(withDeleted));
      documentTexts.set(CURSOR_TYPE_FOOTNOTE, docCursor.getTextOfAllFootnotes(withDeleted));
      documentTexts.set(CURSOR_TYPE_ENDNOTE, docCursor.getTextOfAllEndnotes(withDeleted));
      documentTexts.set(CURSOR_TYPE_HEADER_FOOTER, docCursor.getTextOfAllHeadersAndFooters(withDeleted));
      for (int i = 0; i < NUMBER_CURSOR_TYPES; i++) {
        if(documentTexts.get(i) == null) {
          documentTexts.set(i, new DocumentText());
        }
      }
      nText = documentTexts.get(CURSOR_TYPE_TEXT).paragraphs.size();
      nTable = documentTexts.get(CURSOR_TYPE_TABLE).paragraphs.size();
      nShape = documentTexts.get(CURSOR_TYPE_SHAPE).paragraphs.size();
      nFootnote = documentTexts.get(CURSOR_TYPE_FOOTNOTE).paragraphs.size();
      nEndnote = documentTexts.get(CURSOR_TYPE_ENDNOTE).paragraphs.size();
      nHeaderFooter = documentTexts.get(CURSOR_TYPE_HEADER_FOOTER).paragraphs.size();
      if (debugMode) {
        for (int i = 0; i < NUMBER_CURSOR_TYPES; i++) {
          if (documentTexts.get(i) == null) {
            WtMessageHandler.printToLogFile("DocumentCache: refresh: CursorType: " + i + "; Document Text is Null!");
          } else {
            WtMessageHandler.printToLogFile("DocumentCache: refresh: CursorType: " + i + "; Number of paragraphs: "
                + documentTexts.get(i).paragraphs.size());
          }
        }
      }
      FlatParagraphContainer paragraphContainer = null;
      List<List<List<Integer>>> deletedChars = new ArrayList<>();
      for (DocumentText documentText : documentTexts) {
        List<Integer> hNumbers = new ArrayList<>();
        for (int n : documentText.headingNumbers.keySet()) {
          hNumbers.add(n);
        }
        hNumbers.sort(null);
        chapterBegins.add(hNumbers);
        deletedChars.add(documentText.deletedCharacters);
      }
      headingMap = documentTexts.get(CURSOR_TYPE_TEXT).headingNumbers;
//      MessageHandler.printToLogFile("DocumentCache: refresh: headingMap.size: " + (headingMap == null ? "null" : headingMap.size()));
      if (flatPara == null) {
        flatPara = document.getFlatParagraphTools();
      }
      if (flatPara == null) {
        WtMessageHandler.printToLogFile(
            "WARNING: DocumentCache: refresh: flatPara == null - ParagraphCache not initialised");
        return;
      }
      paragraphContainer = flatPara.getAllFlatParagraphs(fixedLocale);
      if (paragraphContainer == null) {
        WtMessageHandler.printToLogFile(
            "WARNING: DocumentCache: refresh: paragraphContainer == null - ParagraphCache not initialised");
        return;
      }
      if (paragraphContainer.paragraphs == null) {
        WtMessageHandler
            .printToLogFile("WARNING: DocumentCache: refresh: paragraphs in paragraphContainer == null - ParagraphCache not initialised");
        return;
      }
      paragraphs.addAll(paragraphContainer.paragraphs);
      for (Locale locale : paragraphContainer.locales) {
        locales.add(new SerialLocale(locale));
      }
      footnotes.addAll(paragraphContainer.footnotePositions);
      sortedTextIds = paragraphContainer.sortedTextIds;
      
      if (debugMode) {
        int unknown = 0;
        for (DocumentText documentText : documentTexts) {
          if (documentText != null && documentText.paragraphs != null) {
            unknown += documentText.paragraphs.size();
          }
        }
        unknown = paragraphs.size() - unknown;
        WtMessageHandler.printToLogFile("DocumentCache: refresh: unknown paragraphs: " + unknown);
        if (sortedTextIds == null) {
          WtMessageHandler.printToLogFile("DocumentCache: refresh: paragraphContainer.sortedTextIds == null");
        } else {
          for (DocumentText documentText : documentTexts) {
            if (documentText != null && documentText.sortedTextIds != null) {
              for (int n : documentText.sortedTextIds) {
                if (!sortedTextIds.contains(n)) {
                  WtMessageHandler.printToLogFile("DocumentCache: refresh: sortedTextId not in flatparagraph: " + n);
                }
              }
            }
          }
          for (int n : sortedTextIds) {
            boolean found = false;
            for (DocumentText documentText : documentTexts) {
              if (documentText != null && documentText.sortedTextIds != null && documentText.sortedTextIds.contains(n)) {
                found = true;
                break;
              }
            }
            if (!found) {
              WtMessageHandler.printToLogFile("DocumentCache: refresh: sortedTextId not in documentText: " + n);
            }
          }
        }
      }
      if (sortedTextIds == null) {
        List<List<String>> textParas = new ArrayList<>();
        for (DocumentText documentText : documentTexts) {
          textParas.add(documentText.paragraphs);
        }
        mapParagraphs(paragraphs, toTextMapping, toParaMapping, chapterBegins, locales, footnotes, textParas, deletedCharacters, deletedChars);
      } else {
        documentElementsCount = paragraphContainer.documentElementsCount;
        List<List<Integer>> textSortedTextIds = new ArrayList<>();
        for (DocumentText documentText : documentTexts) {
          textSortedTextIds.add(documentText.sortedTextIds);
        }
        mapParagraphsWNI(paragraphs, toTextMapping, toParaMapping, chapterBegins, locales, footnotes, textSortedTextIds, sortedTextIds, deletedCharacters, deletedChars);
      }
      addQuoteInfo(document, documentTexts.get(CURSOR_TYPE_TEXT).paragraphs);
      actualizeCache (paragraphs, chapterBegins, locales, footnotes, toTextMapping, toParaMapping, 
          deletedCharacters, documentTexts.get(CURSOR_TYPE_TEXT).automaticTextParagraphs, sortedTextIds);
//      for (Locale locale : getDifferentLocalesOftext(paragraphContainer.locales)) {
//        document.getMultiDocumentsHandler().handleLtDictionary(getDocAsString(), locale);
//      }
      document.getMultiDocumentsHandler().runShapeCheck(hasUnsupportedText(), fromWhere);
      if (debugModeTm) {
        long endTime = System.currentTimeMillis();
        WtMessageHandler.printToLogFile("Time to generate cache(" + fromWhere + "): " + (endTime - startTime));
      }
    } finally {
    }
  }
  
  /**
   * Actualize cache
   */
  private void actualizeCache (List<String> paragraphs, List<List<Integer>> chapterBegins, List<SerialLocale> locales, 
      List<int[]> footnotes, List<TextParagraph> toTextMapping, List<List<Integer>> toParaMapping, 
      List<List<Integer>> deletedCharacters, List<Integer> automaticParagraphs, List<Integer> sortedTextIds) {
    rwLock.writeLock().lock();
    try {
      clearAnalyzedParagraphs();
      this.paragraphs.clear();
      this.paragraphs.addAll(paragraphs);
      this.chapterBegins.clear();
      this.chapterBegins.addAll(chapterBegins);
      this.locales.clear();
      this.locales.addAll(locales);
      this.footnotes.clear();
      this.footnotes.addAll(footnotes);
      this.toTextMapping.clear();
      this.toTextMapping.addAll(toTextMapping);
      this.toParaMapping.addAll(toParaMapping);
      this.deletedCharacters.clear();
      this.deletedCharacters.addAll(deletedCharacters);
      this.automaticParagraphs.clear();
      this.automaticParagraphs.addAll(automaticParagraphs);
//      this.openingQuotes.addAll(openingQuotes);
//      this.closingQuotes.addAll(closingQuotes);
      if (sortedTextIds != null) {
        if (this.sortedTextIds == null) {
          this.sortedTextIds = new ArrayList<>();
        } else {
          this.sortedTextIds.clear();
        }
        this.sortedTextIds.addAll(sortedTextIds);
      }
      this.docLocale = getMostUsedLanguage(locales);
    } finally {
      rwLock.writeLock().unlock();
    }
  }
  
  private static boolean isEqualTextWithoutZeroSpace(String flatPara, String textPara) {
    if (flatPara.isEmpty() && textPara.isEmpty()) {
      return true;
    }
    flatPara = removeZeroWidthSpace(flatPara);
    textPara = removeZeroWidthSpace(textPara);
    if (flatPara.isEmpty() && textPara.isEmpty()) {
      return true;
    } else if (flatPara.length() != textPara.length()) {
      return false;
    }
    return flatPara.equals(textPara);
  }
  
  public boolean hasUnsupportedText() {
    if (toParaMapping.get(CURSOR_TYPE_SHAPE).size() > 0) {
      return true;
    }
    boolean hasUnsupported = false;
    for (TextParagraph tPara : toTextMapping) {
      if (tPara.type == CURSOR_TYPE_TEXT) {
        break;
      } else if (tPara.type == CURSOR_TYPE_TABLE) {
        hasUnsupported = true;
        break;
      }
    }
    return hasUnsupported;
  }
  
  public static boolean isEqualText(String flatPara, String textPara, int[] footnotes) {
    if (footnotes == null || footnotes.length == 0) {
      return isEqualTextWithoutZeroSpace(flatPara, textPara);
    }
    //  NOTE: flat paragraphs contain footnotes and endnotes as zero space characters
    //        text paragraphs contain footnotes and endnotes as digits or Roman characters
    try {
      if (footnotes[footnotes.length - 1] >= flatPara.length()) {
        WtMessageHandler.printToLogFile("DocumentCache: isEqualWithoutFootnotes: footnotes[footnotes.length - 1] >= flatPara.length()");
        return false;
      }
      int tParaBeg;
      String fPara;
      String tPara;
      textPara = removeZeroWidthSpace(textPara);
      if (footnotes[0] > 0) {
        fPara = flatPara.substring(0, footnotes[0]);
        fPara = removeZeroWidthSpace(fPara);
        if (!fPara.isEmpty()) {
          if (textPara.length() < fPara.length()) {
            return false;
          }
          tPara = textPara.substring(0, fPara.length());
          if (!tPara.equals(fPara)) {
            return false;
          }
        }
      }
      if (footnotes[footnotes.length - 1] < flatPara.length() - 1) {
        fPara = flatPara.substring(footnotes[footnotes.length - 1] + 1);
        fPara = removeZeroWidthSpace(fPara);
        if (!fPara.isEmpty()) {
          tParaBeg = textPara.length() - fPara.length();
          if (tParaBeg < 0) {
            return false;
          }
          tPara = textPara.substring(tParaBeg);
          if (!tPara.equals(fPara)) {
            return false;
          }
        } else {
          tParaBeg = textPara.length() - 1;
        }
      } else {
        tParaBeg = textPara.length() - 1;
      }
      for (int i = footnotes.length - 2; i >= 0; i--) {
        flatPara = flatPara.substring(0, footnotes[i + 1]);
        textPara = textPara.substring(0, tParaBeg);
        fPara = flatPara.substring(footnotes[i] + 1);
        fPara = removeZeroWidthSpace(fPara);
        tParaBeg = textPara.length() - fPara.length();
        boolean isEqual = false; 
        for (int j = 0; j <= MAX_NOTE_CHAR && !isEqual; j++) {
          if (tParaBeg - j < 0) {
            break;
          }
          tPara = textPara.substring(tParaBeg - j, textPara.length() - j);
          isEqual = tPara.equals(fPara);
          if (isEqual) {
            tParaBeg -= j;
          }
        }
        if (!isEqual) {
          return false;
        }
      }
      return true;
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      return false;
    }
  }
  
  /**
   * get the list of different locales out of the list of all locales
   */
  List<Locale> getDifferentLocalesOftext(List<Locale> locales) throws Throwable {
    List<Locale> differentLocales = new ArrayList<>();
    for (Locale locale : locales) {
      if (locale.Variant.startsWith(WtOfficeTools.MULTILINGUAL_LABEL)) {
        locale = new Locale(locale.Language, locale.Country, locale.Variant.substring(WtOfficeTools.MULTILINGUAL_LABEL.length()));
      }
      boolean isInList = false;
      for (Locale difLoc : differentLocales) {
        if (WtOfficeTools.isEqualLocale(locale, difLoc)) {
          isInList = true;
          break;
        }
      }
      if (!isInList) {
        differentLocales.add(locale); 
      }
    }
    return differentLocales;
  }
/*  
  private static boolean isEqualWithoutFootnotes(String flatPara, String textPara, int[] footnotes, int[] n, int level) {
    //  NOTE: flat paragraphs contain footnotes and endnotes as zero space characters
    //        text paragraphs contain footnotes and endnotes as digits or Roman characters
    for(n[level] = 0; n[level] <= MAX_NOTE_CHAR; n[level]++) {
      if (level == 0) {
        String textP = textPara;
        for (int i = footnotes.length - 1; i >= 0; i--) {
          int nDif = 0;
          for (int j = 0; j < i; j++) {
            nDif += n[j] - 1;
          }
          int k = footnotes[i] + nDif;
          if (k + n[i] > textPara.length()) {
            return false;
          }
          textP = textP.substring(0, k) + (k < textP.length() - n[i] ? textP.substring(k + n[i]) : "");
        }
        boolean isEqual = isEqualTextWithoutZeroSpace(flatPara, textP);
        if (isEqual) {
          return true;
        }
      } else {
        boolean isEqual = isEqualWithoutFootnotes(flatPara, textPara, footnotes, n, level - 1);
        if (isEqual) {
          return true;
        }
      }
    }
    return false;
  }

  public static boolean isEqualText(String flatPara, String textPara, int[] footnotes) {
    if (footnotes == null || footnotes.length == 0) {
      return isEqualTextWithoutZeroSpace(flatPara, textPara);
    }
    flatPara = SingleCheck.removeFootnotes(flatPara, footnotes, null);
    if (textPara.isEmpty() || flatPara.length() > textPara.length() || flatPara.length() < (textPara.length() - (footnotes.length * MAX_NOTE_CHAR))) {
      //  NOTE: size of footnote sign is assumed as <= MAX_NOTE_CHAR
      return false;
    }
    
    int[] n = new int[footnotes.length];
    for(int j = 0; j < n.length; j++) {
      n[j] = 0;
    }
    return isEqualWithoutFootnotes(flatPara, textPara, footnotes, n, n.length - 1);
  }
*/
  
  private static int[] getFootnotes(List<int[]> footnotes, int i) {
    return footnotes != null && i < footnotes.size() ? footnotes.get(i) : null;
  }
  
  /**
   * Correct header/footer or table 
   */
  
  private void correctNegativeNumberEntries(int type, List<List<String>> textParas, List<String> paragraphs, 
      List<List<Integer>> toParaMapping, List<TextParagraph> toTextMapping)  {
    boolean isRemoved = false;
    for (int j = textParas.get(type).size() - 1; j >= 0; j--) {
      if (toParaMapping.get(type).get(j) < 0) {
        boolean isMapped = false;
        for (int i = 0; i < toTextMapping.size() && !isMapped; i++) {
          if (toTextMapping.get(i).type == CURSOR_TYPE_UNKNOWN
              && isEqualText(paragraphs.get(i), textParas.get(type).get(j), getFootnotes(footnotes, i))) {
            toParaMapping.get(type).set(j, i);
            toTextMapping.set(i, new TextParagraph(type, j));
            isMapped = true;
          }
        }
        if (!isMapped) {
          toParaMapping.get(type).remove(j);
          isRemoved = true;
        }
      }
    }
    if (isRemoved) {
      for (int j = 0; j < toParaMapping.get(type).size(); j++) {
        int i = toParaMapping.get(type).get(j);
        if (toTextMapping.get(i).number != j) {
          toTextMapping.set(i, new TextParagraph(type, j));
        }
      }
    }
  }
  
  void correctParaMapping (int type, List<List<Integer>> toParaMapping) {
    if (type != CURSOR_TYPE_TEXT) {
      for (int i = toParaMapping.get(type).size() - 1; i >= 0; i--) {
        if (toParaMapping.get(type).get(i) < 0) {
          toParaMapping.get(type).remove(i);
        }
      }
    }
  }
  
  /**
   * Map Text inside a loop of all text paragraphs of a cursor type
   * NOTE: This is needed for all types of cursor other than text and frame because they can be inside a frame disturbing the usual order 
   */
  
  private boolean mapTextParagraphsPerLoop(int type, int nFlat, List<String> paragraphs, List<int[]> footnotes, 
      List<List<String>> textParas, List<TextParagraph> toTextMapping, List<List<Integer>> toParaMapping, 
      List<List<Integer>> nMapped, List<Integer> nNext) {
    if (nMapped.get(type).size() < textParas.get(type).size()) {
      for (int k = nNext.get(type); k < textParas.get(type).size(); k++) {
        if (!nMapped.get(type).contains(k) && isEqualText(paragraphs.get(nFlat), textParas.get(type).get(k), getFootnotes(footnotes, nFlat))) {
          toTextMapping.add(new TextParagraph(type, k));
          toParaMapping.get(type).set(k, nFlat);
          nMapped.get(type).add(k);
          nNext.set(type, k < textParas.get(type).size() - 1 ? k + 1 : 0);
          return true;
        }
      }
      for (int k = 0; k < nNext.get(type); k++) {
        if (!nMapped.get(type).contains(k) && isEqualText(paragraphs.get(nFlat), textParas.get(type).get(k),getFootnotes(footnotes, nFlat))) {
          toTextMapping.add(new TextParagraph(type, k));
          toParaMapping.get(type).set(k, nFlat);
          nMapped.get(type).add(k);
          nNext.set(type, k < textParas.get(type).size() - 1 ? k + 1 : 0);
          return true;
        }
      }
    }
    return false;
  }

  /**
   * Map text paragraphs to flat paragraphs is only used for writer documents
   * Use a heuristic procedure
   */
  private void mapParagraphs(List<String> paragraphs, List<TextParagraph> toTextMapping, List<List<Integer>> toParaMapping,
        List<List<Integer>> chapterBegins, List<SerialLocale> locales, List<int[]> footnotes, List<List<String>> textParas, 
        List<List<Integer>> deletedCharacters, List<List<List<Integer>>> deletedChars) {
    if (textParas != null && !textParas.isEmpty()) {
      List<List<Integer>> nMapped = new ArrayList<>();  // Mapped paragraphs per cursor type
      List<Integer> nNext = new ArrayList<>();          //  Next assumed paragraph number for cursor type
      for (int i = 0; i < textParas.size(); i++) {
        nMapped.add(new ArrayList<>());
        nNext.add(0);
      }
      int nUnknown = 0;
      for (int i = 0; i < NUMBER_CURSOR_TYPES; i++) {
        nUnknown += textParas.get(i).size();
      }
      nUnknown = paragraphs.size() - nUnknown;  // nUnknown: number of headings of graphic elements
      int printCount = 0;
      if (debugMode) {
        WtMessageHandler.printToLogFile("DocumentCache: mapParagraphs: Unknown paragraphs: " + nUnknown);
      }
      for (int j = 0; j < NUMBER_CURSOR_TYPES; j++) {
        if (j != CURSOR_TYPE_TEXT) {
          for (int i = 0; i < textParas.get(j).size(); i++) {
            toParaMapping.get(j).add(-1);
          }
        }
      }
      if (debugMode) {
        for (int i = 0; i < NUMBER_CURSOR_TYPES; i++) {
          for (int j = 0; j < textParas.get(i).size(); j++) {
            WtMessageHandler.printToLogFile("DocumentCache: mapParagraphs: type: " + i + "; text: " + textParas.get(i).get(j));
          }
        }
      }
      int numUnknown = 0;
      boolean thirdTextDone = false;
      for (int i = 0; i < paragraphs.size(); i++) {
        boolean isMapped = false;
        boolean firstTextDone = nMapped.get(CURSOR_TYPE_ENDNOTE).size() == textParas.get(CURSOR_TYPE_ENDNOTE).size()
            && nMapped.get(CURSOR_TYPE_FOOTNOTE).size() == textParas.get(CURSOR_TYPE_FOOTNOTE).size();
        boolean secondTextDone = firstTextDone 
            && nMapped.get(CURSOR_TYPE_SHAPE).size() == textParas.get(CURSOR_TYPE_SHAPE).size()
            && nMapped.get(CURSOR_TYPE_HEADER_FOOTER).size() == textParas.get(CURSOR_TYPE_HEADER_FOOTER).size();
        //  test for footnote, endnote, frame or header/footer
        //  listed before text or embedded tables
        if (!firstTextDone) {
          // test for footnote or endnote
          for (int n = 0; n < 2; n++) {
            if (mapTextParagraphsPerLoop(n, i, paragraphs, footnotes, textParas, toTextMapping, toParaMapping, nMapped, nNext)) {
              isMapped = true;
              break;
            }
          }
          if (isMapped) {
            continue;
          }
        }
        if (!secondTextDone && firstTextDone) {
          // test for header/footer, frame or shape
          for (int n = 2; n < NUMBER_CURSOR_TYPES - 2; n++) {
            if (mapTextParagraphsPerLoop(n, i, paragraphs, footnotes, textParas, toTextMapping, toParaMapping, nMapped, nNext)) {
              isMapped = true;
              break;
            }
          }
          if (isMapped) {
            continue;
          }
        }
        //  test for movable tables
        //  listed before text or embedded tables
        if ((!secondTextDone || !thirdTextDone) && firstTextDone && nMapped.get(CURSOR_TYPE_TABLE).size() < textParas.get(CURSOR_TYPE_TABLE).size()) {
          if (!secondTextDone || textParas.get(CURSOR_TYPE_TEXT).isEmpty() || 
              !isEqualText(paragraphs.get(i), textParas.get(CURSOR_TYPE_TEXT).get(0), getFootnotes(footnotes, i))) {
            if (mapTextParagraphsPerLoop(CURSOR_TYPE_TABLE, i, paragraphs, footnotes, textParas, toTextMapping, toParaMapping, nMapped, nNext)) {
              continue;
            } else if (secondTextDone) {
              nNext.set(CURSOR_TYPE_TABLE, 0);
              thirdTextDone = true;
            }
          }
        }
        //  there are unknown paragraphs before text paragraphs
        if (!secondTextDone) {
          toTextMapping.add(new TextParagraph(CURSOR_TYPE_UNKNOWN, -1));
          numUnknown++;
          if (debugMode) {
            WtMessageHandler.printToLogFile("WARNING: DocumentCache: Could not map Paragraph(" + i + "): '" + paragraphs.get(i) +
                "'; secondTextDone: " + secondTextDone);
            String msg = "DocumentCache: mapParagraphs:\n";
            for (int k = 0; k < NUMBER_CURSOR_TYPES; k++) {
              if (k == CURSOR_TYPE_TEXT) {
                msg += "Actual Cursor Paragraph (Type " + k + "): " + 
                        (nNext.get(k) < textParas.get(k).size() ? textParas.get(k).get(nNext.get(k)) : "no paragraph left") + "\n";
              } else {
                msg += "Actual Cursor Paragraph (Type " + k + "): " + 
                    (nMapped.get(k).size() < textParas.get(k).size() ? textParas.get(k).get(nNext.get(k)) : "no paragraph left") + "\n";
              }
            }
            WtMessageHandler.printToLogFile(msg);
            WtMessageHandler.printToLogFile("Unknown Paragraphs: " + (numUnknown - 1) + " from " + nUnknown);
          }
          continue;
        }
        int j = nNext.get(CURSOR_TYPE_TEXT);
        if (j == textParas.get(CURSOR_TYPE_TEXT).size() && nMapped.get(CURSOR_TYPE_TABLE).size() == textParas.get(CURSOR_TYPE_TABLE).size()) {
          //  no text and tables left ==> unknown paragraph
          toTextMapping.add(new TextParagraph(CURSOR_TYPE_UNKNOWN, -1));
          numUnknown++;
          if (debugMode && (!paragraphs.get(i).isEmpty() && printCount < MAX_PRINTED_PARAS)) {
            printCount++;
            WtMessageHandler.printToLogFile(
                "Warning: DocumentCache: Could not map Paragraph(" + i + "): '" + paragraphs.get(i) + "'; secondTextDone: " + secondTextDone);
          }
          continue;
        } else if (nMapped.get(CURSOR_TYPE_TABLE).size() == textParas.get(CURSOR_TYPE_TABLE).size() && numUnknown >= nUnknown) {
          //  no tables left / text assumed
          toTextMapping.add(new TextParagraph(CURSOR_TYPE_TEXT, j));
          toParaMapping.get(CURSOR_TYPE_TEXT).add(i);
          nNext.set(CURSOR_TYPE_TEXT, j + 1);
          continue;
        } else if (j == textParas.get(CURSOR_TYPE_TEXT).size() && numUnknown >= nUnknown) {
          //  no text left / table assumed
          int k = nNext.get(CURSOR_TYPE_TABLE);
          for (; k < textParas.get(CURSOR_TYPE_TABLE).size() && !nMapped.get(CURSOR_TYPE_TABLE).contains(k); k++); 
          if (k == textParas.get(CURSOR_TYPE_TABLE).size()) {
            for (k = 0; k < nNext.get(CURSOR_TYPE_TABLE) && !nMapped.get(CURSOR_TYPE_TABLE).contains(k); k++);
          }
          nMapped.get(CURSOR_TYPE_TABLE).add(k);
          nNext.set(CURSOR_TYPE_TABLE, k + 1);
          toTextMapping.add(new TextParagraph(CURSOR_TYPE_TABLE, k));
          toParaMapping.get(CURSOR_TYPE_TABLE).set(k, i);
          continue;
        } else {
          //  test for table
          int k = nNext.get(CURSOR_TYPE_TABLE);
          for (; k < textParas.get(CURSOR_TYPE_TABLE).size(); k++) {
            if (!nMapped.get(CURSOR_TYPE_TABLE).contains(k)) {
              break;
            }
          }
          if (k < textParas.get(CURSOR_TYPE_TABLE).size() && 
              isEqualText(paragraphs.get(i), textParas.get(CURSOR_TYPE_TABLE).get(k), getFootnotes(footnotes, i))) {
            boolean equalTable = true;
            boolean equalText = j < textParas.get(CURSOR_TYPE_TEXT).size() &&
                isEqualText(paragraphs.get(i), textParas.get(CURSOR_TYPE_TEXT).get(j), getFootnotes(footnotes, i));
            //  test if table and text are equal
            int mk = k;
            int mj = j;
            int mi = i;
            while (equalTable && equalText && mi < paragraphs.size() - 1) {
              mi++;
              mk++;
              if (mk < textParas.get(CURSOR_TYPE_TABLE).size() && !nMapped.get(CURSOR_TYPE_TABLE).contains(mk)) {
                equalTable = isEqualText(paragraphs.get(mi), textParas.get(CURSOR_TYPE_TABLE).get(mk), getFootnotes(footnotes, mi));
              } else {
                equalTable = false;
              }
              mj++;
              if (mj < textParas.get(CURSOR_TYPE_TEXT).size()) {
                equalText = isEqualText(paragraphs.get(mi), textParas.get(CURSOR_TYPE_TEXT).get(mj), getFootnotes(footnotes, mi));
              } else {
                equalText = false;
              }
            }
            if (!equalText) {
              nMapped.get(CURSOR_TYPE_TABLE).add(k);
              nNext.set(CURSOR_TYPE_TABLE, k + 1);
              toTextMapping.add(new TextParagraph(CURSOR_TYPE_TABLE, k));
              toParaMapping.get(CURSOR_TYPE_TABLE).set(k, i);
              continue;
            }
            isMapped = true;
          }
          //  test for text
          if (!isMapped) {
            if (j < textParas.get(CURSOR_TYPE_TEXT).size() && isEqualText(paragraphs.get(i), textParas.get(CURSOR_TYPE_TEXT).get(j), getFootnotes(footnotes, i))) {
              isMapped = true;
            }
          }
          if (isMapped) {
            toTextMapping.add(new TextParagraph(CURSOR_TYPE_TEXT, j));
            toParaMapping.get(CURSOR_TYPE_TEXT).add(i);
            nNext.set(CURSOR_TYPE_TEXT, j + 1);
            continue;
          }
        }
        //  unknown paragraph
        toTextMapping.add(new TextParagraph(CURSOR_TYPE_UNKNOWN, -1));
        numUnknown++;
        if (debugMode && (!paragraphs.get(i).isEmpty() && printCount < MAX_PRINTED_PARAS)) {
          printCount++;
          WtMessageHandler.printToLogFile(
              "Warning: DocumentCache: Could not map Paragraph(" + i + "): '" + paragraphs.get(i) + "'; secondTextDone: " + secondTextDone);
        }
        if (debugMode) {
          if (nNext.get(CURSOR_TYPE_TABLE) >= textParas.get(CURSOR_TYPE_TABLE).size()) {
            nNext.set(CURSOR_TYPE_TABLE, 0);
          }
          String msg = "DocumentCache: mapParagraphs:\n";
          for (int k = 0; k < NUMBER_CURSOR_TYPES; k++) {
            if (k == CURSOR_TYPE_TEXT) {
              msg += "Actual Cursor Paragraph (Type " + k + "): " + 
                      (nNext.get(k) < textParas.get(k).size() ? textParas.get(k).get(nNext.get(k)) : "no paragraph left") + "\n";
            } else {
              msg += "Actual Cursor Paragraph (Type " + k + "): " + 
                  (nMapped.get(k).size() < textParas.get(k).size() ? textParas.get(k).get(nNext.get(k)) : "no paragraph left") + "\n";
            }
          }
          WtMessageHandler.printToLogFile(msg);
          WtMessageHandler.printToLogFile("Unknown Paragraphs: " + (numUnknown - 1) + " from " + nUnknown);
        }
      }
      boolean isCorrectNonText = true;
      boolean isCorrectMapping = true;
      int nMap = 0;
      for (int i = 0; i < NUMBER_CURSOR_TYPES; i++) {
        if (i != CURSOR_TYPE_TEXT) {
          nMap += nMapped.get(i).size();
        }
      }
      nMap += toParaMapping.get(CURSOR_TYPE_TEXT).size();
      if (nMap == paragraphs.size()) {
        //  remove unmapped entries in to paraMapping
        for (int i = 0; i < NUMBER_CURSOR_TYPES; i++) {
          correctParaMapping(i, toParaMapping);
        }
      } else {
        isCorrectNonText = nMapped.get(CURSOR_TYPE_ENDNOTE).size() == textParas.get(CURSOR_TYPE_ENDNOTE).size()
            && nMapped.get(CURSOR_TYPE_FOOTNOTE).size() == textParas.get(CURSOR_TYPE_FOOTNOTE).size()
            && nMapped.get(CURSOR_TYPE_SHAPE).size() == textParas.get(CURSOR_TYPE_SHAPE).size()
            && nMapped.get(CURSOR_TYPE_HEADER_FOOTER).size() == textParas.get(CURSOR_TYPE_HEADER_FOOTER).size()
            && nMapped.get(CURSOR_TYPE_TABLE).size() == textParas.get(CURSOR_TYPE_TABLE).size();
        if (!isCorrectNonText) {
          //  Try to repair incorrect header/footer mapping
          for (int k = 0; k < NUMBER_CURSOR_TYPES; k++) {
            if (k != CURSOR_TYPE_TEXT) {
              if (nMapped.get(k).size() < textParas.get(k).size()) {
                WtMessageHandler.printToLogFile("Warning: document cache mapping failed: Try to repair mapping of paragraph type " + k);
                correctNegativeNumberEntries(k, textParas, paragraphs, toParaMapping, toTextMapping);
              }
            }
          }
          isCorrectNonText = toParaMapping.get(CURSOR_TYPE_ENDNOTE).size() == textParas.get(CURSOR_TYPE_ENDNOTE).size()
              && toParaMapping.get(CURSOR_TYPE_FOOTNOTE).size() == textParas.get(CURSOR_TYPE_FOOTNOTE).size()
              && toParaMapping.get(CURSOR_TYPE_SHAPE).size() == textParas.get(CURSOR_TYPE_SHAPE).size()
              && toParaMapping.get(CURSOR_TYPE_HEADER_FOOTER).size() == textParas.get(CURSOR_TYPE_HEADER_FOOTER).size()
              && toParaMapping.get(CURSOR_TYPE_TABLE).size() == textParas.get(CURSOR_TYPE_TABLE).size();
        }
        isCorrectMapping = isCorrectNonText && toParaMapping.get(CURSOR_TYPE_TEXT).size() == textParas.get(CURSOR_TYPE_TEXT).size();
        if (!isCorrectMapping && isCorrectNonText) {
          //  Try to repair incorrect text mapping
          WtMessageHandler.printToLogFile("\nWarning: document cache mapping failed: Try to repair mapping of paragraph type text");
          printCount = 0;
          toParaMapping.get(CURSOR_TYPE_TEXT).clear();
          for (int i = 0; i < paragraphs.size(); i++) {
            if (toTextMapping.get(i).type == CURSOR_TYPE_TEXT) {
              toTextMapping.set(i, new TextParagraph(CURSOR_TYPE_UNKNOWN, -1));
            }
          }
          boolean allmapped = true;
          for (int j = 0; j < textParas.get(CURSOR_TYPE_TEXT).size(); j++) {
            boolean ismapped = false;
            for (int i = 0; i < paragraphs.size() && !ismapped; i++) {
              if ((toTextMapping.get(i).type == CURSOR_TYPE_UNKNOWN) && 
                  isEqualText(paragraphs.get(i), textParas.get(CURSOR_TYPE_TEXT).get(j), getFootnotes(footnotes, i))) {
                toTextMapping.set(i, new TextParagraph(CURSOR_TYPE_TEXT, j));
                toParaMapping.get(CURSOR_TYPE_TEXT).add(i);
                ismapped = true;
              }
            }
            if (!ismapped) {
              allmapped = false;
              if (debugMode || printCount < MAX_PRINTED_PARAS) {
                printCount++;
                WtMessageHandler.printToLogFile("Warning: Could not map text paragraph: " + textParas.get(CURSOR_TYPE_TEXT).get(j));
              }
            }
          }
          if (!allmapped) {
            WtMessageHandler.printToLogFile("Warning: unknow non empty paragraphs (max. " + MAX_PRINTED_PARAS + " printed):");
            printCount = 0;
            for (int i = 0; i < paragraphs.size() && printCount < MAX_PRINTED_PARAS; i++) {
              if (toTextMapping.get(i).type == CURSOR_TYPE_UNKNOWN && !paragraphs.get(i).isEmpty()) {
                printCount++;
                WtMessageHandler.printToLogFile(i + ": " + paragraphs.get(i));
                for(int j = 0; j < paragraphs.get(i).length(); j++) {
                  if (!Character.isLetterOrDigit(paragraphs.get(i).codePointAt(j)) && paragraphs.get(i).charAt(j) != ' '  && paragraphs.get(i).charAt(j) != '\t') {
                    WtMessageHandler.printToLogFile("CharAt(" + j + "): " + paragraphs.get(i).codePointAt(j));
                  }
                }
              }
            }  
          }
        }
        isCorrectMapping = isCorrectNonText && toParaMapping.get(CURSOR_TYPE_TEXT).size() == textParas.get(CURSOR_TYPE_TEXT).size();
      }
      if (!isCorrectMapping) {
        numUnknown = 0;
        for (int i = 0; i < NUMBER_CURSOR_TYPES; i++) {
          numUnknown += toParaMapping.get(i).size();
        }
        numUnknown = paragraphs.size() - numUnknown;  // nUnknown: number of headings of graphic elements
        String msg = "An error has occurred in LanguageTool "
            + WtVersionInfo.ltVersion() + " (" + WtVersionInfo.ltBuildDate() + "):\nDocument cache mapping failed:\nParagraphs:\n"
            + "Endnotes: " + toParaMapping.get(CURSOR_TYPE_ENDNOTE).size() + " / " + textParas.get(CURSOR_TYPE_ENDNOTE).size() + "\n"
            + "Footnotes: " + toParaMapping.get(CURSOR_TYPE_FOOTNOTE).size() + " / " + textParas.get(CURSOR_TYPE_FOOTNOTE).size() + "\n"
            + "Headers/Footers: " + toParaMapping.get(CURSOR_TYPE_HEADER_FOOTER).size() + " / " + textParas.get(CURSOR_TYPE_HEADER_FOOTER).size() + "\n"
            + "Shapes: " + toParaMapping.get(CURSOR_TYPE_SHAPE).size() + " / " + textParas.get(CURSOR_TYPE_SHAPE).size() + "\n"
            + "Tables: " + toParaMapping.get(CURSOR_TYPE_TABLE).size() + " / " + textParas.get(CURSOR_TYPE_TABLE).size() + "\n"
            + "Text: " + toParaMapping.get(CURSOR_TYPE_TEXT).size() + " / " + textParas.get(CURSOR_TYPE_TEXT).size() + "\n"
            + "Unknown: " + numUnknown + " / " + nUnknown;
        WtMessageHandler.printToLogFile(msg);
      }
      nText = toParaMapping.get(CURSOR_TYPE_TEXT).size();
      nTable = toParaMapping.get(CURSOR_TYPE_TABLE).size();
      nShape = toParaMapping.get(CURSOR_TYPE_SHAPE).size();
      nFootnote = toParaMapping.get(CURSOR_TYPE_FOOTNOTE).size();
      nEndnote = toParaMapping.get(CURSOR_TYPE_ENDNOTE).size();
      nHeaderFooter = toParaMapping.get(CURSOR_TYPE_HEADER_FOOTER).size();
      mapDeletedCharacters(deletedCharacters, paragraphs, deletedChars, toTextMapping);
      prepareChapterBeginsForText(chapterBegins, toTextMapping, locales);
      if (debugMode) {
        WtMessageHandler.printToLogFile("\nDocumentCache: mapParagraphs: toParaMapping:");
        for (int n = 0; n < NUMBER_CURSOR_TYPES; n++) {
          WtMessageHandler.printToLogFile("Cursor Type: " + n);
          for (int i = 0; i < toParaMapping.get(n).size(); i++) {
            WtMessageHandler
                .printToLogFile("DocumentCache: mapParagraphs: Doc: " + i + " Flat: " + toParaMapping.get(n).get(i));
          }
        }
        WtMessageHandler.printToLogFile("\nDocumentCache: mapParagraphs: toTextMapping:");
        for (int i = 0; i < toTextMapping.size(); i++) {
          WtMessageHandler.printToLogFile("DocumentCache: mapParagraphs: Flat: " + i + " Doc: "
              + toTextMapping.get(i).number + " Type: " + toTextMapping.get(i).type + "; locale: "
              + locales.get(i).Language + "-" + locales.get(i).Country 
              + "; Deleted Chars size: " + (deletedCharacters.get(i) == null ? "null" : deletedCharacters.get(i).size())
              + "; '" + paragraphs.get(i) + "'");
        }
        WtMessageHandler.printToLogFile("\nDocumentCache: mapParagraphs: headings:");
        for (int n = 0; n < NUMBER_CURSOR_TYPES; n++) {
          WtMessageHandler.printToLogFile("\nDocumentCache: mapParagraphs: Cursor Type: " + n);
          for (int i = 0; i < chapterBegins.get(n).size(); i++) {
            WtMessageHandler
                .printToLogFile("DocumentCache: mapParagraphs: Num: " + i + " Heading: " + chapterBegins.get(n).get(i));
          }
        }
        WtMessageHandler.printToLogFile("DocumentCache: mapParagraphs: Number of Flat Paragraphs: " + paragraphs.size());
        WtMessageHandler.printToLogFile(
            "DocumentCache: mapParagraphs: Number of Text Paragraphs: " + toParaMapping.get(CURSOR_TYPE_TEXT).size());
        WtMessageHandler.printToLogFile("DocumentCache: mapParagraphs: Number of footnotes: " + footnotes.size());
        WtMessageHandler.printToLogFile("DocumentCache: mapParagraphs: Number of locales: " + locales.size());
        WtMessageHandler.printToLogFile("DocumentCache: mapParagraphs: Number of Deleted Chars: " + deletedCharacters.size());
      }
    }
  }
  /**
   * Map text paragraphs to flat paragraphs is only used for writer documents
   * Use node indexes of paragraphs (since LO 7.5)
   */
  private void mapParagraphsWNI(List<String> paragraphs, List<TextParagraph> toTextMapping, List<List<Integer>> toParaMapping,
        List<List<Integer>> chapterBegins, List<SerialLocale> locales, List<int[]> footnotes, List<List<Integer>> textSortedTextIds, 
        List<Integer> sortedTextIds, List<List<Integer>> deletedCharacters, List<List<List<Integer>>> deletedChars) {
    isDirty = false;
    List<Integer> nMapped = new ArrayList<>();  // Mapped paragraphs per cursor type
    int nUnknown = 0;
    for (int i = 0; i < NUMBER_CURSOR_TYPES; i++) {
      nMapped.add(0);
      if (textSortedTextIds.get(i) == null) {
        textSortedTextIds.set(i, new ArrayList<>());
        if (debugMode) {
          WtMessageHandler.printToLogFile("Document cache: mapParagraphsWNI: Empty textSortedTextIds for type: " + i);
        }
      }
      nUnknown += textSortedTextIds.get(i).size();
      for (int j = 0; j < textSortedTextIds.get(i).size(); j++) {
        toParaMapping.get(i).add(-1);
      }
    }
    nUnknown = sortedTextIds.size() - nUnknown;
    if (nUnknown < 0) {
      isDirty = true;
      if (debugMode) {
        WtMessageHandler.printToLogFile("WARNING cache is dirty (map on basis of node index); unknown paragraphs:" + nUnknown);
      }
    }
    if (debugMode) {
      for (int i = 0; i < NUMBER_CURSOR_TYPES; i++) {
        WtMessageHandler.printToLogFile("\nDocument cache: mapParagraphsWNI: node indexes for type: " + i);
        for (int j = 0; j < textSortedTextIds.get(i).size(); j++) {
          WtMessageHandler.printToLogFile("Document cache: mapParagraphsWNI: node index: " + textSortedTextIds.get(i).get(j));
        }
      }
      WtMessageHandler.printToLogFile("\nDocument cache: mapParagraphsWNI: unknown paragraphs: " + nUnknown);
      WtMessageHandler.printToLogFile("DocumentCache: mapParagraphs: Number of Flat Paragraphs: " + paragraphs.size());
      WtMessageHandler.printToLogFile(
          "DocumentCache: mapParagraphs: Number of Text Paragraphs: " + toParaMapping.get(CURSOR_TYPE_TEXT).size());
      WtMessageHandler.printToLogFile("DocumentCache: mapParagraphs: Number of footnotes: " + footnotes.size());
      WtMessageHandler.printToLogFile("DocumentCache: mapParagraphs: Number of locales: " + locales.size());
      WtMessageHandler.printToLogFile("DocumentCache: mapParagraphs: Number of Deleted Chars: " + deletedCharacters.size());
    }
    int notMapped = 0;
    for (int n = 0; n < sortedTextIds.size(); n++) {
      boolean found = false;
      for (int i = 0; !found && i < NUMBER_CURSOR_TYPES; i++) {
        List<Integer> txtNdIndexes = textSortedTextIds.get(i);
        for (int j = 0; !found && j < txtNdIndexes.size(); j++) {
          if (((int)txtNdIndexes.get(j)) == ((int)sortedTextIds.get(n))) {
            found = true;
            toParaMapping.get(i).set(j, n);
            toTextMapping.add(new TextParagraph(i, j));
          }
        }
      }
      if (!found) {
        toTextMapping.add(new TextParagraph(CURSOR_TYPE_UNKNOWN, -1));
        notMapped++;
        if (debugMode) {
          WtMessageHandler.printToLogFile("Document cache: mapParagraphsWNI: Not found node: " + sortedTextIds.get(n));
        }
      }
    }
    if (notMapped > 0 && notMapped != nUnknown) {
      if (debugMode) {
        WtMessageHandler.printToLogFile("WARNING not mapped paragraphs (map on basis of node index); notmapped known paragraphs: " + (notMapped - nUnknown)
          + "; unknown paragraphs: " + nUnknown);
      }
    }
    if (isDirty) {
      for (int i = 0; i < NUMBER_CURSOR_TYPES; i++) {
        for (int j = toParaMapping.get(i).size() - 1; j >= 0; j--) {
          if (toParaMapping.get(i).get(j) < 0) {
            toParaMapping.get(i).remove(j);
          }
        }
      }
      nText = toParaMapping.get(CURSOR_TYPE_TEXT).size();
      nTable = toParaMapping.get(CURSOR_TYPE_TABLE).size();
      nShape = toParaMapping.get(CURSOR_TYPE_SHAPE).size();
      nFootnote = toParaMapping.get(CURSOR_TYPE_FOOTNOTE).size();
      nEndnote = toParaMapping.get(CURSOR_TYPE_ENDNOTE).size();
      nHeaderFooter = toParaMapping.get(CURSOR_TYPE_HEADER_FOOTER).size();
    }
    
    mapDeletedCharacters(deletedCharacters, paragraphs, deletedChars, toTextMapping);
    prepareChapterBeginsForText(chapterBegins, toTextMapping, locales);
  }

  /**
   * reset the document cache for impress documents
   */
  private void refreshImpressCalcCache(XComponent xComponent) {
    rwLock.writeLock().lock();
    try {
      isDirty = false;
      ParagraphContainer container;
      if (docType == DocumentType.IMPRESS) {
        container = WtOfficeDrawTools.getAllParagraphs(xComponent);
      } else if (docType == DocumentType.CALC) {
        container = WtOfficeSpreadsheetTools.getAllParagraphs(xComponent);
      } else {
        return;
      }
      clear();
      paragraphs.addAll(container.paragraphs);
      for (int i = 0; i < NUMBER_CURSOR_TYPES - 1; i++) {
        chapterBegins.add(new ArrayList<Integer>());
      }
      chapterBegins.get(CURSOR_TYPE_TEXT).addAll(container.pageBegins);
      for (Locale locale : container.locales) {
        locales.add(new SerialLocale(locale));
      }
      for (int i = 0; i < NUMBER_CURSOR_TYPES; i++) {
        toParaMapping.add(new ArrayList<Integer>());
      }
      for (int i = 0; i < paragraphs.size(); i++) {
        toTextMapping.add(new TextParagraph(CURSOR_TYPE_TEXT, i));
        toParaMapping.get(CURSOR_TYPE_TEXT).add(i);
        footnotes.add(new int[0]);
        deletedCharacters.add(null);
      }
      nText = toParaMapping.get(CURSOR_TYPE_TEXT).size();
      nTable = toParaMapping.get(CURSOR_TYPE_TABLE).size();
      nShape = toParaMapping.get(CURSOR_TYPE_SHAPE).size();
      nFootnote = toParaMapping.get(CURSOR_TYPE_FOOTNOTE).size();
      nEndnote = toParaMapping.get(CURSOR_TYPE_ENDNOTE).size();
      nHeaderFooter = toParaMapping.get(CURSOR_TYPE_HEADER_FOOTER).size();
      docLocale = getMostUsedLanguage(locales);
      if (debugMode) {
        WtMessageHandler.printToLogFile("DocumentCache: reset: isImpress: Number of paragraphse: " + paragraphs.size());
        for (int i = 0; i < NUMBER_CURSOR_TYPES; i++) {
          WtMessageHandler.printToLogFile("DocumentCache: reset: CursorType: " + i + "; Number of paragraphs: " + toParaMapping.get(i).size());
        }
      }
    } catch (Throwable t) {
      isDirty = true;
      WtMessageHandler.showError(t);
    } finally {
      rwLock.writeLock().unlock();
    }
  }
  
  /**
   * Set text level cache to no errors for single paragraph text
   */
  private void setSingleParagraphsCacheToNull(List<WtResultCache> paragraphsCache) {
    for (int i = 0; i < paragraphs.size(); i++) {
      if (isSingleParagraph_intern(i)) {
        for (int n = 1; n < paragraphsCache.size(); n++) {
          paragraphsCache.get(n).put(i, new WtProofreadingError[0]);
        }
      }
    }
  }

  /**
   * Set text level cache for one paragraph to no errors for single paragraph text
   */
  public boolean setSingleParagraphsCacheToNull(int numberFlatParagraph, List<WtResultCache> paragraphsCache) {
    rwLock.readLock().lock();
    try {
      if (isSingleParagraph_intern(numberFlatParagraph)) {
        for (int n = 1; n < paragraphsCache.size(); n++) {
          paragraphsCache.get(n).put(numberFlatParagraph, new WtProofreadingError[0]);
        }
        return true;
      } else {
        return false;
      }
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * wait till reset is finished
   */
  public boolean isResetRunning() {
    return isReset;
  }

  /**
   * wait till reset is finished
   */
  public boolean isFinished() {
    while (isReset) {
      try {
        Thread.sleep(5);
      } catch (InterruptedException e) {
        WtMessageHandler.showError(e);
        return false;
      }
    }
    return paragraphs == null ? false : true;
  }

  /**
   * get Flat Paragraph by Index
   */
  public String getFlatParagraph(int n) {
    rwLock.readLock().lock();
    try {
      String para = n < 0 || n >= paragraphs.size() ? null : paragraphs.get(n);
      return para;
    } finally {
      rwLock.readLock().unlock();
    }
}

  /**
   * set Flat Paragraph at Index
   */
  public void setFlatParagraph(int n, String sPara) {
    rwLock.writeLock().lock();
    try {
      removeAnalyzedParagraph(n);
      if (n >= 0 && n < paragraphs.size()) {
        paragraphs.set(n, sPara);
      }
    } finally {
      rwLock.writeLock().unlock();
    }
  }

  /**
   * set Flat Paragraph and Locale at Index
   */
  public void setFlatParagraph(int n, String sPara, Locale locale) {
    rwLock.writeLock().lock();
    try {
      removeAnalyzedParagraph(n);
      if (n >= 0 && n < locales.size()) {
        locales.set(n, new SerialLocale(locale));
      }
      if (n >= 0 && n < paragraphs.size()) {
        paragraphs.set(n, sPara);
      }
    } finally {
      rwLock.writeLock().unlock();
    }
  }

  /**
   * is multilingual Flat Paragraph
   */
  public boolean isMultilingualFlatParagraph(int n) {
    rwLock.readLock().lock();
    try {
      return isMultilingualFlatParagraphIntern(n);
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * is multilingual Flat Paragraph (only intern)
   */
  private boolean isMultilingualFlatParagraphIntern(int n) {
    return n < 0 || n >= locales.size() ? false : locales.get(n).Variant.startsWith(WtOfficeTools.MULTILINGUAL_LABEL);
  }

  /**
   * set multilingual flag to Flat Paragraph
   */
  public void setMultilingualFlatParagraph(int n) {
    rwLock.writeLock().lock();
    try {
      removeAnalyzedParagraph(n);
      if (n >= 0 && n < locales.size()) {
        SerialLocale locale = locales.get(n);
        if (!locale.Variant.startsWith(WtOfficeTools.MULTILINGUAL_LABEL)) {
          locale.Variant = WtOfficeTools.MULTILINGUAL_LABEL + locale.Variant;
          locales.set(n, locale);
        }
      }
    } finally {
      rwLock.writeLock().unlock();
    }
  }

  /**
   * get Locale of Flat Paragraph by Index
   */
  public Locale getFlatParagraphLocale(int n) {
    rwLock.readLock().lock();
    try {
      return locales.get(n).toLocaleWithoutLabel();
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * set Locale of Flat Paragraph by Index
   */
  public void setFlatParagraphLocale(int n, Locale locale) {
    rwLock.writeLock().lock();
    try {
      removeAnalyzedParagraph(n);
      if (n >= 0 && n < locales.size()) {
        locales.set(n, new SerialLocale(locale));
      }
    } finally {
      rwLock.writeLock().unlock();
    }
  }

  /**
   * get footnotes of Flat Paragraph by Index
   */
  public int[] getFlatParagraphFootnotes(int n) {
    rwLock.readLock().lock();
    try {
      if (n >= 0 && n < footnotes.size()) {
        return footnotes.get(n);
      } else {
        return new int[0];
      }
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * set footnotes of Flat Paragraph by Index
   */
  public void setFlatParagraphFootnotes(int n, int[] footnotePos) {
    rwLock.writeLock().lock();
    try {
      removeAnalyzedParagraph(n);
      if (n >= 0 && n < footnotes.size()) {
        footnotes.set(n, footnotePos);
      }
    } finally {
      rwLock.writeLock().unlock();
    }
  }
  
  /**
   * get deleted characters (report changes) of Flat Paragraph by Index
   */
  public List<Integer> getFlatParagraphDeletedCharacters(int n) {
    rwLock.readLock().lock();
    try {
      if (n >= 0 && n < deletedCharacters.size()) {
        return deletedCharacters.get(n);
      } else {
        return new ArrayList<>();
      }
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get deleted characters (report changes) of Flat Paragraph by Index
   */
  public boolean isAutomaticGenerated(int n, boolean alsoIgnore) {
    rwLock.readLock().lock();
    try {
      if (n >= 0 && n < toTextMapping.size()) {
        if (alsoIgnore && locales.get(n).Language.equals(WtOfficeTools.IGNORE_LANGUAGE)) {
          return true;
        }
        TextParagraph tPara = toTextMapping.get(n);
        if (tPara.type == CURSOR_TYPE_TEXT && automaticParagraphs.contains(tPara.number)) {
          return true;
        }
      }
      return false;
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * set deleted characters (report changes) of Flat Paragraph by Index
   */
  public void setFlatParagraphDeletedCharacters(int n, List<Integer> deletedChars) {
    rwLock.writeLock().lock();
    try {
      removeAnalyzedParagraph(n);
      if (n >= 0 && n < deletedCharacters.size()) {
        deletedCharacters.set(n, deletedChars);
      }
    } finally {
      rwLock.writeLock().unlock();
    }
  }

  /**
   * correct a start point to change of flat paragraph by zero space characters
   */
  public int correctStartPoint(int nStart, int nFPara) {
    rwLock.readLock().lock();
    try {
      int cor = 0;
      for (int i = 0; i < nStart; i++) {
        if (paragraphs.get(nFPara).charAt(i) == WtOfficeTools.ZERO_WIDTH_SPACE_CHAR) {
          boolean isFootnote = false;
          for (int n : footnotes.get(nFPara)) {
            if (n == i) {
              isFootnote = true;
              break;
            }
          }
          if (!isFootnote) {
            cor++;
          }
        }
      }
      return nStart - cor;
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * clear document cache
   */
  private void clear() {
    paragraphs.clear();
    chapterBegins.clear();
    locales.clear();
    footnotes.clear();
    toTextMapping.clear();
    toParaMapping.clear();
    deletedCharacters.clear();
    automaticParagraphs.clear();
    openingQuotes.clear();
    closingQuotes.clear();
    if (sortedTextIds != null) {
      sortedTextIds.clear();
    }
  }
  
  /**
   * Add a document Cache
   */
  private void add(WtDocumentCache in) {
    paragraphs.addAll(in.paragraphs);
    chapterBegins.addAll(in.chapterBegins);
    locales.addAll(in.locales);
    footnotes.addAll(in.footnotes);
    toTextMapping.addAll(in.toTextMapping);
    for (int i = 0; i < NUMBER_CURSOR_TYPES; i++) {
      toParaMapping.add(new ArrayList<Integer>(in.toParaMapping.get(i)));
    }
    deletedCharacters.addAll(in.deletedCharacters);
    if (in.sortedTextIds != null) {
      sortedTextIds = new ArrayList<>(in.sortedTextIds);
    }
    if (in.headingMap != null) {
      headingMap = new HashMap<>(in.headingMap);
    }
    documentElementsCount = in.documentElementsCount;
    nText = in.nText;
    nTable = in.nTable;
    nShape = in.nShape;
    nFootnote = in.nFootnote;
    nEndnote = in.nEndnote;
    nHeaderFooter = in.nHeaderFooter;
  }
  
  /**
   * Replace a document Cache
   */
  public void put(WtDocumentCache in) {
    rwLock.writeLock().lock();
    try {
      clear();
      add(in);
    } finally {
      rwLock.writeLock().unlock();
    }
  }

  /**
   * is DocumentCache empty
   */
  public boolean isEmpty() {
    rwLock.readLock().lock();
    try {
      return paragraphs == null || paragraphs.isEmpty();
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * has no content
   */
  public boolean hasNoContent(boolean lock) {
    if (lock) {
      rwLock.readLock().lock();
    }
    try {
      return paragraphs == null || paragraphs.isEmpty() || (paragraphs.size() == 1 && paragraphs.get(0).isEmpty());
    } finally {
      if (lock) {
        rwLock.readLock().unlock();
      }
    }
  }

  /**
   * size of document cache (number of all flat paragraphs)
   */
  public int size() {
    rwLock.readLock().lock();
    try {
      return paragraphs == null ? 0 : paragraphs.size();
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get the Content of a Text Paragraph
   */
  public String getTextParagraph(TextParagraph textParagraph) {
    rwLock.readLock().lock();
    try {
      if (textParagraph.type == CURSOR_TYPE_UNKNOWN || textParagraph.number < 0) {
        return new String("");
      }
      if (textParagraph.type >= toParaMapping.size() || textParagraph.number >= toParaMapping.get(textParagraph.type).size()) {
        return new String("");
      }
      int nFPara = toParaMapping.get(textParagraph.type).get(textParagraph.number);
      return nFPara < 0 ? new String("") : paragraphs.get(nFPara);
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get Number of Flat Paragraph from Number of Text Paragraph
   */
  public int getFlatParagraphNumber(TextParagraph textParagraph) {
    rwLock.readLock().lock();
    try {
      if (textParagraph.type == CURSOR_TYPE_UNKNOWN || textParagraph.number < 0
          || toParaMapping.size() < NUMBER_CURSOR_TYPES
          || toParaMapping.get(textParagraph.type).size() <= textParagraph.number) {
        return -1;
      }
      return toParaMapping.get(textParagraph.type).get(textParagraph.number);
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get Locale of Text Paragraph by Index
   */
  public Locale getTextParagraphLocale(TextParagraph textParagraph) {
    rwLock.readLock().lock();
    try {
      if (textParagraph.type == CURSOR_TYPE_UNKNOWN || textParagraph.number < 0 ||
          textParagraph.type >= toParaMapping.size() || textParagraph.number >= toParaMapping.get(textParagraph.type).size()) {
        return null;
      }
      int nFPara = toParaMapping.get(textParagraph.type).get(textParagraph.number);
      return nFPara < 0 ? null : locales.get(nFPara).toLocaleWithoutLabel();
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get deleted Characters of Text Paragraph
   */
  public List<Integer> getTextParagraphDeletedCharacters(TextParagraph textParagraph) {
    rwLock.readLock().lock();
    try {
      if (textParagraph.type == CURSOR_TYPE_UNKNOWN || textParagraph.number < 0) {
        return new ArrayList<Integer>();
      }
      int nFPara = toParaMapping.get(textParagraph.type).get(textParagraph.number);
      return nFPara < 0 ? new ArrayList<Integer>() : deletedCharacters.get(nFPara);
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get footnotes of Text Paragraph by Index
   */
  public int[] getTextParagraphFootnotes(TextParagraph textParagraph) {
    rwLock.readLock().lock();
    try {
      if (textParagraph.type == CURSOR_TYPE_UNKNOWN || textParagraph.number < 0) {
        return new int[0];
      }
      int nFPara = toParaMapping.get(textParagraph.type).get(textParagraph.number);
      return nFPara < 0 ? new int[0] : footnotes.get(nFPara);
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * set footnotes of Text Paragraph
   */
  public void setTextParagraphFootnotes(TextParagraph textParagraph, int[] footnotePos) {
    rwLock.writeLock().lock();
    try {
      removeAnalyzedTextParagraph(textParagraph);
      if (textParagraph.type != CURSOR_TYPE_UNKNOWN && textParagraph.number >= 0) {
        footnotes.set(toParaMapping.get(textParagraph.type).get(textParagraph.number), footnotePos);
      }
    } finally {
      rwLock.writeLock().unlock();
    }
  }

  /**
   * get Number of Text Paragraph from Number of Flat Paragraph
   */
  public TextParagraph getNumberOfTextParagraph(int numberOfFlatParagraph) {
    rwLock.readLock().lock();
    try {
      return getNumberOfTextParagraph_(numberOfFlatParagraph);
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * get Number of Text Paragraph from Number of Flat Paragraph
   */
  private TextParagraph getNumberOfTextParagraph_(int numberOfFlatParagraph) {
    if (numberOfFlatParagraph < 0 || numberOfFlatParagraph >= toTextMapping.size()) {
      return new TextParagraph(CURSOR_TYPE_UNKNOWN, -1);
    }
    return toTextMapping.get(numberOfFlatParagraph);
  }

  /**
   * get Type of Paragraph from flat paragraph number
   */
  public int getParagraphType(int numberOfFlatParagraph) {
    rwLock.readLock().lock();
    try {
      return toTextMapping.get(numberOfFlatParagraph).type;
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * size of text cache for a type of text
   */
  public int textSize(int type) {
    rwLock.readLock().lock();
    try {
      return (type < 0 || type >= toParaMapping.size()) ? 0 : toParaMapping.get(type).size();
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * size of text cache for a type of text
   */
  public int textSize(TextParagraph textParagraph) {
    rwLock.readLock().lock();
    try {
      return (textParagraph.type == CURSOR_TYPE_UNKNOWN || textParagraph.type >= toParaMapping.size()) ?
        0 : toParaMapping.get(textParagraph.type).size();
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * Text and local are equal to cache
   */
  public boolean isEqual(int n, String text, Locale locale) {
    rwLock.readLock().lock();
    try {
      return ((n < 0 || n >= locales.size() || locales.get(n) == null) ? false
        : ((isMultilingualFlatParagraphIntern(n) || locales.get(n).equalsLocale(locale)) && text.equals(paragraphs.get(n))));
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * Text, deleted chars and local are equal to cache
   */
  public boolean isEqual(int n, String text, Locale locale, List<Integer> delChars) {
    rwLock.readLock().lock();
    try {
      if (n < 0 || n >= locales.size() || locales.get(n) == null) {
        return false;
      }
      if (!isMultilingualFlatParagraphIntern(n) && !locales.get(n).equalsLocale(locale)) {
        return false;
      }
      if ((delChars != null && deletedCharacters.get(n) == null) || (delChars == null && deletedCharacters.get(n) != null) 
         || (delChars != null && deletedCharacters.get(n) != null && delChars.size() != deletedCharacters.get(n).size())) {
        return false;
      }
      return text.equals(paragraphs.get(n));
    } finally {
      rwLock.readLock().unlock();
    }
  }
  
  /**
   * size of cache has changed?
   */
  public boolean isEqualCacheSize(WtDocumentCursorTools docCursor) {
    rwLock.readLock().lock();
    try {
      if (nText != docCursor.getNumberOfAllTextParagraphs()) {
        return false;
      }
      if (nTable != docCursor.getNumberOfAllTables()) {
        return false;
      }
      if (nShape != docCursor.getNumberOfAllShapes()) {
        return false;
      }
      if (nFootnote != docCursor.getNumberOfAllFootnotes()) {
        return false;
      }
      if (nEndnote != docCursor.getNumberOfAllEndnotes()) {
        return false;
      }
      if (nHeaderFooter != docCursor.getNumberOfAllHeadersAndFooters()) {
        return false;
      }
      return true;
    } finally {
      rwLock.readLock().unlock();
    }
  }
  
  /**
   * get all paragraphs within unsupported text types (shapes, tables in frames)
   * change the cache to new value
   * return all changed paragraphs
   */
  public List<Integer> getChangedUnsupportedParagraphs(WtDocumentCursorTools docCursor, WtResultCache firstResultCache) {
    List<Integer> nChanged = new ArrayList<>();
    rwLock.writeLock().lock();
    try {
      if (docCursor == null) {
        return nChanged;
      }
      if (toParaMapping.get(CURSOR_TYPE_SHAPE).isEmpty() && toParaMapping.get(CURSOR_TYPE_TABLE).isEmpty()) {
        return null;
      }
      List<Integer> nTableParas = new ArrayList<>();
      List<Integer> nShapeParas = new ArrayList<>();
      for (int i = 0; i < paragraphs.size() && toTextMapping.get(i).type != CURSOR_TYPE_TEXT; i++) {
        if (toTextMapping.get(i).type == CURSOR_TYPE_SHAPE) {
          nShapeParas.add(toTextMapping.get(i).number);
        } else if (toTextMapping.get(i).type == CURSOR_TYPE_TABLE) {
          nTableParas.add(toTextMapping.get(i).number);
        }
      }
      if (!nShapeParas.isEmpty()) {
        List<String> fParas = docCursor.getTextOfShapes(nShapeParas);
        if (fParas != null) {
          for (int i = 0; i < fParas.size(); i++) {
            int nFPara = toParaMapping.get(CURSOR_TYPE_SHAPE).get(nShapeParas.get(i));
            if (firstResultCache.getCacheEntry(nFPara) == null || !paragraphs.get(nFPara).equals(fParas.get(i))) {
              removeAnalyzedParagraph(nFPara);
              paragraphs.set(nFPara, fParas.get(i));
              nChanged.add(nFPara);
            }
          }
        }
      }
      if (!nTableParas.isEmpty()) {
        List<String> fParas = docCursor.getTextOfTables(nTableParas);
        if (fParas != null) {
          for (int i = 0; i < fParas.size(); i++) {
            int nFPara = toParaMapping.get(CURSOR_TYPE_TABLE).get(nTableParas.get(i));
            if (firstResultCache.getCacheEntry(nFPara) == null || !paragraphs.get(nFPara).equals(fParas.get(i))) {
              removeAnalyzedParagraph(nFPara);
              paragraphs.set(nFPara, fParas.get(i));
              nChanged.add(nFPara);
            }
          }
        }
      }
    } finally {
      rwLock.writeLock().unlock();
    }
    return nChanged;
  }
/*  
  public List<Integer> getChangedUnsupportedParagraphs(FlatParagraphTools flatPara, ResultCache firstResultCache) {
    rwLock.writeLock().lock();
    List<Integer> nChanged = new ArrayList<>();
    try {
      if (flatPara == null) {
        return nChanged;
      }
      if (toParaMapping.get(CURSOR_TYPE_SHAPE).isEmpty() && toParaMapping.get(CURSOR_TYPE_TABLE).isEmpty()) {
        return null;
      }
      List<Integer> nParas = new ArrayList<>();
      for (int i = 0; i < paragraphs.size() && toTextMapping.get(i).type != CURSOR_TYPE_TEXT; i++) {
        if (toTextMapping.get(i).type == CURSOR_TYPE_SHAPE || toTextMapping.get(i).type == CURSOR_TYPE_TABLE) {
          nParas.add(i);
        }
      }
      List<String> fParas = flatPara.getFlatParagraphs(nParas);
      if (fParas == null) {
        return nChanged;
      }
      for (int i = 0; i < fParas.size(); i++) {
        int nFPara = nParas.get(i);
        if (firstResultCache.getCacheEntry(nFPara) == null || !paragraphs.get(nFPara).equals(fParas.get(i))) {
          paragraphs.set(nFPara, fParas.get(i));
          nChanged.add(nFPara);
        }
      }
    } finally {
      rwLock.writeLock().unlock();
    }
    return nChanged;
  }
*/  
  /**
   * is flat paragraph a single paragraph
   */
  public boolean isSingleParagraph(int numberOfFlatParagraph) {
    rwLock.readLock().lock();
    try {
      return isSingleParagraph_intern(numberOfFlatParagraph);
    } finally {
      rwLock.readLock().unlock();
    }
  }
 
  /**
   * is flat paragraph a single paragraph (intern, not secure)
   */
  private boolean isSingleParagraph_intern(int numberOfFlatParagraph) {
    if (numberOfFlatParagraph < 0 || numberOfFlatParagraph >= toTextMapping.size()) {
      return true;
    }
    TextParagraph textParagraph = toTextMapping.get(numberOfFlatParagraph);
    if (textParagraph.type == CURSOR_TYPE_UNKNOWN) {
      return true;
    }
    for (int n = 0; n < chapterBegins.get(textParagraph.type).size(); n++) {
      if (textParagraph.number == 0 ||
          textParagraph.number == chapterBegins.get(textParagraph.type).get(n)) {
        if (n == chapterBegins.get(textParagraph.type).size() - 1 || 
            chapterBegins.get(textParagraph.type).get(n + 1) == textParagraph.number + 1) {
          return true;
        }
        break;
      } else if (textParagraph.number < chapterBegins.get(textParagraph.type).get(n)) {
        break;
      }
    }
    return false;
  }
 
  /**
   * Gives back the start paragraph for text level check
   */
  public int getStartOfParaCheck(TextParagraph textParagraph, int parasToCheck, boolean checkOnlyParagraph,
      boolean useQueue, boolean addParas) {
    rwLock.readLock().lock();
    try {
      if (textParagraph.number < 0 || toParaMapping.get(textParagraph.type).size() <= textParagraph.number) {
        return -1;
      }
      if (parasToCheck < -1) { // check all flat paragraphs
        return 0;
      }
      if (parasToCheck == 0) {
        return textParagraph.number;
      }
      int headingBefore = 0;
      for (int heading : chapterBegins.get(textParagraph.type)) {
        if (heading > textParagraph.number) {
          break;
        }
        headingBefore = heading;
      }
      if (headingBefore == textParagraph.number || parasToCheck < 0 || (useQueue && !checkOnlyParagraph)) {
        return headingBefore;
      }
      int startPos = textParagraph.number - parasToCheck;
      if (addParas) {
        startPos -= parasToCheck;
      }
      if (startPos < headingBefore) {
        startPos = headingBefore;
      }
      return startPos;
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * Gives back the end paragraph for text level check
   */
  public int getEndOfParaCheck(TextParagraph textParagraph, int parasToCheck, boolean checkOnlyParagraph,
      boolean useQueue, boolean addParas) {
    rwLock.readLock().lock();
    try {
      if (textParagraph.number < 0 || toParaMapping.get(textParagraph.type).size() <= textParagraph.number) {
        return -1;
      }
      if (parasToCheck < -1) { // check all flat paragraphs
        return size();
      }
      int headingAfter = -1;
      if (parasToCheck == 0) {
        return textParagraph.number + 1;
      }
      for (int heading : chapterBegins.get(textParagraph.type)) {
        headingAfter = heading;
        if (heading > textParagraph.number) {
          break;
        }
      }
      if (headingAfter <= textParagraph.number || headingAfter > toParaMapping.get(textParagraph.type).size()) {
        headingAfter = toParaMapping.get(textParagraph.type).size();
      }
      if (parasToCheck < 0 || (useQueue && !checkOnlyParagraph)) {
        return headingAfter;
      }
      int endPos = textParagraph.number + 1 + parasToCheck;
      if (!checkOnlyParagraph) {
        endPos += parasToCheck * WtOfficeTools.CHECK_MULTIPLIKATOR;
      }
      if (addParas) {
        endPos += parasToCheck;
      }
      if (endPos > headingAfter) {
        endPos = headingAfter;
      }
      return endPos;
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * Gives Back the full Text as String sorted by cursor types
   */
  public String getDocAsString(TextParagraph textParagraph, int parasToCheck, boolean checkOnlyParagraph,
      boolean useQueue, boolean hasFootnotes) {
    rwLock.readLock().lock();
    try {
      int startPos = getStartOfParaCheck(textParagraph, parasToCheck, checkOnlyParagraph, useQueue, true);
      int endPos = getEndOfParaCheck(textParagraph, parasToCheck, checkOnlyParagraph, useQueue, true);
      StringBuilder docText;
      if (parasToCheck < -1) { // check all flat paragraphs
        if (startPos < 0 || endPos < 0
            || (hasFootnotes && getFlatParagraph(startPos).isEmpty() && getFlatParagraphFootnotes(startPos).length > 0)) {
          return "";
        }
        docText = new StringBuilder(fixLinebreak(WtSingleCheck.removeFootnotes(getFlatParagraph(startPos),
            (hasFootnotes ? getFlatParagraphFootnotes(startPos) : null), getFlatParagraphDeletedCharacters(startPos))));
        for (int i = startPos + 1; i < endPos; i++) {
          docText.append(WtOfficeTools.END_OF_PARAGRAPH).append(fixLinebreak(
              WtSingleCheck.removeFootnotes(getFlatParagraph(i), (hasFootnotes ? getFlatParagraphFootnotes(i) : null), getFlatParagraphDeletedCharacters(i))));
        }
      } else {
        TextParagraph startPara = new TextParagraph(textParagraph.type, startPos);
        if (startPos < 0 || endPos < 0 || (hasFootnotes && getTextParagraph(startPara).isEmpty()
            && getTextParagraphFootnotes(startPara).length > 0)) {
          return "";
        }
        docText = new StringBuilder(fixLinebreak(WtSingleCheck.removeFootnotes(getTextParagraph(startPara),
            (hasFootnotes ? getTextParagraphFootnotes(startPara) : null), getTextParagraphDeletedCharacters(startPara))));
        for (int i = startPos + 1; i < endPos; i++) {
          TextParagraph tPara = new TextParagraph(textParagraph.type, i);
          docText.append(WtOfficeTools.END_OF_PARAGRAPH).append(fixLinebreak(WtSingleCheck
              .removeFootnotes(getTextParagraph(tPara), (hasFootnotes ? getTextParagraphFootnotes(tPara) : null), getTextParagraphDeletedCharacters(tPara))));
        }
      }
      return docText.toString();
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      return "";
   } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * Gives Back the full Text as String
   */
  public String getDocAsString() {
    rwLock.readLock().lock();
    try {
      StringBuilder docText = new StringBuilder(paragraphs.get(0));
      for (int i = 1; i < paragraphs.size(); i++) {
        docText.append(WtOfficeTools.END_OF_PARAGRAPH).append(paragraphs.get(i));
      }
      return docText.toString();
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * Change manual linebreak to distinguish from end of paragraph
   */
  public static String fixLinebreak(String text) {
    return text.replaceAll(WtOfficeTools.SINGLE_END_OF_PARAGRAPH, WtOfficeTools.MANUAL_LINEBREAK);
  }

  /**
   * remove zero width space
   */
  public static String removeZeroWidthSpace(String text) {
    return text.replaceAll(WtOfficeTools.ZERO_WIDTH_SPACE, "");
  }

  /**
   * Gives Back the StartPosition of Paragraph
   */
  public int getStartOfParagraph(int nPara, TextParagraph textParagraph, int parasToCheck, boolean checkOnlyParagraph,
      boolean useQueue, boolean hasFootnotes) {
    rwLock.readLock().lock();
    try {
      if (nPara < 0 || (parasToCheck > -2 && toParaMapping.get(textParagraph.type).size() <= nPara)
          || (parasToCheck < -1 && paragraphs.size() <= nPara)) {
        return -1;
      }
      int startPos = getStartOfParaCheck(textParagraph, parasToCheck, checkOnlyParagraph, useQueue, true);
      if (startPos < 0) {
        return -1;
      }
      int pos = 0;
      if (parasToCheck < -1) {
        for (int i = startPos; i < nPara; i++) {
          pos += WtSingleCheck.removeFootnotes(getFlatParagraph(i), 
                  (hasFootnotes ? getFlatParagraphFootnotes(i) : null), getFlatParagraphDeletedCharacters(i)).length() + WtOfficeTools.NUMBER_PARAGRAPH_CHARS;
        }
      } else {
        for (int i = startPos; i < nPara; i++) {
          TextParagraph tPara = new TextParagraph(textParagraph.type, i);
          pos += WtSingleCheck.removeFootnotes(getTextParagraph(tPara), 
                  (hasFootnotes ? getTextParagraphFootnotes(tPara) : null), getTextParagraphDeletedCharacters(tPara)).length() + WtOfficeTools.NUMBER_PARAGRAPH_CHARS;
        }
      }
      return pos;
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      return -1;
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * Map deleted characters to flat paragraphs
   */
  private void mapDeletedCharacters(List<List<Integer>> deletedCharacters, List<String> paragraphs,
      List<List<List<Integer>>> deletedChars, List<TextParagraph> toTextMapping) {
    if (deletedChars == null) {
      for (int i = 0; i < toTextMapping.size(); i++) {
        deletedCharacters.add(null);
      }
    } else {
      for (int i = 0; i < toTextMapping.size(); i++) {
        if (toTextMapping.get(i).type == CURSOR_TYPE_UNKNOWN) {
          deletedCharacters.add(null);
          if (debugMode || !paragraphs.get(i).isEmpty()) {
            WtMessageHandler.printToLogFile("Warning: CURSOR_TYPE_UNKNOWN at Paragraph " + i + ": deleted Characters set to null");
            WtMessageHandler.printToLogFile("         Paragraph : '" + paragraphs.get(i) + "'");
          }
        } else {
          deletedCharacters.add(deletedChars.get(toTextMapping.get(i).type).get(toTextMapping.get(i).number));
        }
      }
    }
  }
  
  /**
   * For cursor type text: Add the next chapter begin after Heading and changes of
   * language to the chapter begins
   */
  private void prepareChapterBeginsForText(List<List<Integer>> chapterBegins, List<TextParagraph> toTextMapping, List<SerialLocale> locales) {
    List<Integer> prepChBegins = new ArrayList<Integer>(chapterBegins.get(CURSOR_TYPE_TEXT));
    for (int begin : chapterBegins.get(CURSOR_TYPE_TEXT)) {
      if (!prepChBegins.contains(begin + 1)) {
        prepChBegins.add(begin + 1);
      }
    }
    if (locales.size() > 0) {
      SerialLocale lastLocale = locales.get(0);
      for (int i = 1; i < locales.size(); i++) {
        if (locales != null && !locales.get(i).equalWithoutLabel(lastLocale)) {
          TextParagraph nText = toTextMapping.get(i);
          if (nText.type == CURSOR_TYPE_TEXT && nText.number >= 0) {
            if (!prepChBegins.contains(nText.number)) {
              prepChBegins.add(nText.number);
            }
            lastLocale = locales.get(i);
            if (debugMode) {
              WtMessageHandler.printToLogFile(
                  "DocumentCache: prepareChapterBeginsForText: Paragraph(" + i + "): Locale changed to: "
                      + lastLocale.Language + (lastLocale.Country == null ? "" : ("-" + lastLocale.Country)));
            }
          }
        }
      }
    }
    prepChBegins.sort(null);
    chapterBegins.set(CURSOR_TYPE_TEXT, prepChBegins);
  }

  public TextParagraph createTextParagraph(int type, int paragraph) {
    return new TextParagraph(type, paragraph);
  }
  
  /**
   * Refresh the cache and compare with the old
   * Give back the range of difference and size
   */
  public ChangedRange refreshAndCompare(WtSingleDocument document, Locale fixedLocale, Locale docLocale, XComponent xComponent, int fromWhere) {
    WtDocumentCache oldCache = new WtDocumentCache(this);
    this.refresh(document, fixedLocale, docLocale, xComponent, true, fromWhere);
    rwLock.readLock().lock();
    try {
      if (paragraphs == null || paragraphs.isEmpty() || oldCache.paragraphs == null || oldCache.paragraphs.isEmpty()) {
        return null;
      }
      int from = 0;
      int to = 1;
      // to prevent spontaneous recheck of nearly the whole text
      // the change of text contents has to be checked, if the size of CURSOR_TYPE_HEADER_FOOTER has not changed
      // ignore headers and footers and the change of function inside of them
      if (oldCache.toParaMapping.get(CURSOR_TYPE_HEADER_FOOTER).size() != toParaMapping.get(CURSOR_TYPE_HEADER_FOOTER).size()) {
        while (from < paragraphs.size() && from < oldCache.paragraphs.size()
            && (toTextMapping.get(from).type != WtDocumentCache.CURSOR_TYPE_HEADER_FOOTER
            || paragraphs.get(from).equals(oldCache.paragraphs.get(from)))) {
          from++;
        }
        while (to <= paragraphs.size() && to <= oldCache.paragraphs.size()
            && (toTextMapping.get(paragraphs.size() - to).type != WtDocumentCache.CURSOR_TYPE_HEADER_FOOTER
            || paragraphs.get(paragraphs.size() - to).equals(
                oldCache.paragraphs.get(oldCache.paragraphs.size() - to)))) {
          to++;
        }
      } else {
        while (from < paragraphs.size() && from < oldCache.paragraphs.size()
            && (toTextMapping.get(from).type == CURSOR_TYPE_HEADER_FOOTER
            || paragraphs.get(from).equals(oldCache.paragraphs.get(from)))) {
          from++;
        }
        while (to <= paragraphs.size() && to <= oldCache.paragraphs.size()
            && (toTextMapping.get(paragraphs.size() - to).type == WtDocumentCache.CURSOR_TYPE_HEADER_FOOTER
            || paragraphs.get(paragraphs.size() - to).equals(
                    oldCache.paragraphs.get(oldCache.paragraphs.size() - to)))) {
          to++;
        }
      }
      to = paragraphs.size() - to + 1;
      if (to < from) {
        to = from;
      }
      this.removeAndShiftAnalyzedParagraph(from, to, oldCache.paragraphs.size(), paragraphs.size());
      return new ChangedRange(from, to, oldCache.paragraphs.size(), paragraphs.size());
    } finally {
      rwLock.readLock().unlock();
    }
  }
  
  /**
   * has nearest paragraph changed
   *//*
  public boolean nearestParagraphHasChanged(int numberOfFlatParagraph, FlatParagraphTools flatPara) {
    if (flatPara == null || numberOfFlatParagraph < 0 || numberOfFlatParagraph > paragraphs.size() - 1) {
      return true;
    }
    if (paragraphs.size() == 1) {
      return false;
    }
    int pNum = numberOfFlatParagraph == paragraphs.size() - 1 ? numberOfFlatParagraph - 1 : numberOfFlatParagraph + 1;
    if (!flatPara.getFlatParagraphAt(pNum).getText().equals(paragraphs.get(pNum))) {
      return true;
    }
    return false;
  }
*/
  /**
   * Get Map of Headings (only cursor type text)
   */
  public Map<Integer, Integer> getHeadingMap() {
    return headingMap;
  }
  
  /**
   * Return nearest sorted text Id
   */
  public int getNearestSortedTextId(int sortedTextId) {
    rwLock.readLock().lock();
    try {
      if (sortedTextIds == null) {
        return -1;
      }
      for (int i = 0; i < sortedTextIds.size(); i++) {
        if (sortedTextIds.get(i) == sortedTextId) {
          if (i == sortedTextIds.size() - 1) {
            return sortedTextIds.get(i - 1);
          } else {
            return sortedTextIds.get(i + 1);
          }
        }
      }
      return -1;
    } finally {
      rwLock.readLock().unlock();
    }
  }
  
  /**
   * Return Number of flat Paragraph from node index
   */
  public int getFlatparagraphFromSortedTextId(int sortedTextId) {
    rwLock.readLock().lock();
    try {
      if (sortedTextIds == null) {
        return -1;
      }
      for (int i = 0; i < sortedTextIds.size(); i++) {
        if (sortedTextIds.get(i) == sortedTextId) {
          return i;
        }
      }
      return -1;
    } finally {
      rwLock.readLock().unlock();
    }
  }

  /**
   * Return false if cache has to be actualized
   */
  public boolean isActual(int documentElementsCount) {
    rwLock.readLock().lock();
    try {
      if (isDirty || sortedTextIds == null || documentElementsCount == -1 || this.documentElementsCount != documentElementsCount) {
        return false;
      }
      return true;
    } finally {
      rwLock.readLock().unlock();
    }
  }
  
  private SerialLocale getMostUsedLanguage(List<SerialLocale> locales) {
    Map<SerialLocale, Integer> localesMap = new HashMap<>();
    for (SerialLocale locale : locales) {
      boolean localeExists = false;
      for (SerialLocale loc : localesMap.keySet()) {
        if (loc.equalsLocale(locale)) {
          localesMap.put(loc, localesMap.get(loc) + 1);
          localeExists = true;
          break;
        }
      }
      if (!localeExists) {
        localesMap.put(locale, 1);
      }
    }
    int max = 0;
    SerialLocale maxLocale = null;
    for (SerialLocale loc : localesMap.keySet()) {
      if (localesMap.get(loc) > max && WtDocumentsHandler.hasLocale(loc.toLocale())) {
        max = localesMap.get(loc);
        maxLocale = loc;
      }
    }
    if (maxLocale != null) {
      return maxLocale;
    }
    return null;
  }
  
  public Locale getDocumentLocale() {
    rwLock.readLock().lock();
    try {
      if (docLocale == null) {
        return null;
      }
      return docLocale.toLocaleWithoutLabel();
    } finally {
      rwLock.readLock().unlock();
    }
  }
  
  /**
   * Remove all analyzed paragraphs
   */
  public void clearAnalyzedParagraphs() {
    analyzedParagraphs.clear();
  }
  
  /**
   * Get all analyzed paragraphs
   */
  public Map<Integer, List<AnalyzedSentence>> getAllAnalyzedParagraphs() {
    return analyzedParagraphs;
  }
  
  /**
   * Get an analyzed paragraphs
   */
  public List<AnalyzedSentence> getAnalyzedParagraph(int nFPara) {
    rwLock.readLock().lock();
    try {
      if (nFPara >= 0 && nFPara < analyzedParagraphs.size()) {
        return analyzedParagraphs.get(nFPara);
      } else {
        return null;
      }
    } finally {
      rwLock.readLock().unlock();
    }
  }
  
  /**
   * Has analyzed paragraph same length as text
   */
  public boolean isCorrectAnalyzedParagraphLength(int nFPara, String text) throws Throwable { 
    List<AnalyzedSentence> analyzedSentences = getAnalyzedParagraph(nFPara);
    if (analyzedSentences == null) {
      return false;
    }
    text = fixLinebreak(WtSingleCheck.removeFootnotes(text, 
        getFlatParagraphFootnotes(nFPara), getFlatParagraphDeletedCharacters(nFPara)));
    int len = 0;
    for (AnalyzedSentence analyzedSentence : analyzedSentences) {
      String sentence = analyzedSentence.getText();
      len += sentence.length();
    }
    return len == text.length();
  }

  /**
   * Remove an analyzed paragraph
   */
  private void removeAnalyzedParagraph(int nFPara) {
    analyzedParagraphs.remove(nFPara);
  }
  
  /**
   * Remove an analyzed text paragraph
   */
  private void removeAnalyzedTextParagraph(TextParagraph textParagraph) {
    if (textParagraph.type != CURSOR_TYPE_UNKNOWN && textParagraph.number >= 0 
        && textParagraph.number < toParaMapping.get(textParagraph.type).size()) {
      analyzedParagraphs.remove(toParaMapping.get(textParagraph.type).get(textParagraph.number));
    }
  }
  
  /**
   * Put an analyzed paragraphs
   */
  public void putAnalyzedParagraph(int nFPara, List<AnalyzedSentence> analyzedParagraph) {
    rwLock.writeLock().lock();
    try {
      analyzedParagraphs.put(nFPara, analyzedParagraph);
    } finally {
      rwLock.writeLock().unlock();
    }
  }
  
  /**
   * Remove and shift analyzed paragraphs by a range
   */
  private void removeAndShiftAnalyzedParagraph(int fromParagraph, int toParagraph, int oldSize, int newSize) {
    if (analyzedParagraphs == null || analyzedParagraphs.isEmpty()) {
      return;
    }
    int shift = newSize - oldSize;
    if (fromParagraph < 0 && toParagraph >= newSize) {
      return;
    }
    Map<Integer, List<AnalyzedSentence>> tmpParagraphs = new HashMap<>(analyzedParagraphs);
    analyzedParagraphs.clear();
    if (shift < 0) {   // new size < old size
      for (int i : tmpParagraphs.keySet()) {
        if (i < fromParagraph) {
          analyzedParagraphs.put(i, tmpParagraphs.get(i));
        } else if (i >= toParagraph - shift) {
          analyzedParagraphs.put(i + shift, tmpParagraphs.get(i));
        }
      }
    } else {
      for (int i : tmpParagraphs.keySet()) {
        if (i < fromParagraph - 1) {  // remove also the last paragraph before (it was changed in many cases)
          analyzedParagraphs.put(i, tmpParagraphs.get(i));
        } else if (i >= toParagraph + 1) {  // remove also the last paragraph after (it was changed in many cases)
          analyzedParagraphs.put(i + shift, tmpParagraphs.get(i));
        }
      }
    }
  }

  /**
   * create an analyzed paragraph and store it in analyzed Cache
   */
  public List<AnalyzedSentence> createAnalyzedParagraph(int nFPara, WtLanguageTool lt) {
    try {
      String paraText = getFlatParagraph(nFPara);
      if (paraText == null) {
        return null;
      }
      paraText = WtSingleCheck.removeFootnotes(paraText, 
          getFlatParagraphFootnotes(nFPara), getFlatParagraphDeletedCharacters(nFPara));
      return createAnalyzedParagraph(nFPara, paraText, lt);
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      return null;
    }
  }

  private List<AnalyzedSentence> createAnalyzedParagraph(int nFPara, String paraText, WtLanguageTool lt) throws IOException {
    List<AnalyzedSentence> analyzedParagraph = lt.analyzeText(paraText.replace("\u00AD", ""));
    putAnalyzedParagraph(nFPara, analyzedParagraph);
    return analyzedParagraph;
  }

  /**
   * Get an analyzed paragraph from analyzed Cache
   * if the requested paragraph doesn't exist create it
   */
  public AnalysedText getOrCreateAnalyzedParagraph(int nFPara, WtLanguageTool lt) {
    try {
      String paraText = getFlatParagraph(nFPara);
      if (paraText == null) {
        return null;
      }
      paraText = fixLinebreak(WtSingleCheck.removeFootnotes(paraText, 
          getFlatParagraphFootnotes(nFPara), getFlatParagraphDeletedCharacters(nFPara)));
      paraText = paraText.replace("\u00AD", "");
      List<AnalyzedSentence> analyzedSentences = getAnalyzedParagraph(nFPara);
      List<String> sentences = new ArrayList<>();
      if (analyzedSentences == null) {
        analyzedSentences = createAnalyzedParagraph(nFPara, paraText, lt);
      }
      int len = 0;
      for (AnalyzedSentence analyzedSentence : analyzedSentences) {
        String sentence = analyzedSentence.getText();
        len += sentence.length();
        sentences.add(sentence);
      }
      if (len != paraText.length()) {
  /*
        String anSenText = "";
        for (String txt : sentences) {
          anSenText += txt;
        }
        MessageHandler.printToLogFile("DocumentCache: getOrCreateAnalyzedParagraph: Different length: (" + len + "/" + paraText.length()
            + "):\nparaText: '" + paraText + "'\nAnalysed: '" + anSenText + "'");
  */
        analyzedSentences = createAnalyzedParagraph(nFPara, paraText, lt);
        sentences.clear();
        for (AnalyzedSentence analyzedSentence : analyzedSentences) {
          sentences.add(analyzedSentence.getText());
        }
      }
      return new AnalysedText(analyzedSentences, sentences, paraText);
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      return null;
    }
  }

  /**
   * Get a range of analyzed paragraphs from analyzed Cache
   * if the requested paragraphs don't exist create it
   */
  public AnalysedText getAnalyzedParagraphs(TextParagraph from, TextParagraph to, WtLanguageTool lt) throws IOException {
    List<AnalyzedSentence> analyzedParagraphs = new ArrayList<>();
    List<String> sentences = new ArrayList<>();
    StringBuilder docText = new StringBuilder();
    for (int i = from.number; i < to.number; i++) {
      int n = getFlatParagraphNumber(new TextParagraph(from.type, i));
      AnalysedText analyzedParagraph = getOrCreateAnalyzedParagraph(n, lt);
      if (analyzedParagraph == null) {
        return null;
      }
      if (analyzedParagraph.analyzedSentences.size() == 0) {
        if (analyzedParagraphs.size() > 0) {
          int last = analyzedParagraphs.size() - 1;
          AnalyzedSentence sentence = analyzedParagraphs.get(last);
          AnalyzedTokenReadings[] tokens = sentence.getTokens();
          int len = tokens.length;
          AnalyzedTokenReadings[] newTokens = new AnalyzedTokenReadings[len + 2];
          for (int k = 0; k < len; k++) {
            newTokens[k] = tokens[k];
          }
          int startPos = tokens[len - 1].getEndPos();
          newTokens[len] = new AnalyzedTokenReadings(new AnalyzedToken("\n", null, null), startPos);
          newTokens[len + 1] = new AnalyzedTokenReadings(new AnalyzedToken("\n", null, null), startPos + 1);
          sentence = new AnalyzedSentence(newTokens);
          analyzedParagraphs.set(last, sentence);
          sentences.set(last, sentence.getText());
        } else {
          AnalyzedTokenReadings[] newTokens = new AnalyzedTokenReadings[3];
          newTokens[0] = new AnalyzedTokenReadings(new AnalyzedToken("", JLanguageTool.SENTENCE_START_TAGNAME, null), 0);
          newTokens[1] = new AnalyzedTokenReadings(new AnalyzedToken("\n", null, null), 0);
          newTokens[2] = new AnalyzedTokenReadings(new AnalyzedToken("\n", JLanguageTool.SENTENCE_END_TAGNAME, null), 1);
          newTokens[2].setParagraphEnd();
          AnalyzedSentence sentence = new AnalyzedSentence(newTokens);
          analyzedParagraphs.add(sentence);
          sentences.add(sentence.getText());
        }
      } else {
        for (int j = 0; j < analyzedParagraph.analyzedSentences.size(); j++) {
          AnalyzedSentence sentence = analyzedParagraph.analyzedSentences.get(j);
          if (j == analyzedParagraph.analyzedSentences.size() - 1 && i < to.number - 1) {
            AnalyzedTokenReadings[] tokens = sentence.getTokens();
            int len = tokens.length;
            AnalyzedTokenReadings[] newTokens = new AnalyzedTokenReadings[len + 2];
            for (int k = 0; k < len; k++) {
              newTokens[k] = tokens[k];
            }
            int startPos = tokens[len - 1].getEndPos();
            newTokens[len] = new AnalyzedTokenReadings(new AnalyzedToken("\n", null, null), startPos);
            newTokens[len + 1] = new AnalyzedTokenReadings(new AnalyzedToken("\n", null, null), startPos + 1);
            sentence = new AnalyzedSentence(newTokens);
            sentences.add(sentence.getText());
          } else {
            sentences.add(analyzedParagraph.sentences.get(j));
          }
          analyzedParagraphs.add(sentence);
        }
      }
      docText.append(analyzedParagraph.text);
      if (i < to.number - 1) {
        docText.append(WtOfficeTools.END_OF_PARAGRAPH);
      }
    }
    return new AnalysedText(analyzedParagraphs, sentences, docText.toString());
  }
  
  public static class TextParagraph implements Serializable {
    private static final long serialVersionUID = 1L;
    public int type;
    public int number;

    public TextParagraph(int type, int number) {
      this.type = type;
      this.number = number;
    }
  }
/*
  public static void printTokenizedSentences(List<AnalyzedSentence> sentences) {
    for (AnalyzedSentence sentence : sentences) {
      String str = "";
      for (AnalyzedTokenReadings token : sentence.getTokens()) {
        str += "'" + token.getToken(); 
        if (token.isSentenceStart()) {
          str += "{sent start}";
        }
        if (token.isSentenceEnd()) {
          str += "{sent end}";
        }
        if (token.isParagraphEnd()) {
          str += "{para end}";
        }
        str += "' ";
      }
      MessageHandler.printToLogFile("Sentence: " + str);
    }
  }
*/
  /**
   * add information about opening and closing quotes
   */
  private void addQuoteInfo(WtSingleDocument document, List<String> textParagraphs) {
    boolean needQuoteInfo = document.getMultiDocumentsHandler().getConfiguration().getCheckDirectSpeech() 
        != WtConfiguration.CHECK_DIRECT_SPEECH_YES;
    if (needQuoteInfo) {
      WtQuotesDetection quotesDetector = new WtQuotesDetection();
      quotesDetector.analyzeTextParagraphs(textParagraphs, openingQuotes, closingQuotes);
    }
  }

  /**
   * update information about opening and closing quotes
   */
  public void updateQuoteInfo(WtSingleDocument document, int nPara) {
    try {
      boolean needQuoteInfo = document.getMultiDocumentsHandler().getConfiguration().getCheckDirectSpeech() 
          != WtConfiguration.CHECK_DIRECT_SPEECH_YES;
      if (needQuoteInfo) {
        TextParagraph tPara = getNumberOfTextParagraph(nPara);
        if (tPara.type == CURSOR_TYPE_TEXT) {
          WtQuotesDetection quotesDetector = new WtQuotesDetection();
          quotesDetector.updateTextParagraph(getFlatParagraph(nPara), nPara, openingQuotes, closingQuotes);
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }

  /**
   * is opening quote before nStart
   * nTPara numbers the text paragraphs
   */
  private boolean isOpenQuote(int nTPara, int nStart) {
    if (openingQuotes.size() <= nTPara || openingQuotes.get(nTPara) == null || openingQuotes.get(nTPara).size() == 0) {
      return false;
    }
    for (int i = openingQuotes.get(nTPara).size() - 1; i >= 0; i--) {
      if (openingQuotes.get(nTPara).get(i) <= nStart) {
        if (closingQuotes.size() <= nTPara || closingQuotes.get(nTPara) == null || closingQuotes.get(nTPara).size() == 0) {
          return true;
        }
        for (int j = closingQuotes.get(nTPara).size() - 1; j >= 0; j--) {
          if (closingQuotes.get(nTPara).get(j) <= nStart) {
            if (closingQuotes.get(nTPara).get(j) > openingQuotes.get(nTPara).get(i)) {
              return false;
            } else {
              return true;
            }
          }
        }
        return true;
      }
    }
    return false;
  }
  
  /**
   * filter sentence in quotes our of error array
   */
  public WtProofreadingError[] filterDirectSpeech (WtProofreadingError[] errorArray, int nPara, WtConfiguration config) {
    if (config.getCheckDirectSpeech() == WtConfiguration.CHECK_DIRECT_SPEECH_YES 
        || errorArray == null || errorArray.length == 0) {
      return errorArray;
    }
    TextParagraph tPara = getNumberOfTextParagraph(nPara);
    if (tPara.type != CURSOR_TYPE_TEXT) {
      return errorArray;
    }
    List<WtProofreadingError> errors = new ArrayList<>();
    boolean isNoStyleDirectSpeech = config.getCheckDirectSpeech() == WtConfiguration.CHECK_DIRECT_SPEECH_NO_STYLE;
    for (WtProofreadingError error : errorArray) {
      if (error.bPunctuationRule || (!error.bStyleRule && isNoStyleDirectSpeech) || !isOpenQuote(tPara.number, error.nErrorStart)) {
        errors.add(error);
      }
    }
    return errors.toArray(new WtProofreadingError[0]);
  }

  /**
   * Class of serializable locale needed to save cache
   */
  public static class SerialLocale implements Serializable {

    private static final long serialVersionUID = 1L;
    String Country;
    String Language;
    String Variant;

    SerialLocale(Locale locale) {
      this.Country = locale.Country;
      this.Language = locale.Language;
      this.Variant = locale.Variant;
    }

    SerialLocale(String country, String language, String variant) {
      this.Country = country;
      this.Language = language;
      this.Variant = variant;
    }

    /**
     *  Get a String from SerialLocale
     */
    public String toString() {
      return Language + (Country.isEmpty() ? "" : "-" + Country) + (Variant.isEmpty() ? "" : "-" + Variant);
    }
    
    /**
     *  Gives SerialLocale without multilingual label
     */
    public SerialLocale withoutLabel() {
      if (Variant.startsWith(WtOfficeTools.MULTILINGUAL_LABEL)) {
        return new SerialLocale(Language, Country, Variant.substring(WtOfficeTools.MULTILINGUAL_LABEL.length()));
      }
      return new SerialLocale(Language, Country, Variant);
    }

    /**
     *  Gives SerialLocale without multilingual label
     */
    public boolean equalWithoutLabel(SerialLocale locale) {
      String var = Variant.startsWith(WtOfficeTools.MULTILINGUAL_LABEL) ? Variant.substring(WtOfficeTools.MULTILINGUAL_LABEL.length()) : Variant;
      String locVar = locale.Variant.startsWith(WtOfficeTools.MULTILINGUAL_LABEL) ? 
          locale.Variant.substring(WtOfficeTools.MULTILINGUAL_LABEL.length()) : locale.Variant;
      return Language.equals(locale.Language) && Country.equals(locale.Country) && var.equals(locVar);
    }

    /**
     * return the Locale as String
     */
    Locale toLocale() {
      return new Locale(Language, Country, Variant);
    }

    /**
     * return the Locale as String
     */
    Locale toLocaleWithoutLabel() {
      if (Variant.startsWith(WtOfficeTools.MULTILINGUAL_LABEL)) {
        return new Locale(Language, Country, Variant.substring(WtOfficeTools.MULTILINGUAL_LABEL.length()));
      }
      return new Locale(Language, Country, Variant);
    }

    /**
     * True if the Language is the same as Locale
     */
    boolean equalsLocale(Locale locale) {
      return ((locale == null || Language == null || Country == null || Variant == null) ? false
          : Language.equals(locale.Language) && Country.equals(locale.Country) && Variant.equals(locale.Variant));
    }

    boolean equalsLocale(SerialLocale locale) {
      return ((locale == null || Language == null || Country == null || Variant == null) ? false
          : Language.equals(locale.Language) && Country.equals(locale.Country) && Variant.equals(locale.Variant));
    }
    
  }

  public static class ChangedRange {
    final public int from;
    final public int to;
    final public int oldSize;
    final public int newSize;
    
    ChangedRange(int from, int to, int oldSize, int newSize) {
      this.from = from;
      this.to = to;
      this.oldSize = oldSize;
      this.newSize = newSize;
    }
  }

  public static class AnalysedText {
    final public List<AnalyzedSentence> analyzedSentences;
    final public List<String> sentences;
    final public String text;
    
    AnalysedText(List<AnalyzedSentence> analyzedSentences, List<String> sentences, String text) {
      this.analyzedSentences = analyzedSentences;
      this.sentences = sentences;
      this.text = text;
    }
  }
 
}

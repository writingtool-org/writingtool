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
package org.writingtool.aisupport;

import java.util.List;

import org.languagetool.Language;
import org.languagetool.UserConfig;
import org.writingtool.WtDocumentCache;
import org.writingtool.WtDocumentsHandler;
import org.writingtool.WtSingleDocument;
import org.writingtool.WtLanguageTool;
import org.writingtool.WtLinguisticServices;
import org.writingtool.WtSingleCheck;
import org.writingtool.WtTextLevelCheckQueue;
import org.writingtool.aisupport.WtAiErrorDetection.DetectionType;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtOfficeTools.DocumentType;

/**
 * Class of a queue to handle check of AI error detection
 * @since 1.0
 * @author Fred Kruse
 */
public class WtAiCheckQueue extends WtTextLevelCheckQueue {
  
  private final static int MIN_TEXT_LENGTH = 300;

  private int debugMode = WtOfficeTools.DEBUG_MODE_AI;   //  should be false except for testing
  private boolean singleParagraphMode = false;

  public WtAiCheckQueue(WtDocumentsHandler multiDocumentsHandler, boolean singleParagraphMode) {
    super(multiDocumentsHandler);
    this.singleParagraphMode = singleParagraphMode;
    WtMessageHandler.printToLogFile("AI Queue started");
  }
  
  public void setSingleParagraphMode(boolean singleParagraphMode) {
    this.singleParagraphMode = singleParagraphMode;
  }
  
  /**
   * Add a new entry to queue
   * add it only if the new entry is not identical with the last entry or the running
   */
   public void addQueueEntry(TextParagraph nTPara, WtSingleDocument document, boolean next) {
     try {
       String docId = document.getDocID();
       if (nTPara == null || nTPara.type < 0  || nTPara.number < 0 || docId == null || interruptCheck) {
         if (debugMode > 1) {
           WtMessageHandler.printToLogFile("AiCheckQueue: addQueueEntry: Return without add to queue: nCache = " + WtOfficeTools.CACHE_AI
               + ", nTPara = " + (nTPara == null ? "null" : ("(" + nTPara.number + "/" + nTPara.type + ")")) + ", docId = " + docId);
         }
         return;
       }
       QueueEntry queueEntry = createAiQueueEntry(nTPara, WtOfficeTools.CACHE_AI, 0, document, false);
       if (!textRuleQueue.isEmpty()) {
         synchronized(textRuleQueue) {
           for (int i = 0; i < textRuleQueue.size(); i++) {
             QueueEntry entry = textRuleQueue.get(i);
             if (entry.equals(queueEntry)) {
               if (debugMode > 1) {
                 WtMessageHandler.printToLogFile("AiCheckQueue: addQueueEntry: Entry removed: nCache = " + WtOfficeTools.CACHE_AI
                     + ", nTPara = (" + nTPara.number + "/" + nTPara.type + "), docId = " + docId);
               }
               textRuleQueue.remove(i);
               break;
             }
           }
         }
       }
       synchronized(textRuleQueue) {
         if (next) {
           textRuleQueue.add(0, queueEntry);
         } else {
           textRuleQueue.add(queueEntry);
         }
         if (debugMode > 1) {
           WtMessageHandler.printToLogFile("AiCheckQueue: addQueueEntry: Entry added: nCache = " + WtOfficeTools.CACHE_AI
               + ", nTPara = (" + nTPara.number + "/" + nTPara.type + "), docId = " + docId);
         }
       }
       interruptCheck = false;
       wakeupQueue();
     } catch (Throwable t) {
       WtMessageHandler.showError(t);
     }
   }
  
   /**
    *  get an entry for the next unchecked paragraphs
    */
  @Override
  protected QueueEntry getNextQueueEntry(TextParagraph nPara, String docId) {
    try {
      List<WtSingleDocument> documents = multiDocHandler.getDocuments();
      int nDoc = 0;
      for (int n = 0; n < documents.size(); n++) {
        if ((docId == null || docId.equals(documents.get(n).getDocID())) && !documents.get(n).isDisposed() && documents.get(n).getDocumentType() == DocumentType.WRITER) {
          QueueEntry queueEntry = documents.get(n).getNextAiQueueEntry(nPara);
          if (queueEntry != null) {
            return queueEntry;
          }
          nDoc = n;
          break;
        }
      }
      for (int i = nDoc + 1; i < documents.size(); i++) {
        if (!documents.get(i).isDisposed() && documents.get(i).getDocumentType() == DocumentType.WRITER) {
          QueueEntry queueEntry = documents.get(i).getNextAiQueueEntry(null);
          if (queueEntry != null) {
            return queueEntry;
          }
        }
      }
      for (int i = 0; i < nDoc; i++) {
        if (!documents.get(i).isDisposed() && documents.get(i).getDocumentType() == DocumentType.WRITER) {
          QueueEntry queueEntry = documents.get(i).getNextAiQueueEntry(null);
          if (queueEntry != null) {
            return queueEntry;
          }
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return null;
  }
  
  /**
   * initialize languagetool for text level iteration
   */
  @Override
  public void initLangtool(Language language) throws Throwable {
    lt = null;
    try {
      if (debugMode > 1) {
        WtMessageHandler.printToLogFile("TextLevelCheckQueue: initLangtool: language = " + (language == null ? "null" : language.getShortCodeWithCountryAndVariant()));
      }
      WtConfiguration config = WtDocumentsHandler.getConfiguration(language);
      WtLinguisticServices linguServices = multiDocHandler.getLinguisticServices();
      linguServices.setNoSynonymsAsSuggestions(config.noSynonymsAsSuggestions() || multiDocHandler.isTestMode());
      Language fixedLanguage = config.getDefaultLanguage();
      if (fixedLanguage != null) {
        language = fixedLanguage;
      }
      lt = new WtLanguageTool(language, config.getMotherTongue(),
          new UserConfig(config.getConfigurableValues(), linguServices), config, multiDocHandler.getExtraRemoteRules(), 
          !config.useLtSpellChecker(), false, multiDocHandler.isTestMode(), true);
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }
  
  /**
   *  run a queue entry for the specific document
   */
  @Override
  protected void runQueueEntry(QueueEntry qEntry, WtDocumentsHandler multiDocHandler, WtLanguageTool lt) throws Throwable {
    try {
      if (debugMode > 1) {
        WtMessageHandler.printToLogFile("Try to run AI Queue Entry for " + qEntry.nStart.number);
      }
      WtSingleDocument document = getSingleDocument(qEntry.docId);
      TextParagraph startPara = qEntry.nStart;
      if (document != null && !document.isDisposed() && startPara != null) {
        WtDocumentCache docCache = document.getDocumentCache();
        if (docCache != null) {
          WtAiErrorDetection aiError = new WtAiErrorDetection(document, multiDocHandler.getConfiguration(), lt);
          if (singleParagraphMode || startPara.type == WtDocumentCache.CURSOR_TYPE_UNKNOWN) {
            int nFPara = startPara.type == WtDocumentCache.CURSOR_TYPE_UNKNOWN ? startPara.number : docCache.getFlatParagraphNumber(startPara);
            if (debugMode > 1) {
              WtMessageHandler.printToLogFile("Run AI Queue Entry for " 
                  + ", nTPara = (" + startPara.number + "/" + startPara.type + "), docId = " + qEntry.docId
                  + ", nFPara = " + nFPara);
            }
            aiError.addAiRuleMatchesForParagraph(nFPara, DetectionType.GRAMMAR);
            if (multiDocHandler.useAiSuggestion()) {
              aiError.addAiRuleMatchesForParagraph(nFPara, DetectionType.REWRITE);
            }
          } else {
            TextParagraph endPara = qEntry.nEnd;
            String textToCheck = getDocAsString(startPara, docCache);
            aiError.addAiRuleMatchesForText(startPara, endPara, textToCheck, DetectionType.GRAMMAR);
            if (multiDocHandler.useAiSuggestion()) {
              aiError.addAiRuleMatchesForText(startPara, endPara, textToCheck, DetectionType.REWRITE);
            }
          }
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }
  
  /**
   * Create and give back a new queue entry
   * @throws Throwable 
   */
  public QueueEntry createAiQueueEntry(TextParagraph nTPara, int nCache, int nCheck, WtSingleDocument document, boolean overrideRunning) throws Throwable {
    TextParagraph nStart = !singleParagraphMode ? new TextParagraph(nTPara.type, getStartOfParaCheck(nTPara, document.getDocumentCache())) : nTPara;
    TextParagraph nEnd = !singleParagraphMode ? new TextParagraph(nTPara.type, getEndOfParaCheck(nTPara, document.getDocumentCache())) : nTPara;
    return (new QueueEntry(nStart, nEnd, nCache, nCheck, document.getDocID(), overrideRunning));
  }
  
  private static int getStartOfParaCheck(TextParagraph textParagraph, WtDocumentCache docCache) throws Throwable {
    if (textParagraph.type == WtDocumentCache.CURSOR_TYPE_UNKNOWN) {
      return textParagraph.number;
    }
    int startChapter = docCache.getStartOfParaCheck(textParagraph, -1, false, false, false);
    int endChapter = docCache.getEndOfParaCheck(textParagraph, -1, false, false, false);
    int tType = textParagraph.type;
    int j = 0;
    for (int i = startChapter; i < endChapter;) {
      int len = 0;
      for (j = i; j < endChapter && len < MIN_TEXT_LENGTH; j++) {
        len += docCache.getTextParagraph(new TextParagraph(tType, j)).length();
        if(j >= textParagraph.number) {
          return i;
        }
      }
      i = j;
    }
    return startChapter;
  }
  
  private static int getEndOfParaCheck(TextParagraph textParagraph, WtDocumentCache docCache) throws Throwable {
    if (textParagraph.type == WtDocumentCache.CURSOR_TYPE_UNKNOWN) {
      return textParagraph.number + 1;
    }
    int startChapter = docCache.getStartOfParaCheck(textParagraph, -1, false, false, false);
    int endChapter = docCache.getEndOfParaCheck(textParagraph, -1, false, false, false);
    int tType = textParagraph.type;
    int j = 0;
    for (int i = startChapter; i < endChapter;) {
      int len = 0;
      for (j = i; j < endChapter && len < MIN_TEXT_LENGTH; j++) {
        len += docCache.getTextParagraph(new TextParagraph(tType, j)).length();
      }
      if(j >= textParagraph.number) {
        return j;
      }
      i = j;
    }
    return endChapter;
  }

  /**
   * Gives Back the StartPosition of Paragraph
   */
  public static int getStartOfParagraph(int nPara, TextParagraph textParagraph, WtDocumentCache docCache) throws Throwable {
    if (nPara < 0 || docCache.textSize(textParagraph) <= nPara) {
      return -1;
    }
    int startPos = getStartOfParaCheck(textParagraph, docCache);
    if (startPos < 0) {
      return -1;
    }
    int pos = 0;
    for (int i = startPos; i < nPara; i++) {
      TextParagraph tPara = new TextParagraph(textParagraph.type, i);
      pos += WtSingleCheck.removeFootnotes(docCache.getTextParagraph(tPara), 
              docCache.getTextParagraphFootnotes(tPara), docCache.getTextParagraphDeletedCharacters(tPara)).length() + WtOfficeTools.NUMBER_PARAGRAPH_CHARS;
    }
    return pos;
  }

  /**
   * Gives Back the full Text as String sorted by cursor types
   */
  private String getDocAsString(TextParagraph textParagraph, WtDocumentCache docCache) {
    try {
      if (textParagraph.type == WtDocumentCache.CURSOR_TYPE_UNKNOWN) {
        return docCache.getFlatParagraph(textParagraph.number);
      }
      if (singleParagraphMode) {
        return docCache.getTextParagraph(textParagraph);
      }
      int startPos = getStartOfParaCheck(textParagraph, docCache);
      int endPos = getEndOfParaCheck(textParagraph, docCache);
      StringBuilder docText;
      TextParagraph startPara = new TextParagraph(textParagraph.type, startPos);
      if (startPos < 0 || endPos < 0) {
        return "";
      }
      docText = new StringBuilder(WtDocumentCache.fixLinebreak(WtSingleCheck.removeFootnotes(docCache.getTextParagraph(startPara),
          docCache.getTextParagraphFootnotes(startPara), docCache.getTextParagraphDeletedCharacters(startPara))));
      for (int i = startPos + 1; i < endPos; i++) {
        TextParagraph tPara = new TextParagraph(textParagraph.type, i);
        docText.append(WtOfficeTools.END_OF_PARAGRAPH).append(WtDocumentCache.fixLinebreak(WtSingleCheck
            .removeFootnotes(docCache.getTextParagraph(tPara), docCache.getTextParagraphFootnotes(tPara), docCache.getTextParagraphDeletedCharacters(tPara))));
      }
      return docText.toString();
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      return "";
    }
  }

}

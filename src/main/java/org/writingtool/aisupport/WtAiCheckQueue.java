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
package org.writingtool.aisupport;

import java.util.List;

import org.languagetool.Language;
import org.languagetool.UserConfig;
import org.writingtool.WtDocumentCache;
import org.writingtool.WtDocumentsHandler;
import org.writingtool.WtSingleDocument;
import org.writingtool.WtLanguageTool;
import org.writingtool.WtLinguisticServices;
import org.writingtool.WtTextLevelCheckQueue;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtOfficeTools.DocumentType;

/**
 * Class of a queue to handle check of AI error detection
 * @since 6.5
 * @author Fred Kruse
 */
public class WtAiCheckQueue extends WtTextLevelCheckQueue {

  private int debugMode = WtOfficeTools.DEBUG_MODE_AI;   //  should be false except for testing

  public WtAiCheckQueue(WtDocumentsHandler multiDocumentsHandler) {
    super(multiDocumentsHandler);
    WtMessageHandler.printToLogFile("AI Queue started");
  }
  
  /**
   * Add a new entry to queue
   * add it only if the new entry is not identical with the last entry or the running
   */
   public void addQueueEntry(TextParagraph nTPara, String docId, boolean next) {
     try {
       if (nTPara == null || nTPara.type < 0  || nTPara.number < 0 || docId == null || interruptCheck) {
         if (debugMode > 1) {
           WtMessageHandler.printToLogFile("AiCheckQueue: addQueueEntry: Return without add to queue: nCache = " + WtOfficeTools.CACHE_AI
               + ", nTPara = " + (nTPara == null ? "null" : ("(" + nTPara.number + "/" + nTPara.type + ")")) + ", docId = " + docId);
         }
         return;
       }
       QueueEntry queueEntry = new QueueEntry(nTPara, nTPara, WtOfficeTools.CACHE_AI, 0, docId, false);
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
      TextParagraph nTPara = qEntry.nStart;
      if (document != null && !document.isDisposed() && nTPara != null) {
        WtDocumentCache docCache = document.getDocumentCache();
        if (docCache != null) {
          int nFPara = docCache.getFlatParagraphNumber(nTPara);
          if (debugMode > 1) {
            WtMessageHandler.printToLogFile("Run AI Queue Entry for " 
                + ", nTPara = (" + nTPara.number + "/" + nTPara.type + "), docId = " + qEntry.docId
                + ", nFPara = " + nFPara);
          }
          WtAiErrorDetection aiError = new WtAiErrorDetection(document, multiDocHandler.getConfiguration(), lt);
          aiError.addAiRuleMatchesForParagraph(nFPara);
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }

}

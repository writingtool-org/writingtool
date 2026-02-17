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

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.languagetool.Language;
import org.languagetool.UserConfig;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtOfficeTools.DocumentType;

import com.sun.star.lang.Locale;

/**
 * Class of a queue to handle parallel check of text level rules
 * @since 1.0
 * @author Fred Kruse
 */
public class WtTextLevelCheckQueue {
  
  public static final int NO_FLAG = 0;
  public static final int RESET_FLAG = 1;
  public static final int STOP_FLAG = 2;
  public static final int DISPOSE_FLAG = 3;

  private static final int HEAP_CHECK_INTERVAL = 50;
  private static final int MAX_CHECK_PER_THREAD = 50;

  protected List<QueueEntry> textRuleQueue = Collections.synchronizedList(new ArrayList<QueueEntry>());  //  Queue to check text rules in a separate thread
//  private Object queueWakeup = new Object();
  protected WtDocumentsHandler multiDocHandler;
  WtSingleDocument document;

  private QueueIterator queueIterator = null;
  private TextParagraph lastStart = null;
  private TextParagraph lastEnd = null;
  private int lastCache = -1;
  private String lastDocId = null;
  protected WtLanguageTool lt;
  protected Language lastLanguage = null;
  protected boolean interruptCheck = false;
  private boolean queueRuns = false;
//  private boolean queueWaits = false;
  
  private int numSinceHeapTest = 0;

  private static boolean debugMode = false;   //  should be false except for testing
  private static boolean debugModeTm;         // time measurement should be false except for testing
  
  protected WtTextLevelCheckQueue(WtDocumentsHandler multiDocumentsHandler) {
    multiDocHandler = multiDocumentsHandler;
//    queueIterator = new QueueIterator();
//    queueIterator.start();
    debugMode = WtOfficeTools.DEBUG_MODE_TQ;
    debugModeTm = WtOfficeTools.DEBUG_MODE_TM;
  }
 
 /**
  * Add a new entry to queue
  * add it only if the new entry is not identical with the last entry or the running
  */
  public void addQueueEntry(TextParagraph nStart, TextParagraph nEnd, int nCache, int nCheck,
      String docId, boolean overrideRunning) throws Throwable {
    if (nStart == null || nEnd == null || nStart.type != nEnd.type || nStart.number < 0 
        || nEnd.number <= nStart.number || nCache < 0 || docId == null
        || interruptCheck) {
      if (debugMode) {
        WtMessageHandler.printToLogFile(nameOfQueue() + ": addQueueEntry: Return without add to queue: nCache = " + nCache
            + ", nStart = " + nStart + ", nEnd = " + nEnd 
            + ", nCheck = " + nCheck + ", docId = " + docId + ", overrideRunning = " + overrideRunning);
      }
      return;
    }
    QueueEntry queueEntry = new QueueEntry(nStart, nEnd, nCache, nCheck, docId, overrideRunning);
    if (!textRuleQueue.isEmpty()) {
      if (!overrideRunning && lastStart != null && nStart.type == lastStart.type && nStart.number >= lastStart.number 
          && nEnd.number <= lastEnd.number && nCache == lastCache && lastDocId != null && docId.equals(lastDocId)) {
        return;
      }
      synchronized(textRuleQueue) {
        for (int i = 0; i < textRuleQueue.size(); i++) {
          QueueEntry entry = textRuleQueue.get(i);
          if (entry.isObsolete(queueEntry)) {
            if (!overrideRunning) {
              return;
            } else {
              textRuleQueue.remove(i);
              if (debugMode) {
                WtMessageHandler.printToLogFile(nameOfQueue() + ": addQueueEntry: remove queue entry: docId = " + entry.docId 
                    + ", nStart.type = " + entry.nStart.type + ", nStart.number = " + entry.nStart.number + ", nEnd.number = " + entry.nEnd.number 
                    + ", nCache = " + entry.nCache + ", nCheck = " + entry.nCheck + ", overrideRunning = " + entry.overrideRunning);
              }
            }
          }
        }
        if (overrideRunning) {
          for (int i = 0; i < textRuleQueue.size(); i++) {
            QueueEntry entry = textRuleQueue.get(i);
            if(!entry.isEqualButSmallerCacheNumber(queueEntry)) {
              textRuleQueue.add(i, queueEntry);
              if (debugMode) {
                WtMessageHandler.printToLogFile(nameOfQueue() + ": addQueueEntry: add queue entry at position: " + i + "; docId = " + queueEntry.docId 
                    + ", nStart.type = " + queueEntry.nStart.type + ", nStart.number = " + queueEntry.nStart.number + ", nEnd.number = " + queueEntry.nEnd.number 
                    + ", nCache = " + queueEntry.nCache + ", nCheck = " + queueEntry.nCheck + ", overrideRunning = " + queueEntry.overrideRunning);
              }
              return;
            }
          }
        }
      }
    }
    synchronized(textRuleQueue) {
      textRuleQueue.add(queueEntry);
      if (debugMode) {
        WtMessageHandler.printToLogFile(nameOfQueue() + ": addQueueEntry: add queue entry at position: " + (textRuleQueue.size() - 1) + "; docId = " + queueEntry.docId 
            + ", nStart.type = " + queueEntry.nStart.type + ", nStart.number = " + queueEntry.nStart.number + ", nEnd.number = " + queueEntry.nEnd.number 
            + ", nCache = " + queueEntry.nCache + ", nCheck = " + queueEntry.nCheck + ", overrideRunning = " + queueEntry.overrideRunning);
      }
    }
    interruptCheck = false;
    wakeupQueue();
  }
  
  /**
   * Create and give back a new queue entry
   */
  public QueueEntry createQueueEntry(TextParagraph nStart, TextParagraph nEnd, int nCache, int nCheck, String docId, boolean overrideRunning) {
    return (new QueueEntry(nStart, nEnd, nCache, nCheck, docId, overrideRunning));
  }
  
  public QueueEntry createQueueEntry(TextParagraph nStart, TextParagraph nEnd, int cacheNum, int nCheck, String docId) {
    return createQueueEntry(nStart, nEnd, cacheNum, nCheck, docId, false);
  }
  
  /**
   * for debugMode set name of Queue
   */
  protected String nameOfQueue() {
    return "TextLevelCheckQueue";
  }
  
  /**
   * wake up the waiting iteration of the queue
   */
  protected void wakeupQueue() {
    try {
//    synchronized(queueWakeup) {
      if (debugMode) {
        WtMessageHandler.printToLogFile(nameOfQueue() + ": wakeupQueue: wake queue");
      }
//      queueWakeup.notify();
//    }
      if (queueIterator == null) {
        queueIterator = new QueueIterator();
        queueIterator.start();
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }

  /**
   * wake up the waiting iteration of the queue for a specific document
   */
  public void wakeupQueue(String docId) throws Throwable {
    if (lastDocId == null) {
      lastDocId = docId;
    }
    wakeupQueue();
  }

  /**
   * Set a stop flag to get a definite ending of the iteration
   */
  public void setStop(boolean disposed) throws Throwable {
    if (queueRuns) {
      interruptCheck = true;
      if (disposed && document != null) {
        document.dispose(true);
      }
      QueueEntry queueEntry = new QueueEntry();
      queueEntry.setStop();
      if (debugMode) {
        WtMessageHandler.printToLogFile(nameOfQueue() + ": setStop: stop queue");
      }
      textRuleQueue.add(0, queueEntry);
    }
//    wakeupQueue();
  }
  
  /**
   * Reset the queue
   * all entries are removed; LanguageTool is new initialized
   */
  public void setReset() throws Throwable {
    if (queueRuns) {
      interruptCheck = true;
      QueueEntry queueEntry = new QueueEntry();
      queueEntry.setReset();
      if (debugMode) {
        WtMessageHandler.printToLogFile(nameOfQueue() + ": setReset: reset queue");
      }
      synchronized(textRuleQueue) {
        textRuleQueue.clear();
      }
      textRuleQueue.add(queueEntry);
    }
    lt = null;
    wakeupQueue();
  }
  
  /**
   * remove all entries for the disposed docId (gone document)
   * @param docId
   */
  public void interruptCheck(String docId, boolean wait) throws Throwable {
    if (debugMode) {
      WtMessageHandler.printToLogFile(nameOfQueue() + ": interruptCheck: interrupt queue");
    }
    if (!textRuleQueue.isEmpty()) {
      synchronized(textRuleQueue) {
        for (int i = textRuleQueue.size() - 1; i >= 0; i--) {
          QueueEntry queueEntry = textRuleQueue.get(i);
          if (docId.equals(queueEntry.docId)) {
            textRuleQueue.remove(queueEntry);
          }
        }
      }
    }
//    if (wait && !queueWaits && lastStart != null && lastDocId != null && lastDocId.equals(docId)) {
    if (wait && queueRuns && lastStart != null && lastDocId != null && lastDocId.equals(docId)) {
      lastDocId = null;
    }
  }
  
  /**
   *  get the document by ID
   */
  protected WtSingleDocument getSingleDocument(String docId) throws Throwable {
    for (WtSingleDocument document : multiDocHandler.getDocuments()) {
      if (docId.equals(document.getDocID())) {
        return document;
      }
    }
    return null;
  }
  
  /**
   *  get language of document by ID
   */
  protected Language getLanguage(String docId, TextParagraph nStart) throws Throwable {
    WtSingleDocument document = getSingleDocument(docId);
    WtDocumentCache docCache = null;
    if (document != null) {
      docCache = document.getDocumentCache();
      if (docCache != null && nStart.number < docCache.textSize(nStart)) {
        if (docCache.isAutomaticGenerated(docCache.getFlatParagraphNumber(nStart), true)) {
          return null;
        }
        Locale locale = nStart.type == WtDocumentCache.CURSOR_TYPE_UNKNOWN && nStart.number >= 0 ? 
                docCache.getFlatParagraphLocale(nStart.number) : docCache.getTextParagraphLocale(nStart);
        if (locale != null && WtDocumentsHandler.hasLocale(locale)) {
          return WtDocumentsHandler.getLanguage(locale);
        }
        WtMessageHandler.printToLogFile(nameOfQueue() + ": getLanguage: return null: locale = " 
            + (locale == null ? "null" : WtOfficeTools.localeToString(locale))
            + ", nStart.type = " + nStart.type + ", nStart.number = " + nStart.number);
      }
    }
    if (debugMode) {
      if (document == null) {
        WtMessageHandler.printToLogFile(nameOfQueue() + ": getLanguage: document == null: return null");
      } else if (docCache == null) {
        WtMessageHandler.printToLogFile(nameOfQueue() + ": getLanguage: docCache == null: return null");
      } else if (nStart.number >= docCache.textSize(nStart)) {
        WtMessageHandler.printToLogFile(nameOfQueue() + ": getLanguage: nStart.number >= docCache.textSize(nStart): return null");
      }
    }
    return null;
  }
  
  /**
   * initialize languagetool for text level iteration
   */
  public void initLangtool(Language language) throws Throwable {
    if (debugMode) {
      WtMessageHandler.printToLogFile(nameOfQueue() + ": initLangtool: language = " + (language == null ? "null" : language.getShortCodeWithCountryAndVariant()));
    }
    lt = null;
    try {
      WtConfiguration config = WtDocumentsHandler.getConfiguration(language);
      WtLinguisticServices linguServices = multiDocHandler.getLinguisticServices();
      linguServices.setNoSynonymsAsSuggestions(config.noSynonymsAsSuggestions() || multiDocHandler.isTestMode());
      Language fixedLanguage = config.getDefaultLanguage();
      if (fixedLanguage != null) {
        language = fixedLanguage;
      }
      lt = new WtLanguageTool(language, config.getMotherTongue(),
          new UserConfig(config.getConfigurableValues(), linguServices), config, multiDocHandler.getExtraRemoteRules(), 
          !config.useLtSpellChecker(), false, multiDocHandler.isTestMode());
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }
  
  /**
   * gives back information if queue is interrupted
   */
  public boolean isInterrupted() {
    return interruptCheck;
  }
  
  /**
   * gives back information if queue is running
   */
  public boolean isRunning() {
    return queueRuns;
  }
  
  /**
   * gives back information if queue is waiting
   */
  public boolean isWaiting() {
//    return queueWaits;
    return !queueRuns;
  }
  
  /**
   *  get an entry for the next unchecked paragraphs
   */
  protected QueueEntry getNextQueueEntry(TextParagraph nPara, String docId) {
    List<WtSingleDocument> documents = multiDocHandler.getDocuments();
    int nDoc = 0;
    for (int n = 0; n < documents.size(); n++) {
      if ((docId == null || docId.equals(documents.get(n).getDocID())) && !documents.get(n).isDisposed() && documents.get(n).getDocumentType() == DocumentType.WRITER) {
        QueueEntry queueEntry = documents.get(n).getNextQueueEntry(nPara);
        if (queueEntry != null) {
          return queueEntry;
        }
        nDoc = n;
        break;
      }
    }
    for (int n = 0; n < documents.size(); n++) {
      if (docId != null && docId.equals(documents.get(n).getDocID()) && !documents.get(n).isDisposed() && documents.get(n).getDocumentType() == DocumentType.WRITER) {
        QueueEntry queueEntry = documents.get(n).getQueueEntryForChangedParagraph();
        if (queueEntry != null) {
          return queueEntry;
        }
        nDoc = n;
        break;
      }
    }
    for (int i = nDoc + 1; i < documents.size(); i++) {
      if (!documents.get(i).isDisposed() && documents.get(i).getDocumentType() == DocumentType.WRITER) {
        QueueEntry queueEntry = documents.get(i).getNextQueueEntry(null);
        if (queueEntry != null) {
          return queueEntry;
        }
      }
    }
    for (int i = 0; i < nDoc; i++) {
      if (!documents.get(i).isDisposed() && documents.get(i).getDocumentType() == DocumentType.WRITER) {
        QueueEntry queueEntry = documents.get(i).getNextQueueEntry(null);
        if (queueEntry != null) {
          return queueEntry;
        }
      }
    }
    return null;
  }
  
  /**
   * run heap space test, in intervals
   */
  private boolean testHeapSpace() {
    if (numSinceHeapTest > HEAP_CHECK_INTERVAL) {
      numSinceHeapTest = 0;
      if (!multiDocHandler.isEnoughHeapSpace()) {
        return false;
      }
    } else {
      numSinceHeapTest++;
    }
    return true;
  }

  /**
   *  run a queue entry for the specific document
   */
  protected void runQueueEntry(QueueEntry qEntry, WtDocumentsHandler multiDocHandler, WtLanguageTool lt) throws Throwable {
    if (testHeapSpace()) {
      document = getSingleDocument(qEntry.docId);
      if (document != null && !document.isDisposed()) {
        if (debugMode) {
          WtMessageHandler.printToLogFile(nameOfQueue() + ": runQueueEntry: nstart = " + qEntry.nStart.number + "; nEnd = "  + qEntry.nEnd.number 
              + "; nCache = "  + qEntry.nCache + "; nCheck = "  + qEntry.nCheck + "; overrideRunning = "  + qEntry.overrideRunning);
        }
        document.runQueueEntry(qEntry.nStart, qEntry.nEnd, qEntry.nCache, qEntry.nCheck, qEntry.overrideRunning, lt);
        document = null;
        multiDocHandler.setCacheStatusColor(document);
      }
    } else {
      WtMessageHandler.printToLogFile("Warning: Not enough heap space; text level queue stopped!");
      setStop(false);
    }
  }
  
  /**
   *  run a queue entry to get synonyms (not used in text level check queue)
   */
  protected void runQueueEntrySynonyms(QueueEntry qEntry, WtDocumentsHandler multiDocHandler, WtLanguageTool lt) throws Throwable { }
  
  /**
   * Internal class to store queue entries
   */
  protected static class QueueEntry {
    public TextParagraph nStart;
    public TextParagraph nEnd;
    public int nCache;
    public int nCheck;
    public String docId;
    public WtProofreadingError error;
    public boolean overrideRunning;
    public int special = WtTextLevelCheckQueue.NO_FLAG;

    public QueueEntry(TextParagraph nStart, TextParagraph nEnd, int nCache, int nCheck, String docId, boolean overrideRunning) {
      this(nStart, nEnd, nCache, nCheck, docId, overrideRunning, null);
    }
    
    public QueueEntry(TextParagraph nStart, int nCache, String docId, WtProofreadingError error) {
      this(nStart, nStart, nCache, 0, docId, false, error);
    }
    
    public QueueEntry(TextParagraph nStart, TextParagraph nEnd, int nCache, int nCheck, String docId, 
        boolean overrideRunning, WtProofreadingError error) {
      this.nStart = nStart;
      this.nEnd = nEnd;
      this.nCache = nCache;
      this.nCheck = nCheck;
      this.docId = docId;
      this.overrideRunning = overrideRunning;
      this.error = error;
    }
    
    QueueEntry(TextParagraph nStart, TextParagraph nEnd, int nCache, int nCheck, String docId) {
      this(nStart, nEnd, nCache, nCheck, docId, false);
    }
    
    QueueEntry() {
    }

    /**
     * Set reset flag
     */
    void setReset() {
      special = WtTextLevelCheckQueue.RESET_FLAG;
    }
    
    /**
     * Set stop flag
     */
    void setStop() {
      special = WtTextLevelCheckQueue.STOP_FLAG;
    }
    
    /**
     * Set dispose flag
     */
    void setDispose(String docId) {
      special = WtTextLevelCheckQueue.DISPOSE_FLAG;
    }
    
    /**
     * Define equal queue entries
     */
    @Override
    public boolean equals(Object o) {
      if (o == null || !(o instanceof QueueEntry)) {
        return false;
      }
      QueueEntry e = (QueueEntry) o;
      if (nStart == null || nEnd == null || e.nStart == null || e.nEnd == null) {
        return false;
      }
      if ((error == null && e.error != null) || (error != null && e.error == null)) {
        return false;
      }
      if (nStart.type == e.nStart.type && nStart.number == e.nStart.number && nEnd.number == e.nEnd.number
          && nCache == e.nCache && nCheck == e.nCheck && docId.equals(e.docId)
          && ((error == null && e.error == null) || error.equals(e.error)) ) {
        return true;
      }
      return false;
    }

    /**
     * entry is equal but number of cache is smaller then new entry e
     */
    public boolean isEqualButSmallerCacheNumber(QueueEntry e) {
      if (e == null || nStart == null || nEnd == null 
          || e.nStart == null || e.nEnd == null || nStart.type != e.nStart.type) {
        return false;
      }
      if (nStart.number >= e.nStart.number && nEnd.number <= e.nEnd.number && nCache < e.nCache && docId.equals(e.docId)) {
        return true;
      }
      return false;
    }

    /**
     * entry is obsolete and should be replaced by new entry e
     */
    public boolean isObsolete(QueueEntry e) {
      if (e == null || nStart == null  || nEnd == null || e.nStart == null || e.nEnd == null
          || nCheck != e.nCheck || nCache != e.nCache || nStart.type != e.nStart.type || docId == null || !docId.equals(e.docId)) {
        return false;
      }
      if (nCheck < -1 || (nCheck == -1 && e.nStart.number >= nStart.number && e.nStart.number <= nEnd.number) 
          || (nCheck >= 0 && nStart.number == e.nStart.number && nEnd.number == e.nEnd.number)) {
        return true;
      }
      return false;
    }

  }

  /**
   * class for automatic iteration of the queue
   */
  private class QueueIterator extends Thread {
    
    private int numCheck = 0;

      
    public QueueIterator() {
    }
    
    /**
     * Run queue for check with text
     */
    @Override
    public void run() {
      try {
        long startTime = 0;
        queueRuns = true;
        if (debugMode) {
          WtMessageHandler.printToLogFile(nameOfQueue() + ": run: queue started");
        }
        while (numCheck < MAX_CHECK_PER_THREAD) {
//          queueWaits = false;
          if (interruptCheck) {
/*            
            MessageHandler.printToLogFile(nameOfQueue() + ": run: Interrupt check - queue ended");
            textRuleQueue.clear();
            interruptCheck = false;
            queueRuns = false;
            queueIterator = null;
            return;
*/
            try {
              Thread.sleep(50);
            } catch (InterruptedException e) {
              WtMessageHandler.printException(e);
            }
            interruptCheck = false;
            continue;
          }
          if (textRuleQueue.isEmpty()) {
            synchronized(textRuleQueue) {
              if (lastDocId != null) {
                QueueEntry queueEntry = null;
                try {
                  if (debugModeTm) {
                    startTime = System.currentTimeMillis();
                  }
                  if (!interruptCheck) {
                    queueEntry = getNextQueueEntry(lastStart, lastDocId);
                  }
                  if (debugModeTm) {
                    long runTime = System.currentTimeMillis() - startTime;
                    if (runTime > WtOfficeTools.TIME_TOLERANCE) {
                      WtMessageHandler.printToLogFile("WtTextLevelCheckQueue: Time to run Text Level Check Queue (get Next Queue Entry): " + runTime);
                    }
                  }
                } catch (Throwable e) {
                  //  there may be exceptions because of timing problems
                  //  catch them and write to log file but don't stop the queue
                  if (debugMode) {
                    WtMessageHandler.showError(e);
                  } else {
                    WtMessageHandler.printException(e);
                  }
                }
                if (queueEntry != null && !interruptCheck) {
                  textRuleQueue.add(queueEntry);
                  queueEntry = null;
                  continue;
                }
              }
            }
//            synchronized(queueWakeup) {
              try {
                if (debugMode) {
                  WtMessageHandler.printToLogFile(nameOfQueue() + ": run: queue waits");
                }
                lastStart = null;
                lastEnd = null;
//                queueWaits = true;
//                queueWakeup.wait();
                queueRuns = false;
                queueIterator = null;
                return;
              } catch (Throwable e) {
                WtMessageHandler.showError(e);
                queueRuns = false;
                queueIterator = null;
                return;
              }
//            }
          } else {
            QueueEntry queueEntry;
            synchronized(textRuleQueue) {
              if (!textRuleQueue.isEmpty()) {
                queueEntry = textRuleQueue.get(0);
                textRuleQueue.remove(0);
              } else {
                continue;
              }
            }
            if (queueEntry.special == STOP_FLAG) {
              if (debugMode) {
                WtMessageHandler.printToLogFile(nameOfQueue() + ": run: queue ended");
              }
              queueRuns = false;
              queueIterator = null;
              interruptCheck = false;
              return;
            } else if (queueEntry.special == RESET_FLAG) {
              if (debugMode) {
                WtMessageHandler.printToLogFile(nameOfQueue() + ": run: reset queue");
              }
//              synchronized(queueWakeup) {
                try {
                  if (debugMode) {
                    WtMessageHandler.printToLogFile(nameOfQueue() + ": run: queue waits");
                  }
                  lastStart = null;
                  lastEnd = null;
                  lastLanguage = null;
//                  queueWaits = true;
                  interruptCheck = false;
                  lt = null;
//                  queueWakeup.wait();
                  continue;
                } catch (Throwable e) {
                  WtMessageHandler.showError(e);
                  queueRuns = false;
                  queueIterator = null;
                  return;
                }
//              }
            } else {
              if (debugMode) {
                WtMessageHandler.printToLogFile(nameOfQueue() + ": run: run queue entry: docId = " + queueEntry.docId + ", nStart.type = " + queueEntry.nStart.type 
                    + ", nStart.number = " + queueEntry.nStart.number + ", nEnd.number = " + queueEntry.nEnd.number 
                    + ", nCheck = " + queueEntry.nCheck + ", overrideRunning = " + queueEntry.overrideRunning);
                if (queueEntry.nStart.number + 1 == queueEntry.nEnd.number) {
                  WtSingleDocument document = getSingleDocument(queueEntry.docId);
                  WtMessageHandler.printToLogFile(nameOfQueue() + ": run: Paragraph(" + queueEntry.nStart.number + "): '" 
                      + document.getDocumentCache().getTextParagraph(queueEntry.nStart) + "'");
                }
              }
              try {
                if (debugModeTm) {
                  startTime = System.currentTimeMillis();
                }
                Language entryLanguage = null;
                if (!interruptCheck) {
                  entryLanguage = getLanguage(queueEntry.docId, queueEntry.nStart);
                }
                if (entryLanguage != null) {
                  if (lt == null || lastLanguage == null || !lastLanguage.equals(entryLanguage)) {
                    lastLanguage = entryLanguage;
                    if (!interruptCheck) {
                      initLangtool(lastLanguage);
                      if (lt == null) {
                        WtMessageHandler.printToLogFile(nameOfQueue() + ": run: lt == null: lastLanguage == " 
                            + (lastLanguage == null ? "null" : lastLanguage.getShortCodeWithCountryAndVariant()));
                        queueRuns = false;
                        queueIterator = null;
                        return;
                      }
                    }
                    if (!interruptCheck) {
                      lt.activateTextRulesByIndex(queueEntry.nCache);
                    }
                  } else if (lastCache != queueEntry.nCache && !interruptCheck) {
                    lt.activateTextRulesByIndex(queueEntry.nCache);
                  }
                }
                lastDocId = queueEntry.docId;
  	            lastStart = queueEntry.nStart;
  	            lastEnd = queueEntry.nEnd;
  	            lastCache = queueEntry.nCache;
  	            // entryLanguage == null: language is not supported by LT
  	            // lt is set to null - results in empty entry in result cache
                if (debugMode && entryLanguage == null) {
                  WtMessageHandler.printToLogFile(nameOfQueue() + ": run: entryLanguage == null: lt set to null"); 
                }
                if (!interruptCheck) {
                  numCheck++;
                  if (queueEntry.error == null) {
                    runQueueEntry(queueEntry, multiDocHandler, entryLanguage == null ? null : lt);
                  } else {
                    runQueueEntrySynonyms(queueEntry, multiDocHandler, entryLanguage == null ? null : lt);
                  }
                }
                queueEntry = null;
                if (debugModeTm) {
                  long runTime = System.currentTimeMillis() - startTime;
                  if (runTime > WtOfficeTools.TIME_TOLERANCE) {
                    WtMessageHandler.printToLogFile("WtTextLevelCheckQueue: Time to run Text Level Check Queue (run Queue Entry): " + runTime);
                  }
                }
              } catch (Throwable e) {
                //  there may be exceptions because of timing problems
                //  catch them and write to log file but don't stop the queue
                if (debugMode) {
                  WtMessageHandler.showError(e);
                } else {
                  WtMessageHandler.printException(e);
                }
              }
            }
          }
        }
      } catch (Throwable e) {
        WtMessageHandler.showError(e);
        queueRuns = false;
        queueIterator = null;
      }
      queueRuns = false;
      queueIterator = null;
      if (numCheck >= MAX_CHECK_PER_THREAD) {
        wakeupQueue();
      }
    }
    
  }
    
}

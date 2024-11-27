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

import java.io.File;
import java.util.Map;
import java.util.ResourceBundle;

import org.writingtool.WtDocumentCache;
import org.writingtool.WtSingleDocument;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.WtDocumentsHandler.WaitDialogThread;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;

/**
 * Class for a new document filled by AI
 * @since WT 1.0
 * @author Fred Kruse
 */

public class WtAiTextToSpeech extends Thread {
  
  private static final ResourceBundle messages = WtOfficeTools.getMessageBundle();

  private boolean debugMode = WtOfficeTools.DEBUG_MODE_AI;   //  should be false except for testing

  public final static String WAIT_TITLE = messages.getString("loAiWaitDialogTitle");
  public final static String WAIT_MESSAGE = messages.getString("loAiWaitDialogMessage");
  
  public final static String TMP_DIRNAME = "audioOut";
  
  private WaitDialogThread waitDialog = null;
  private final WtSingleDocument document;
  private final WtDocumentCache docCache;
  private final Map<Integer, Integer> headingMap;
//  private XComponentContext xContext;
//  private WtDocumentCursorTools docCursor;
//  private XComponent xToComponent;
  
  public WtAiTextToSpeech(WtSingleDocument document, ResourceBundle messages) {
//-    this.xContext = xContext;
    this.document = document;
    docCache = document.getDocumentCache();
    headingMap = docCache.getHeadingMap();
  }
  
  @Override
  public void run() {
    try {
      waitDialog = new WaitDialogThread(WAIT_TITLE, WAIT_MESSAGE);
      waitDialog.start();
/*      
      WtConfiguration config = document.getMultiDocumentsHandler().getConfiguration();
      XComponent xComponent = document.getXComponent();
      WtViewCursorTools viewCursor = new WtViewCursorTools(xComponent);
      TextParagraph tPara = viewCursor.getViewCursorParagraph();
      WtDocumentCache docCache = document.getDocumentCache();
      String text = docCache.getTextParagraph(tPara);
*/
      WtConfiguration config = document.getMultiDocumentsHandler().getConfiguration();
      WtAiRemote aiRemote = new WtAiRemote(document.getMultiDocumentsHandler(), config);
      String audioDir = getAudioDir();
      int nParaStart = 0;
      int nFile = 0;
      int maxPara = docCache.textSize(WtDocumentCache.CURSOR_TYPE_TEXT);
      waitDialog.initializeProgressBar(0, maxPara);
      while ((nParaStart = createAudioFile(nParaStart, nFile, audioDir, aiRemote)) < maxPara) {
        nFile++;
        waitDialog.setValueForProgressBar(nParaStart, true);
      }
      waitDialog.close();
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
  }
  
  private String getAudioDir() {
    File audioDir = new File(WtOfficeTools.getWtConfigDir(), TMP_DIRNAME);
    if (audioDir.exists()) {
      File newDir;
      int i = 1;
      while ((newDir = new File(audioDir.getAbsolutePath() + "." + i)).exists()) {
        i++;
      };
      audioDir.renameTo(newDir);
      audioDir = new File(WtOfficeTools.getWtConfigDir(), TMP_DIRNAME);
    }
    audioDir.mkdirs();
    return audioDir.getAbsolutePath() + "/";
  }
  
  /** is Header of chapter
   * NOTE: nPara is number of text paragraph
   */
  private boolean isRealHeader(int nPara) {
    return (headingMap.containsKey(nPara) && headingMap.get(nPara) > 0);
  }
  
  private int createAudioFile(int nParaStart, int nFile, String audioDir, WtAiRemote aiRemote) {
    int n = nParaStart;
    String filename;
    try {
    if (isRealHeader(nParaStart)) {
      filename = docCache.getTextParagraph(new TextParagraph(WtDocumentCache.CURSOR_TYPE_TEXT, nParaStart));
      filename.replaceAll("[^\\w\\d]", "_");
    } else {
      filename = "Intro";
    }
    filename = audioDir + (nFile < 10 ? "0" : "") + nFile + "_" + filename;
    StringBuilder sb = new StringBuilder();
    if (n < docCache.textSize(WtDocumentCache.CURSOR_TYPE_TEXT)) {
      sb.append(docCache.getTextParagraph(new TextParagraph(WtDocumentCache.CURSOR_TYPE_TEXT, n))).append('\n');
      n++;
    }
    for (; n < docCache.textSize(WtDocumentCache.CURSOR_TYPE_TEXT) && !isRealHeader(n); n++) {
      sb.append(docCache.getTextParagraph(new TextParagraph(WtDocumentCache.CURSOR_TYPE_TEXT, n))).append('\n');
    }
    String text = sb.toString();
    if (debugMode) {
      WtMessageHandler.printToLogFile("AiTextToSpeech: runInstruction: text: " + text);
    }
    aiRemote.runTtsInstruction(text, filename);
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      n = docCache.textSize(WtDocumentCache.CURSOR_TYPE_TEXT);
    }
    return n;
  }

}

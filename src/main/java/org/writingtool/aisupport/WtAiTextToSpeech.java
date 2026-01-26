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

import java.io.File;
import java.net.URI;
import java.net.URISyntaxException;
import java.util.Map;
import java.util.ResourceBundle;
import java.util.regex.Pattern;

import javax.swing.JFileChooser;
import javax.swing.SwingUtilities;

import org.writingtool.WtDocumentCache;
import org.writingtool.WtDocumentsHandler;
import org.writingtool.WtSingleDocument;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.WtDocumentsHandler.WaitDialogThread;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;

import com.sun.star.frame.XModel;
import com.sun.star.uno.UnoRuntime;

/**
 * Class for a new document filled by AI
 * @since 1.0
 * @author Fred Kruse
 */

public class WtAiTextToSpeech extends Thread {
  
  private static final ResourceBundle messages = WtOfficeTools.getMessageBundle();
  private final static String ParaEndSign = ".";   //  Sign to force speech to make a small pause
  private final static String PauseSign = " - . - . - .";

  public final static String WAIT_TITLE = messages.getString("loAiWaitDialogTitle");
  public final static String WAIT_MESSAGE = messages.getString("loAiWaitDialogMessage");
  
  public final static String TMP_DIRNAME = "audioOut";
  public final static String FILE_EXTENSION = ".wav";

  private int debugMode = WtOfficeTools.DEBUG_MODE_AI;   //  should be false except for testing
  
  private final Pattern ALPHA_NUM = Pattern.compile("[\\p{L}\\d]");
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
      String audioDir = getAudioDir();
      if (audioDir == null) {
        return;
      }
      waitDialog = WtDocumentsHandler.getWaitDialog(WAIT_TITLE, WAIT_MESSAGE);
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
      int nParaStart = 0;
      int nFile = 0;
      int maxPara = docCache.textSize(WtDocumentCache.CURSOR_TYPE_TEXT);
      SwingUtilities.invokeLater(() -> { waitDialog.initializeProgressBar(0, maxPara); });
      
      while ((nParaStart = createAudioFile(nParaStart, nFile, audioDir, aiRemote)) < maxPara) {
        nFile++;
        if (waitDialog.canceled()) {
          waitDialog.close();
          return;
        }
        int nValue = nParaStart;
        SwingUtilities.invokeLater(() -> { waitDialog.setValueForProgressBar(nValue, true); });
      }
      waitDialog.close();
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
  }
  
  private String getAudioDir() {
    String sAudioDir = null;
    XModel xModel = UnoRuntime.queryInterface(XModel.class, document.getXComponent());
    if (xModel != null) {
      String url = xModel.getURL();
      if (url != null && url.startsWith("file://")) {
        URI uri;
        try {
          uri = new URI(url);
          sAudioDir = uri.getPath();
          int nDir = sAudioDir.lastIndexOf(".");
          sAudioDir = sAudioDir.substring(0, nDir) + "_audio";
        } catch (URISyntaxException e) {
        }
      }
    }
    if (sAudioDir == null) {
      sAudioDir = System.getProperty("user.home");
      if (sAudioDir == null) {
        return null;
      }
      sAudioDir += "/" + TMP_DIRNAME;
    }
    File audioDir = new File(sAudioDir);
    if (audioDir.exists()) {
      String path = audioDir.getAbsolutePath();
      int i = 1;
      while ((audioDir = new File(path + "." + i)).exists()) {
        i++;
      };
    }
    JFileChooser fileChooser = new JFileChooser(audioDir.getParent());
    fileChooser.setSelectedFile(audioDir);
    int choose = fileChooser.showSaveDialog(null);
    if (choose != JFileChooser.APPROVE_OPTION) {
      return null;
    }
    audioDir = fileChooser.getSelectedFile();
    audioDir.mkdirs();
    return audioDir.getAbsolutePath() + "/";
  }
  
  /** is Header of chapter
   * NOTE: nPara is number of text paragraph
   */
  private boolean isRealHeader(int nPara) {
    return (headingMap.containsKey(nPara) && headingMap.get(nPara) > 0);
  }
  
  private boolean isPause(String sPara) {
    return !sPara.isBlank() && !ALPHA_NUM.matcher(sPara).find();
  }
  
  private StringBuilder addToText(int nPara, StringBuilder sb) {
    String sPara = docCache.getTextParagraph(new TextParagraph(WtDocumentCache.CURSOR_TYPE_TEXT, nPara));
    sb.append(sPara);
    if (!sPara.endsWith(ParaEndSign)) {
      sb.append(ParaEndSign);
    }
    sb.append(" ");
    if (isRealHeader(nPara) || isPause(sPara)) {
      sb.append(PauseSign);
    }
    return sb;
  }
  
  private int createAudioFile(int nParaStart, int nFile, String audioDir, WtAiRemote aiRemote) {
    int n = nParaStart;
    String filename;
    try {
    if (isRealHeader(nParaStart)) {
      filename = docCache.getTextParagraph(new TextParagraph(WtDocumentCache.CURSOR_TYPE_TEXT, nParaStart));
      filename = filename.replaceAll("[^\\w\\d]", "_");
    } else {
      filename = "Intro";
    }
    filename = audioDir + (nFile < 10 ? "0" : "") + nFile + "_" + filename + FILE_EXTENSION;
    StringBuilder sb = new StringBuilder();
    if (n < docCache.textSize(WtDocumentCache.CURSOR_TYPE_TEXT)) {
      addToText(n, sb);
      n++;
    }
    for (; n < docCache.textSize(WtDocumentCache.CURSOR_TYPE_TEXT) && !isRealHeader(n); n++) {
      addToText(n, sb);
    }
    String text = sb.toString();
    if (debugMode > 1) {
      WtMessageHandler.printToLogFile("AiTextToSpeech: runInstruction: text: " + text);
    }
    aiRemote.runTtsInstruction(text, filename, false);
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      n = docCache.textSize(WtDocumentCache.CURSOR_TYPE_TEXT);
    }
    return n;
  }

}

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

import java.util.ResourceBundle;

import org.writingtool.WtDocumentCache;
import org.writingtool.WtSingleDocument;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.WtDocumentsHandler.WaitDialogThread;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtDocumentCursorTools;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtViewCursorTools;

import com.sun.star.lang.XComponent;

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
  
  public final static String TMP_FILENAME = "tmpOut.wav";
  
  private WaitDialogThread waitDialog = null;
//  private XComponentContext xContext;
  private WtSingleDocument document;
  WtDocumentCursorTools docCursor;
  //  private XComponent xToComponent;
  
  public WtAiTextToSpeech(WtSingleDocument document, ResourceBundle messages) {
//-    this.xContext = xContext;
    this.document = document;
  }
  
  @Override
  public void run() {
    try {
      waitDialog = new WaitDialogThread(WAIT_TITLE, WAIT_MESSAGE);
      waitDialog.start();
      WtConfiguration config = document.getMultiDocumentsHandler().getConfiguration();
      XComponent xComponent = document.getXComponent();
      WtViewCursorTools viewCursor = new WtViewCursorTools(xComponent);
      TextParagraph tPara = viewCursor.getViewCursorParagraph();
      WtDocumentCache docCache = document.getDocumentCache();
      String text = docCache.getTextParagraph(tPara);
      
      String filename = WtOfficeTools.getWtConfigDir().getAbsolutePath();
      filename += (filename.endsWith("/") ? "" : "/") + TMP_FILENAME;

      WtAiRemote aiRemote = new WtAiRemote(document.getMultiDocumentsHandler(), config);
      if (debugMode) {
        WtMessageHandler.printToLogFile("AiTextToSpeech: runInstruction: text: " + text);
      }
      String output = aiRemote.runTtsInstruction(text, filename);
      waitDialog.close();
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
  }

}

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

import java.awt.Component;
import java.util.ResourceBundle;

import javax.swing.SwingUtilities;

import org.writingtool.WtDocumentCache;
import org.writingtool.WtSingleDocument;
import org.writingtool.dialogs.WtAiTranslationDialog;
import org.writingtool.dialogs.WtAiTranslationDialog.TranslationOptions;
import org.writingtool.dialogs.WtOptionPane;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.WtDocumentsHandler.WaitDialogThread;
import org.writingtool.tools.WtDocumentCursorTools;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;

import com.sun.star.awt.XWindow;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.frame.XModel;
import com.sun.star.frame.XStorable;
import com.sun.star.io.IOException;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.Locale;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XComponent;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.util.XModifiable;

/**
 * Class for a new document filled by AI
 * @since 1.0
 * @author Fred Kruse
 */

public class WtAiTranslateDocument extends Thread {
  
  public static String TRANSLATE_INSTRUCTION = "Print the translation of the following text in the language ";
  public static String TRANSLATE_INSTRUCTION_POST = " (without comments)";
  
  private final ResourceBundle messages;
  private WaitDialogThread waitDialog = null;
  private WtSingleDocument document;
  private WtDocumentCursorTools docCursor;
  private Locale locale;
  private float temperature;
  private String fromUrl;
  
  public WtAiTranslateDocument(WtSingleDocument document, ResourceBundle messages) {
    this.document = document;
    this.messages = messages;
    XModel xModel = UnoRuntime.queryInterface(XModel.class, document.getXComponent());
    if (xModel == null) {
      WtMessageHandler.printToLogFile("WtAiTranslateDocument: XModel not found!");
      return;
    }
    fromUrl = xModel.getURL();
  }
  
  @Override
  public void run() {
    try {
      if (fromUrl == null || fromUrl.isBlank()) {
        WtOptionPane.showMessageDialog (null, messages.getString("loAiDialogTranslateFileError"));
        return;
      }
      if (isDocModified()) {
        int ret = WtOptionPane.showConfirmDialog(null, messages.getString("loAiDialogSaveFileMessage"), 
            messages.getString("guiOOSaveDocument"), WtOptionPane.OK_CANCEL_OPTION);
        if (ret == WtOptionPane.CANCEL_OPTION) {
          return;
        }
        fileSave();
      }
      WtMessageHandler.printToLogFile("WtAiTranslateDocument: fromUrl: " + fromUrl);
      TranslationOptions transOpt = getTranslationOptions();
      if (transOpt != null) {
        locale = transOpt.locale;
        temperature = transOpt.temperature;
        WtMessageHandler.printToLogFile("Locale: " + WtOfficeTools.localeToString(locale));
        fileSaveAs();
        XWindow xWindow = WtOfficeTools.getCurrentWindow(document.getMultiDocumentsHandler().getContext());
        fileLoad();
//        XWindow xNewWindow = WtOfficeTools.getCurrentWindow(document.getMultiDocumentsHandler().getContext());
        xWindow.setFocus();
//        xNewWindow.setVisible(true);
        writeText();
        fileSaveAs();
      } else {
        WtMessageHandler.printToLogFile("Locale: null");
      }
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
  }
  private TranslationOptions getTranslationOptions() {
    WtAiTranslationDialog langDialog = new WtAiTranslationDialog(document, messages);
    return langDialog.run();
  }

  private void replaceParagraph(TextParagraph textPara, String str, Locale locale) 
       throws UnknownPropertyException, PropertyVetoException, IllegalArgumentException, WrappedTargetException {
    XParagraphCursor pCursor = docCursor.getParagraphCursor(textPara);
    if (pCursor == null) {
      docCursor = new WtDocumentCursorTools(document.getXComponent());
      pCursor = docCursor.getParagraphCursor(textPara);
    }
    pCursor.gotoStartOfParagraph(false);
    pCursor.gotoEndOfParagraph(true);
    pCursor.setString(str);
    XPropertySet xCursorProps = UnoRuntime.queryInterface(
        XPropertySet.class, pCursor );
    xCursorProps.setPropertyValue ( "CharLocale", locale);
  }
  
  private String outUrl(String inUrl, String lang) {
    int n = inUrl.lastIndexOf('.');
    if (n < 0) {
      inUrl += ".odt";
      n = inUrl.lastIndexOf('.');
    }
    String name = inUrl.substring(0, n);
    String ext = inUrl.substring(n);
    return name + "_" + lang + ext;
  }
  
  private void fileSaveAs() throws IOException {
    XStorable xStore = UnoRuntime.queryInterface (com.sun.star.frame.XStorable.class, document.getXComponent());
    PropertyValue[] sProps = new PropertyValue[1];
    sProps[0] = new PropertyValue();
    sProps[0].Name = "Overwrite";
    sProps[0].Value = true; 
    xStore.storeAsURL (outUrl(fromUrl, locale.Language), sProps);
  }
  
  private void fileSave() throws IOException {
    XStorable xStore = UnoRuntime.queryInterface (com.sun.star.frame.XStorable.class, document.getXComponent());
    PropertyValue[] sProps = new PropertyValue[1];
    sProps[0] = new PropertyValue();
    sProps[0].Name = "Overwrite";
    sProps[0].Value = true; 
    xStore.store();
  }
  
  private XComponent fileLoad() {
    try {
      XDesktop xDesktop = WtOfficeTools.getDesktop(document.getMultiDocumentsHandler().getContext());
      if (xDesktop == null) {
        WtMessageHandler.printToLogFile("WtAiTranslateDocument: fileLoad: xDesktop == null");   
        return null;
      }
      XComponentLoader xLoader = UnoRuntime.queryInterface (XComponentLoader.class, xDesktop);
/*      
      PropertyValue[] sProps = new PropertyValue[1];
      sProps[0] = new PropertyValue();
      sProps[0].Name = "Hidden";
      sProps[0].Value = true; 
*/      
      PropertyValue[] sProps = new PropertyValue[0];
      WtMessageHandler.printToLogFile("WtAiTranslateDocument: fileLoad: Load file: " + fromUrl.toString());
      return xLoader.loadComponentFromURL(fromUrl, "_blank", 0, sProps);
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
    return null;
  }
  
  private boolean isDocModified() {
    XModifiable xModifiable = UnoRuntime.queryInterface (XModifiable.class, document.getXComponent());
    return xModifiable.isModified();
  }
  
  private void writeText() {
    try {
      WtDocumentCache fromCache = document.getDocumentCache();
      waitDialog = new WaitDialogThread(WtAiParagraphChanging.WAIT_TITLE, WtAiParagraphChanging.WAIT_MESSAGE);
      waitDialog.start();
      WtAiRemote aiRemote = new WtAiRemote(document.getMultiDocumentsHandler(), document.getMultiDocumentsHandler().getConfiguration());
      docCursor = document.getDocumentCursorTools();
      String instruction = TRANSLATE_INSTRUCTION + locale.Language + TRANSLATE_INSTRUCTION_POST;
      SwingUtilities.invokeLater(() -> { waitDialog.initializeProgressBar(0, fromCache.size()); });
      for(int i = 0; i < fromCache.size(); i++) {
        if (waitDialog.canceled()) {
          break;
        }
        String str = fromCache.getFlatParagraph(i);
        if(str.trim().isEmpty()) {
          continue;
        }
        String out = aiRemote.runInstruction(instruction, str, temperature, 1, locale, true);
        TextParagraph textPara = fromCache.getNumberOfTextParagraph(i);
        replaceParagraph(textPara, out, locale);
        int nValue = i;
        SwingUtilities.invokeLater(() -> { waitDialog.setValueForProgressBar(nValue, true); });
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
    if (waitDialog != null) {
      waitDialog.close();
    }

  }

}

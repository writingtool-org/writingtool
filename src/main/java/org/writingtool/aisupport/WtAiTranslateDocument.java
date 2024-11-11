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
import org.writingtool.dialogs.WtAiLanguageDialog;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.tools.WtDocumentCursorTools;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.frame.XModel;
import com.sun.star.frame.XStorable;
import com.sun.star.io.IOException;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.Locale;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 * Class for a new document filled by AI
 * @since WT 1.0
 * @author Fred Kruse
 */

public class WtAiTranslateDocument extends Thread {
  
  private static final ResourceBundle messages = WtOfficeTools.getMessageBundle();

  private String TRANSLATE_INSTRUCTION = "Print the translation of the following text in ";
  private String TRANSLATE_INSTRUCTION_POST = " (without comments)";
  
//  private XComponentContext xContext;
  private WtSingleDocument document;
//  private XComponent xToComponent;
  private Locale locale;
  private String fromUrl;
  
  public WtAiTranslateDocument(XComponentContext xContext, WtSingleDocument document) {
//-    this.xContext = xContext;
    this.document = document;
    XModel xModel = UnoRuntime.queryInterface(XModel.class, document.getXComponent());
    if (xModel == null) {
      WtMessageHandler.printToLogFile("CacheIO: getDocumentPath: XModel not found!");
      return;
    }
    fromUrl = xModel.getURL();

  }
  
  @Override
  public void run() {
    try {
//      xToComponent = openNewFile();
      locale = getLocale();
      if (locale != null) {
        saveFile(locale);
        writeText(locale);
      }
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
  }
  
  private Locale getLocale() {
    WtAiLanguageDialog langDialog = new WtAiLanguageDialog(document, messages);
    return langDialog.run();
  }
/*  
  private XComponent openNewFile() throws IOException, IllegalArgumentException {
    XDesktop xDesktop = WtOfficeTools.getDesktop(xContext);
    XComponentLoader xComponentLoader = (XComponentLoader)UnoRuntime.queryInterface(
        XComponentLoader.class, xDesktop);
    String loadURL = "private:factory/swriter";
    
    // the boolean property Hidden tells the office to open a file in hidden mode
    PropertyValue[] loadProps = new PropertyValue[1];
    loadProps[0] = new PropertyValue();
    loadProps[0].Name = "Hidden";
    loadProps[0].Value = false; 
    return xComponentLoader.loadComponentFromURL(loadURL, "_blank", 0, loadProps);
  }

  private void insertParagraph(String str, XText docText) {
    docText.getEnd().setString(str);
    docText.getEnd().setString("\n");
  }
*//*  
  private void replaceParagraph(int nFPara, int oldLength, String str, Locale locale) {
    document.getFlatParagraphTools().changeTextAndLocaleOfParagraph(nFPara, 0, oldLength, str, locale);
  }
*/
  private void replaceParagraph(TextParagraph textPara, String str, Locale locale, 
      WtDocumentCursorTools docCursor) throws UnknownPropertyException, PropertyVetoException, IllegalArgumentException, WrappedTargetException {
    XParagraphCursor pCursor = docCursor.getParagraphCursor(textPara);
    pCursor.gotoStartOfParagraph(false);
    pCursor.gotoEndOfParagraph(true);
    pCursor.setString(str);
    XPropertySet xCursorProps = UnoRuntime.queryInterface(
        XPropertySet.class, pCursor );
    xCursorProps.setPropertyValue ( "CharLocale", locale);
  }
  
  private String outUrl(String inUrl, String lang) {
    int n = inUrl.lastIndexOf('.');
    String name = inUrl.substring(0, n);
    String ext = inUrl.substring(n);
    return name + "_" + lang + ext;
  }
  
  private void saveFile(Locale locale) throws IOException {
    XStorable xStore = UnoRuntime.queryInterface (com.sun.star.frame.XStorable.class, document.getXComponent());
    PropertyValue[] sProps = new PropertyValue[1];
    sProps[0] = new PropertyValue();
    sProps[0].Name = "Overwrite";
    sProps[0].Value = true; 
    xStore.storeAsURL (outUrl(fromUrl, locale.Language), sProps);
  }
  
  private void writeText(Locale locale) {
    try {
      WtAiRemote aiRemote = new WtAiRemote(document.getMultiDocumentsHandler(), document.getMultiDocumentsHandler().getConfiguration());
      WtDocumentCache fromCache = document.getDocumentCache();
      WtDocumentCursorTools docCursor = document.getDocumentCursorTools();
      for(int i = 0; i < fromCache.size(); i++) {
        String str = fromCache.getFlatParagraph(i);
        String instruction = TRANSLATE_INSTRUCTION + locale.Language + TRANSLATE_INSTRUCTION_POST;
        String out = aiRemote.runInstruction(instruction, str, 0, 1, locale, true);
        TextParagraph textPara = fromCache.getNumberOfTextParagraph(i);
        replaceParagraph(textPara, out, locale, docCursor);
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
  }

}

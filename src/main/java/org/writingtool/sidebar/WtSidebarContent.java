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
package org.writingtool.sidebar;

import java.awt.SystemColor;
import java.util.ArrayList;
import java.util.List;
import java.util.ResourceBundle;
import java.util.SortedMap;
import java.util.TreeMap;

import javax.swing.SwingUtilities;

import com.sun.star.accessibility.XAccessible;
import com.sun.star.awt.ActionEvent;
import com.sun.star.awt.ImagePosition;
import com.sun.star.awt.Rectangle;
import com.sun.star.awt.WindowEvent;
import com.sun.star.awt.XButton;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XFixedText;
import com.sun.star.awt.XTextComponent;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.awt.tab.XTabPageContainer;
import com.sun.star.awt.tab.XTabPageContainerModel;
import com.sun.star.lang.DisposedException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lib.uno.helper.ComponentBase;
import com.sun.star.text.XTextDocument;
import com.sun.star.ui.LayoutSize;
import com.sun.star.ui.XSidebarPanel;
import com.sun.star.ui.XToolPanel;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import org.libreoffice.ext.unohelper.common.UNO;
import org.libreoffice.ext.unohelper.common.UnoHelperException;
import org.libreoffice.ext.unohelper.dialog.adapter.AbstractActionListener;
import org.libreoffice.ext.unohelper.dialog.adapter.AbstractWindowListener;
import org.libreoffice.ext.unohelper.ui.GuiFactory;
import org.libreoffice.ext.unohelper.ui.layout.ControlLayout;
import org.libreoffice.ext.unohelper.ui.layout.HorizontalLayout;
import org.libreoffice.ext.unohelper.ui.layout.Layout;
import org.libreoffice.ext.unohelper.ui.layout.VerticalLayout;
import org.libreoffice.ext.unohelper.util.UnoProperty;
import org.writingtool.WritingTool;
import org.writingtool.WtDocumentCache;
import org.writingtool.WtDocumentCache.TextParagraph;
import org.writingtool.WtDocumentsHandler;
import org.writingtool.WtLanguageTool;
import org.writingtool.WtProofreadingError;
import org.writingtool.WtResultCache;
import org.writingtool.WtSingleDocument;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtViewCursorTools;

/**
 * Create the window for the WT sidebar panel
 * 
 * @since 1.3
 * @author Fred Kruse
 */
public class WtSidebarContent extends ComponentBase implements XToolPanel, XSidebarPanel {

  private static final ResourceBundle messages = WtOfficeTools.getMessageBundle();
  
  private static String OpeningFormatSign = "[";
  private static String ClosingFormatSign = "]";

  private XComponentContext xContext;           //  the component context
  private XMultiComponentFactory xMCF;          //  The component factory
  private XControlContainer controlContainer;   //  The container of the controls
  private Layout layout;                        //  The layout of the controls
  private Layout buttonContainerlayout;         //  The layout of the buttons
  private XTextComponent paragraphBox;          //  Box to show text

  public WtSidebarContent(XComponentContext context, XWindow parentWindow) {
    xContext = context;
    AbstractWindowListener windowAdapter = new AbstractWindowListener() {
      @Override
      public void windowResized(WindowEvent e) {
        layout.layout(parentWindow.getPosSize());
      }
    };
    parentWindow.addWindowListener(windowAdapter);
    layout = new VerticalLayout(5, 5, 5, 5, 5);

    xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class, context.getServiceManager());
    XWindowPeer parentWindowPeer = UNO.XWindowPeer(parentWindow);

    if (parentWindowPeer == null) {
        return;
    }

    XToolkit parentToolkit = parentWindowPeer.getToolkit();
    controlContainer = UNO
            .XControlContainer(GuiFactory.createControlContainer(xMCF, context, new Rectangle(0, 0, 0, 0), null));
    XControl xContainer = UNO.XControl(controlContainer);
    xContainer.createPeer(parentToolkit, parentWindowPeer);
/*    
    XControl tabControl = GuiFactory.createTabPageContainer(xMCF, context);
    XTabPageContainer tabContainer = UNO.XTabPageContainer(tabControl);
    XTabPageContainerModel tabModel = UNO.XTabPageContainerModel(tabControl.getModel());
    GuiFactory.createTab(xMCF, context, tabModel, "Test1", (short) 1, 100);
    controlContainer.addControl("tabs", UNO.XControl(tabContainer));
    layout.addControl(UNO.XControl(tabContainer));
*/
    // Add button container
    XControl xButtonContainer = GuiFactory.createControlContainer(xMCF, context, new Rectangle(0, 0, 80, 40), null);
    XControlContainer buttonContainer = UNO.XControlContainer(xButtonContainer);
    XWindow containerWindow = UNO.XWindow(xContainer);
    buttonContainerlayout = new HorizontalLayout(5, 5, 5, 10, 5);
    AbstractWindowListener buttonContainerAdapter = new AbstractWindowListener() {
      @Override
      public void windowResized(WindowEvent e) {
        buttonContainerlayout.layout(containerWindow.getPosSize());
      }
    };
    containerWindow.addWindowListener(buttonContainerAdapter);
    controlContainer.addControl("buttonContainer", xButtonContainer);
    layout.addControl(xButtonContainer);
    
    // Add buttons to container
    
    addButtonToContainer("nextError", "WTNextSmall.png", "loContextMenuNextError", buttonContainer, buttonContainerlayout);
    
    addButtonToContainer("checkDialog", "WTCheckSmall.png", "loContextMenuGrammarCheck", buttonContainer, buttonContainerlayout);
    
    addButtonToContainer("checkAgainDialog", "WTCheckAgainSmall.png", "loContextMenuGrammarCheckAgain", buttonContainer, buttonContainerlayout);
    
    addButtonToContainer("refreshCheck", "WTRefreshSmall.png", "loContextMenuRefreshCheck", buttonContainer, buttonContainerlayout);
    
    addButtonToContainer("configure", "WTOptionsSmall.png", "loContextMenuOptions", buttonContainer, buttonContainerlayout);
    
    addButtonToContainer("about", "WTAboutSmall.png", "loContextMenuAbout", buttonContainer, buttonContainerlayout);
    
    // Add text field
    SortedMap<String, Object> props = new TreeMap<>();
    props.put("TextColor", SystemColor.textInactiveText.getRGB() & ~0xFF000000);
    props.put("Autocomplete", false);
    props.put("HideInactiveSelection", true);
    props.put("MultiLine", true);
    props.put("VScroll", true);
    props.put("HScroll", true);
//    XFixedText searchBox = UNO.XFixedText(
//        GuiFactory.createLabel(xMCF, context, WtOfficeTools.getFormatedTextVersionInformation(), new Rectangle(0, 0, 100, 32), props));
    paragraphBox = UNO.XTextComponent(
            GuiFactory.createTextfield(xMCF, context, "", new Rectangle(0, 0, 100, 32), props, null));
    controlContainer.addControl("paragraphBox", UNO.XControl(paragraphBox));
    layout.addControl(UNO.XControl(paragraphBox));
    setTestToBox();
/*
    // Add button
    XControl button = GuiFactory.createButton(xMCF, context, "Insert", null, new Rectangle(0, 0, 100, 32), null);
    XButton xbutton = UNO.XButton(button);
    AbstractActionListener xButtonAction = event -> {
      String text = paragraphBox.getText();
      try {
        UNO.init(xMCF);
        XTextDocument doc = UNO.getCurrentTextDocument();
        doc.getText().setString(text);
      } catch (UnoHelperException e1) {
        WtMessageHandler.printException(e1);
      }
    };
    xbutton.addActionListener(xButtonAction);
    controlContainer.addControl("button0", button);
    layout.addControl(button);
*/
  }
  
  private void addButtonToContainer(String cmd, String imageFile, String helpText, 
      XControlContainer buttonContainer, Layout containerLayout) {
    SortedMap<String, Object> bProps = new TreeMap<>();
    bProps.put(UnoProperty.IMAGE_URL, "vnd.sun.star.extension://org.openoffice.writingtool.oxt/images/" + imageFile);
    bProps.put(UnoProperty.HELP_TEXT, messages.getString(helpText));
    bProps.put("ImagePosition", ImagePosition.Centered);
    XControl button = GuiFactory.createButton(xMCF, xContext, "", null, new Rectangle(0, 0, 16, 16), bProps);
    XButton xButton = UNO.XButton(button);
    AbstractActionListener xButtonOptionsAction = event -> {
      try {
        WtOfficeTools.dispatchCmd("service:org.writingtool.WritingTool?" + cmd, xContext);
        setTestToBox();
      } catch (Throwable e1) {
        WtMessageHandler.showError(e1);
      }
    };
    xButton.addActionListener(xButtonOptionsAction);
    buttonContainer.addControl(cmd, button);
    containerLayout.addControl(button);
  }
  
  private WtProofreadingError[] getErrorsOfParagraph(TextParagraph tPara) {
    WtDocumentsHandler documents = WritingTool.getDocumentsHandler();
    WtLanguageTool lt = documents.getLanguageTool();
    WtSingleDocument document = documents.getCurrentDocument();
    if (document == null) {
      return new WtProofreadingError[0];
    }
    WtDocumentCache docCache = document.getDocumentCache();
    if (docCache == null) {
      return new WtProofreadingError[0];
    }
    int nFPara = docCache.getFlatParagraphNumber(tPara);
    List<WtProofreadingError[]> errors = new ArrayList<>();
    for (int cacheNum = 0; cacheNum < document.getParagraphsCache().size(); cacheNum++) {
      if (!docCache.isAutomaticGenerated(nFPara, true) && ((cacheNum == WtOfficeTools.CACHE_SINGLE_PARAGRAPH 
              || (lt.isSortedRuleForIndex(cacheNum) && !document.getDocumentCache().isSingleParagraph(nFPara)))
          || (cacheNum == WtOfficeTools.CACHE_AI && documents.useAi()))) {
        WtProofreadingError[] pErrors = document.getParagraphsCache().get(cacheNum).getSafeMatches(nFPara);
        //  Note: unsafe matches are needed to prevent the thread to get into a read lock
        errors.add(pErrors);
      } else {
        errors.add(new WtProofreadingError[0]);
      }
    }
    return document.mergeErrors(errors, nFPara);
  }
  
  private String formatText(String orgText, WtProofreadingError[] errors) {
    if (errors == null || errors.length == 0) {
      return orgText;
    }
    String formatedText = "";
    int lastChar = 0;
    for (WtProofreadingError error : errors) {
      if (lastChar <= error.nErrorStart) {
        if (error.nErrorStart > 0) {
          formatedText += orgText.substring(lastChar, error.nErrorStart);
        }
        lastChar = error.nErrorStart + error.nErrorLength;
        formatedText += OpeningFormatSign + orgText.substring(error.nErrorStart, lastChar) + ClosingFormatSign;
      }
    }
    if (lastChar < orgText.length()) {
      formatedText += orgText.substring(lastChar);
    }
    return formatedText;
  }
  
  private void setTestToBox() {
    try {
      XComponent xComponent = WtOfficeTools.getCurrentComponent(xContext);
      WtViewCursorTools vCursor = new WtViewCursorTools(xComponent);
      String orgText = vCursor.getViewCursorParagraphText();
      TextParagraph tPara = vCursor.getViewCursorParagraph();
      WtProofreadingError[] errors = getErrorsOfParagraph(tPara);
      String formatedText = formatText(orgText, errors);
      if (formatedText == null) {
        formatedText = "";
      }
      paragraphBox.setText(formatedText);
    } catch (Throwable e1) {
      WtMessageHandler.showError(e1);
    }
  }

  @Override
  public XAccessible createAccessible(XAccessible a) {
    return UNO.XAccessible(getWindow());
  }

  @Override
  public XWindow getWindow() {
    if (controlContainer == null) {
      throw new DisposedException("", this);
    }
    return UNO.XWindow(controlContainer);
  }

  @Override
  public LayoutSize getHeightForWidth(int width) {
    int height = layout.getHeightForWidth(width);
    return new LayoutSize(height, height, height);
  }

  @Override
  public int getMinimalWidth() {
    return 300;
  }
  
  
}

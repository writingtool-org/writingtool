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

import com.sun.star.accessibility.XAccessible;
import com.sun.star.awt.ImagePosition;
import com.sun.star.awt.PosSize;
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
import com.sun.star.lang.DisposedException;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lib.uno.helper.ComponentBase;
import com.sun.star.ui.LayoutSize;
import com.sun.star.ui.XSidebarPanel;
import com.sun.star.ui.XToolPanel;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import org.libreoffice.ext.unohelper.common.UNO;
import org.libreoffice.ext.unohelper.dialog.adapter.AbstractActionListener;
import org.libreoffice.ext.unohelper.dialog.adapter.AbstractWindowListener;
import org.libreoffice.ext.unohelper.ui.GuiFactory;
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
  
  private final static String OPENING_FORMAT_SIGN = "[";
  private final static String CLOSING_FORMAT_SIGN = "]";
  private final static String LINE_BREAK = "\n";
  
  private final static int LINE_MAX_CHAR = 40;
  
  private final static int MINIMAL_WIDTH = 300;

  private final static int BUTTON_MARGIN_TOP = 5;
  private final static int BUTTON_MARGIN_BOTTOM = 5;
  private final static int BUTTON_MARGIN_LEFT = 5;
  private final static int BUTTON_MARGIN_RIGHT = 5;
  private final static int BUTTON_MARGIN_BETWEEN = 5;

  private final static int BUTTON_WIDTH = 32;
  private final static int BUTTON_CONTAINER_WIDTH = MINIMAL_WIDTH - BUTTON_MARGIN_LEFT - BUTTON_MARGIN_RIGHT;
  private final static int BUTTON_CONTAINER_HEIGHT = BUTTON_WIDTH + BUTTON_MARGIN_LEFT + BUTTON_MARGIN_RIGHT;

  private final static int CONTAINER_MARGIN_TOP = 5;
  private final static int CONTAINER_MARGIN_BOTTOM = 5;
  private final static int CONTAINER_MARGIN_LEFT = 5;
  private final static int CONTAINER_MARGIN_RIGHT = 5;
  private final static int CONTAINER_MARGIN_BETWEEN = 5;


  private XComponentContext xContext;           //  the component context
  
  private XWindow parentWindow;                 //  the parent window
  private XWindow contentWindow;                //  the window of the control container
  private XWindow buttonContainerWindow;        //  the window of the button container
  private XWindow paragraphBoxWindow;           //  the window of the paragraph box container
  private XMultiComponentFactory xMCF;          //  The component factory
  private XControlContainer controlContainer;   //  The container of the controls
//  private Layout layout;                        //  The layout of the controls
  private Layout buttonContainerlayout;         //  The layout of the buttons
  private XFixedText paragraphBox;              //  Box to show text

  public WtSidebarContent(XComponentContext context, XWindow parentWindow) {
    xContext = context;
    this.parentWindow = parentWindow;
    try {
      AbstractWindowListener windowAdapter = new AbstractWindowListener() {
        @Override
        public void windowResized(WindowEvent e) {
          Rectangle rect = parentWindow.getPosSize();
//          layout.layout(rect);
          contentWindow.setPosSize(rect.X, rect.Y, rect.Width, rect.Height, PosSize.POSSIZE);
          buttonContainerWindow.setPosSize(CONTAINER_MARGIN_LEFT, CONTAINER_MARGIN_TOP, 
              BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT, PosSize.POSSIZE);
          int paraBoxX = CONTAINER_MARGIN_LEFT;
          int paraBoxY = CONTAINER_MARGIN_TOP + BUTTON_CONTAINER_HEIGHT + CONTAINER_MARGIN_BETWEEN;
          int paraBoxWidth = rect.Width - CONTAINER_MARGIN_LEFT - CONTAINER_MARGIN_RIGHT;
          int paraBoxHeight = rect.Height - CONTAINER_MARGIN_TOP - CONTAINER_MARGIN_BOTTOM;
          paragraphBoxWindow.setPosSize(paraBoxX, paraBoxY, paraBoxWidth, paraBoxHeight, PosSize.POSSIZE);
        }
      };
      parentWindow.addWindowListener(windowAdapter);
//      layout = new VerticalLayout(5, 5, 5, 5, 5);
  
      xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class, context.getServiceManager());
      XWindowPeer parentWindowPeer = UNO.XWindowPeer(parentWindow);
  
      if (parentWindowPeer == null) {
          return;
      }
  
      XToolkit parentToolkit = parentWindowPeer.getToolkit();
//      Rectangle parentWindowSize = parentWindow.getPosSize();
//      controlContainer = UNO
//          .XControlContainer(GuiFactory.createControlContainer(xMCF, context, parentWindowSize, null));
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
      XControl xButtonContainer = GuiFactory.createControlContainer(xMCF, context, 
          new Rectangle(CONTAINER_MARGIN_LEFT, CONTAINER_MARGIN_TOP, BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT), null);
      XControlContainer buttonContainer = UNO.XControlContainer(xButtonContainer);
      contentWindow = UNO.XWindow(xContainer);
      buttonContainerlayout = new HorizontalLayout(5, 5, 5, 10, 5);
      AbstractWindowListener buttonContainerAdapter = new AbstractWindowListener() {
        @Override
        public void windowResized(WindowEvent e) {
          buttonContainerlayout.layout(contentWindow.getPosSize());
        }
      };
      contentWindow.addWindowListener(buttonContainerAdapter);
      controlContainer.addControl("buttonContainer", xButtonContainer);
      buttonContainerWindow = UNO.XWindow(xButtonContainer);
      buttonContainerWindow.setPosSize(CONTAINER_MARGIN_LEFT, CONTAINER_MARGIN_TOP, BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT, PosSize.POSSIZE);
//      layout.addControl(xButtonContainer);
      
      // Add buttons to container
      int num = 0;
      addButtonToContainer(num, "nextError", "WTNextSmall.png", "loContextMenuNextError", buttonContainer, buttonContainerlayout);
      num++;
      addButtonToContainer(num, "checkDialog", "WTCheckSmall.png", "loContextMenuGrammarCheck", buttonContainer, buttonContainerlayout);
      num++;
      addButtonToContainer(num, "checkAgainDialog", "WTCheckAgainSmall.png", "loContextMenuGrammarCheckAgain", buttonContainer, buttonContainerlayout);
      num++;
      addButtonToContainer(num, "refreshCheck", "WTRefreshSmall.png", "loContextMenuRefreshCheck", buttonContainer, buttonContainerlayout);
      num++;
      addButtonToContainer(num, "configure", "WTOptionsSmall.png", "loContextMenuOptions", buttonContainer, buttonContainerlayout);
      num++;
      addButtonToContainer(num, "about", "WTAboutSmall.png", "loContextMenuAbout", buttonContainer, buttonContainerlayout);

      // Add text field
      SortedMap<String, Object> props = new TreeMap<>();
      props.put("TextColor", SystemColor.textInactiveText.getRGB() & ~0xFF000000);
      props.put("Autocomplete", false);
      props.put("HideInactiveSelection", true);
      props.put("MultiLine", true);
      props.put("VScroll", true);
      props.put("HScroll", true);
      Rectangle containerSize = contentWindow.getPosSize();
      int paraBoxX = CONTAINER_MARGIN_LEFT;
      int paraBoxY = CONTAINER_MARGIN_TOP + BUTTON_CONTAINER_HEIGHT + CONTAINER_MARGIN_BETWEEN;
      int paraBoxWidth = containerSize.Width - CONTAINER_MARGIN_LEFT - CONTAINER_MARGIN_RIGHT;
      int paraBoxHeight = containerSize.Height - CONTAINER_MARGIN_TOP - CONTAINER_MARGIN_BOTTOM;
      
      XControl xParagraphBox = GuiFactory.createLabel(xMCF, context, WtOfficeTools.getFormatedTextVersionInformation(), 
          new Rectangle(paraBoxX, paraBoxY, paraBoxWidth, paraBoxHeight), props); 
      paragraphBox = UNO.XFixedText(xParagraphBox);
  //    paragraphBox = UNO.XTextComponent(
  //            GuiFactory.createTextfield(xMCF, context, "", new Rectangle(0, 0, 80, 32), props, null));
      controlContainer.addControl("paragraphBox", UNO.XControl(paragraphBox));
  
//      layout.addControl(UNO.XControl(paragraphBox));
      paragraphBoxWindow = UNO.XWindow(xParagraphBox);
      paragraphBoxWindow.setPosSize(paraBoxX, paraBoxY, paraBoxWidth, paraBoxHeight, PosSize.POSSIZE);
      setTextToBox();
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
    } catch (Throwable t) {
      WtMessageHandler.printException(t);
    }
      
  }
  
  private void addButtonToContainer(int num, String cmd, String imageFile, String helpText, 
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
        setTextToBox();
      } catch (Throwable e1) {
        WtMessageHandler.showError(e1);
      }
    };
    xButton.addActionListener(xButtonOptionsAction);
    buttonContainer.addControl(cmd, button);
    XWindow xWindow = UNO.XWindow(button);
    xWindow.setPosSize(BUTTON_MARGIN_LEFT + num * (BUTTON_WIDTH + BUTTON_MARGIN_BETWEEN), 
        BUTTON_MARGIN_TOP, BUTTON_WIDTH, BUTTON_WIDTH, PosSize.POSSIZE);
//    containerLayout.addControl(button);
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
        formatedText += OPENING_FORMAT_SIGN + orgText.substring(error.nErrorStart, lastChar) + CLOSING_FORMAT_SIGN;
      }
    }
    if (lastChar < orgText.length()) {
      formatedText += orgText.substring(lastChar);
    }
    return formatedText;
  }
  
  private void setTextToBox() {
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
      paragraphBox.setText(addLineBreaks(formatedText));
//      layout.layout(parentWindow.getPosSize());
    } catch (Throwable e1) {
      WtMessageHandler.showError(e1);
    }
  }
  
  private String addLineBreaks(String inpText) {
    if (inpText.length() <= LINE_MAX_CHAR) {
      return inpText;
    }
    String outpText = "";
    int lastChar = 0;
    while (lastChar < inpText.length()) {
      if (lastChar + LINE_MAX_CHAR >= inpText.length()) {
        outpText += inpText.substring(lastChar);
        lastChar = inpText.length();
      } else {
        int i;
        for (i = lastChar + LINE_MAX_CHAR; i > lastChar && !Character.isWhitespace(inpText.charAt(i)) && inpText.charAt(i) != '-'; i--);
//        WtMessageHandler.printToLogFile("lastChar = " + ", i = " + i);
        if (i == 0) {
          outpText += inpText.substring(lastChar, lastChar + LINE_MAX_CHAR) + LINE_BREAK;
          lastChar += LINE_MAX_CHAR;
        } else {
          outpText += inpText.substring(lastChar, i + 1) + LINE_BREAK;
          lastChar = i + 1;
        }
      }
    }
    return outpText;
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
//    int height = layout.getHeightForWidth(width);
//    return new LayoutSize(height, height, height);
    int max = parentWindow.getPosSize().Height;
    return new LayoutSize(300, max, max);
  }

  @Override
  public int getMinimalWidth() {
    return MINIMAL_WIDTH;
  }
  
  
}

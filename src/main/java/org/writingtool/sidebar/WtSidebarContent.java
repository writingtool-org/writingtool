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
import com.sun.star.awt.ActionEvent;
import com.sun.star.awt.ImagePosition;
import com.sun.star.awt.PosSize;
import com.sun.star.awt.Rectangle;
import com.sun.star.awt.WindowEvent;
import com.sun.star.awt.XActionListener;
import com.sun.star.awt.XButton;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XFixedText;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowListener;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.lang.DisposedException;
import com.sun.star.lang.EventObject;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lib.uno.helper.ComponentBase;
import com.sun.star.ui.LayoutSize;
import com.sun.star.ui.XSidebarPanel;
import com.sun.star.ui.XToolPanel;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import org.languagetool.Language;
import org.libreoffice.ext.unohelper.common.UNO;
import org.libreoffice.ext.unohelper.ui.GuiFactory;
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
import org.writingtool.tools.WtOfficeTools.DocumentType;

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
  
  private final static int LINE_MAX_CHAR = 46;
  
  private final static int MINIMAL_WIDTH = 300;

  private final static int BUTTON_MARGIN_TOP = 5;
  private final static int BUTTON_MARGIN_BOTTOM = 5;
  private final static int BUTTON_MARGIN_LEFT = 5;
  private final static int BUTTON_MARGIN_RIGHT = 5;
  private final static int BUTTON_MARGIN_BETWEEN = 5;

  private final static int BUTTON_WIDTH = 32;
  private final static int BUTTON_CONTAINER_WIDTH = MINIMAL_WIDTH - BUTTON_MARGIN_LEFT - BUTTON_MARGIN_RIGHT;
  private final static int BUTTON_CONTAINER_HEIGHT = BUTTON_WIDTH + BUTTON_MARGIN_TOP + BUTTON_MARGIN_BOTTOM;

  private final static int LABEL_TOP = 2 * BUTTON_MARGIN_TOP + 2 * (BUTTON_CONTAINER_HEIGHT + BUTTON_MARGIN_BETWEEN) + BUTTON_MARGIN_BOTTOM;
  private final static int LABEL_LEFT = BUTTON_MARGIN_LEFT;
  private final static int LABEL_WIDTH = MINIMAL_WIDTH - BUTTON_MARGIN_LEFT - BUTTON_MARGIN_RIGHT;;
  private final static int LABEL_HEIGHT = 20;

  private final static int CONTAINER_MARGIN_TOP = 5;
  private final static int CONTAINER_MARGIN_BOTTOM = 5;
  private final static int CONTAINER_MARGIN_LEFT = 5;
  private final static int CONTAINER_MARGIN_RIGHT = 5;
  private final static int CONTAINER_MARGIN_BETWEEN = 5;
  private final static int CONTAINER_TOP = CONTAINER_MARGIN_TOP + LABEL_TOP + LABEL_HEIGHT;
  private final static int CONTAINER_HEIGHT = 200;
  


  private XComponentContext xContext;           //  the component context
  private WtDocumentsHandler documents;
  
  private XWindow parentWindow;                 //  the parent window
  private XWindow contentWindow;                //  the window of the control container
  private XWindow buttonContainer1Window;       //  the window of the first row button container
  private XWindow buttonContainer2Window;       //  the window of the second row button container
  private XWindow buttonContainerAiWindow;      //  the window of the button container for AI action
  private XWindow paragraphLabelWindow;         //  the window of the paragraph label
  private XWindow paragraphBoxWindow;           //  the window of the paragraph box
  private XWindow aiLabelWindow;                //  the window of the AI label
  private XWindow aiResultBoxWindow = null;     //  the window of the AI result box
  private XWindow buttonAutoOnWindow;           //  the window of button for auto check on
  private XWindow buttonAutoOffWindow;          //  the window of button for auto check off
  private XWindow buttonStatAnWindow;           //  the window of button for statistical Analysis
  private XWindow buttonAiGeneralWindow;        //  the window of button for AI general
  private XWindow buttonAiTranslateTextWindow;  //  the window of button for AI text to speech
  private XWindow buttonAiTextToSpeechWindow;   //  the window of button for AI text to speech
  private XMultiComponentFactory xMCF;          //  The component factory
  private XControlContainer controlContainer;   //  The container of the controls
  private XFixedText paragraphBox;              //  Box to show text

  public WtSidebarContent(XComponentContext context, XWindow parentWindow) {
    xContext = context;
    this.parentWindow = parentWindow;
    try {
      XWindowListener windowAdapter = new XWindowListener() {
        @Override
        public void windowResized(WindowEvent e) {
          resizeContainer();
        }
        @Override
        public void disposing(EventObject arg0) { }
        @Override
        public void windowHidden(EventObject arg0) { }
        @Override
        public void windowMoved(WindowEvent arg0) { }
        @Override
        public void windowShown(EventObject arg0) { }
      };
      
      parentWindow.addWindowListener(windowAdapter);
  
      xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class, context.getServiceManager());
      XWindowPeer parentWindowPeer = UNO.XWindowPeer(parentWindow);
  
      if (parentWindowPeer == null) {
          return;
      }
  
      XToolkit parentToolkit = parentWindowPeer.getToolkit();
      SortedMap<String, Object> props = new TreeMap<>();
      props.put("BackgroundColor", SystemColor.menu.getRGB() & ~0xFF000000);

      //  Get basic control container
      controlContainer = UNO
          .XControlContainer(GuiFactory.createControlContainer(xMCF, context, new Rectangle(0, 0, 0, 0), props));
      XControl xContainer = UNO.XControl(controlContainer);
      xContainer.createPeer(parentToolkit, parentWindowPeer);
      contentWindow = UNO.XWindow(xContainer);
      
      // Add first row button container
      XControl xButtonContainer = GuiFactory.createControlContainer(xMCF, context, 
          new Rectangle(CONTAINER_MARGIN_LEFT, CONTAINER_MARGIN_TOP, BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT), props);
      XControlContainer buttonContainer = UNO.XControlContainer(xButtonContainer);

      controlContainer.addControl("buttonContainer1", xButtonContainer);
      buttonContainer1Window = UNO.XWindow(xButtonContainer);
      buttonContainer1Window.setPosSize(CONTAINER_MARGIN_LEFT, CONTAINER_MARGIN_TOP, BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT, PosSize.POSSIZE);
      
      // Add buttons to container
      int num = 0;
      addButtonToContainer(num, "nextError", "WTNextSmall.png", "loContextMenuNextError", buttonContainer);
      num++;
      addButtonToContainer(num, "checkDialog", "WTCheckSmall.png", "loContextMenuGrammarCheck", buttonContainer);
      num++;
      addButtonToContainer(num, "checkAgainDialog", "WTCheckAgainSmall.png", "loContextMenuGrammarCheckAgain", buttonContainer);
      num++;
      addButtonToContainer(num, "refreshCheck", "WTRefreshSmall.png", "loContextMenuRefreshCheck", buttonContainer);
      num++;
      addButtonToContainer(num, "configure", "WTOptionsSmall.png", "loContextMenuOptions", buttonContainer);
      num++;
      addButtonToContainer(num, "about", "WTAboutSmall.png", "loContextMenuAbout", buttonContainer);
      num++;
      buttonAutoOnWindow = addButtonToContainer(num, "backgroundCheckOn", "WTBackCheckOnSmall.png", "loMenuEnableBackgroundCheck", buttonContainer);
      buttonAutoOffWindow = addButtonToContainer(num, "backgroundCheckOff", "WTBackCheckOffSmall.png", "loMenuDisableBackgroundCheck", buttonContainer);
      documents = WritingTool.getDocumentsHandler();
      buttonAutoOnWindow.setVisible(documents.isBackgroundCheckOff());
      buttonAutoOffWindow.setVisible(!documents.isBackgroundCheckOff());
      
      boolean isAiSupport = documents.getConfiguration().useAiSupport() || documents.getConfiguration().useAiTtsSupport();
//      boolean isAiSupport = true;

      // Add second row button container
      xButtonContainer = GuiFactory.createControlContainer(xMCF, context, 
          new Rectangle(CONTAINER_MARGIN_LEFT, CONTAINER_MARGIN_TOP + BUTTON_CONTAINER_HEIGHT + CONTAINER_MARGIN_BETWEEN, 
              BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT), props);
      buttonContainer = UNO.XControlContainer(xButtonContainer);
      controlContainer.addControl("buttonContainer2", xButtonContainer);
      buttonContainer2Window = UNO.XWindow(xButtonContainer);
      buttonContainer2Window.setPosSize(CONTAINER_MARGIN_LEFT, CONTAINER_MARGIN_TOP + BUTTON_CONTAINER_HEIGHT + CONTAINER_MARGIN_BETWEEN, 
          BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT, PosSize.POSSIZE);
      
      // Add buttons to container
      num = 0;
      buttonStatAnWindow = addButtonToContainer(num, "statisticalAnalyses", "WTStatAnSmall.png", "loStatisticalAnalysis", buttonContainer);
      buttonStatAnWindow.setEnable(hasStatAn());
      num++;
      buttonAiGeneralWindow = addButtonToContainer(num, "aiGeneralCommand", "WTAiGeneralSmall.png", "loMenuAiGeneralCommand", buttonContainer);
      buttonAiGeneralWindow.setEnable(isAiSupport);
      num++;
      buttonAiTranslateTextWindow = addButtonToContainer(num, "aiTranslateText", "WTAiTranslateTextSmall.png", "loMenuAiTranslateCommand", buttonContainer);
      buttonAiTranslateTextWindow.setEnable(isAiSupport);
      num++;
      buttonAiTextToSpeechWindow = addButtonToContainer(num, "aiTextToSpeech", "WTAiTextToSpeechSmall.png", "loMenuAiTextToSpeechCommand", buttonContainer);
      buttonAiTextToSpeechWindow.setEnable(documents.getConfiguration().useAiTtsSupport());
      
      // Add Label
      props.put("FontStyleName", "Bold");
      XControl xParagraphLabel = GuiFactory.createLabel(xMCF, context, messages.getString("loSidebarParagraphLabel") + ":", 
          new Rectangle(LABEL_LEFT, LABEL_TOP, LABEL_WIDTH, LABEL_HEIGHT), props); 
      controlContainer.addControl("paragraphBoxLabel", xParagraphLabel);
      paragraphLabelWindow = UNO.XWindow(xParagraphLabel);
      
      // Add text field
      props = new TreeMap<>();
      props.put("TextColor", SystemColor.textInactiveText.getRGB() & ~0xFF000000);
      props.put("BackgroundColor", SystemColor.control.getRGB() & ~0xFF000000);
      props.put("Autocomplete", false);
      props.put("HideInactiveSelection", true);
      props.put("MultiLine", true);
      props.put("VScroll", true);
      props.put("HScroll", true);
      Rectangle containerSize = contentWindow.getPosSize();
      int paraBoxX = CONTAINER_MARGIN_LEFT;
      int paraBoxY = CONTAINER_TOP;
      int paraBoxWidth = containerSize.Width - CONTAINER_MARGIN_LEFT - CONTAINER_MARGIN_RIGHT;
      int paraBoxHeight = isAiSupport ? CONTAINER_HEIGHT : containerSize.Height - CONTAINER_MARGIN_TOP - CONTAINER_MARGIN_BOTTOM;
      
      XControl xParagraphBox = GuiFactory.createLabel(xMCF, context, "", 
          new Rectangle(paraBoxX, paraBoxY, paraBoxWidth, paraBoxHeight), props); 
      paragraphBox = UNO.XFixedText(xParagraphBox);
  //    paragraphBox = UNO.XTextComponent(
  //            GuiFactory.createTextfield(xMCF, context, "", new Rectangle(0, 0, 80, 32), props, null));
      controlContainer.addControl("paragraphBox", xParagraphBox);
  
      paragraphBoxWindow = UNO.XWindow(xParagraphBox);
      paragraphBoxWindow.setPosSize(paraBoxX, paraBoxY, paraBoxWidth, paraBoxHeight, PosSize.POSSIZE);
      setTextToBox();
      
      if (isAiSupport) {
        // Add AI Button container
        props = new TreeMap<>();
        props.put("BackgroundColor", SystemColor.menu.getRGB() & ~0xFF000000);
        xButtonContainer = GuiFactory.createControlContainer(xMCF, context, 
            new Rectangle(CONTAINER_MARGIN_LEFT, CONTAINER_TOP + CONTAINER_HEIGHT + CONTAINER_MARGIN_BETWEEN, 
                BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT), props);
        buttonContainer = UNO.XControlContainer(xButtonContainer);
        controlContainer.addControl("buttonContainerAi", xButtonContainer);
        buttonContainerAiWindow = UNO.XWindow(xButtonContainer);
        buttonContainerAiWindow.setPosSize(CONTAINER_MARGIN_LEFT, CONTAINER_TOP + CONTAINER_HEIGHT + CONTAINER_MARGIN_BETWEEN, 
            BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT, PosSize.POSSIZE);

        // Add AI buttons
        num = 0;
        addButtonToContainer(num, "aiBetterStyle", "WTAiBetterStyleSmall.png", "loMenuAiStyleCommand", buttonContainer);
        num++;
        addButtonToContainer(num, "aiReformulateText", "WTAiReformulateSmall.png", "loMenuAiReformulateCommand", buttonContainer);
        num++;
        addButtonToContainer(num, "aiAdvanceText", "WTAiExpandSmall.png", "loMenuAiExpandCommand", buttonContainer);

        // Add AI Label
        props.put("FontStyleName", "Bold");
        int labelTop = CONTAINER_TOP + CONTAINER_HEIGHT + 2 * CONTAINER_MARGIN_BETWEEN + BUTTON_CONTAINER_HEIGHT;
        XControl xAiLabel = GuiFactory.createLabel(xMCF, context, messages.getString("loAiDialogResultLabel") + ":", 
            new Rectangle(LABEL_LEFT, labelTop, LABEL_WIDTH, LABEL_HEIGHT), props); 
        controlContainer.addControl("aiLabel", xAiLabel);
        aiLabelWindow = UNO.XWindow(xAiLabel);
        
        // Add text field
        props = new TreeMap<>();
        props.put("TextColor", SystemColor.textInactiveText.getRGB() & ~0xFF000000);
        props.put("BackgroundColor", SystemColor.control.getRGB() & ~0xFF000000);
        props.put("Autocomplete", false);
        props.put("HideInactiveSelection", true);
        props.put("MultiLine", true);
        props.put("VScroll", true);
        props.put("HScroll", true);
        int aiBoxX = CONTAINER_MARGIN_LEFT;
        int aiBoxY = CONTAINER_TOP + CONTAINER_HEIGHT + 2 * CONTAINER_MARGIN_BETWEEN + LABEL_HEIGHT + BUTTON_CONTAINER_HEIGHT;
        int aiBoxWidth = containerSize.Width - CONTAINER_MARGIN_LEFT - CONTAINER_MARGIN_RIGHT;
        int aiBoxHeight = CONTAINER_HEIGHT;
        XControl xAiBox = GuiFactory.createLabel(xMCF, context, "", 
            new Rectangle(aiBoxX, aiBoxY, aiBoxWidth, aiBoxHeight), props); 
        controlContainer.addControl("aiBox", xAiBox);
        aiResultBoxWindow = UNO.XWindow(xAiBox);
        aiResultBoxWindow.setPosSize(aiBoxX, aiBoxY, aiBoxWidth, aiBoxHeight, PosSize.POSSIZE);
        
      }
      
      
      
      
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
  
  private XWindow addButtonToContainer(int num, String cmd, String imageFile, String helpText, XControlContainer buttonContainer) {
    SortedMap<String, Object> bProps = new TreeMap<>();
    bProps.put(UnoProperty.IMAGE_URL, "vnd.sun.star.extension://org.openoffice.writingtool.oxt/images/" + imageFile);
    bProps.put(UnoProperty.HELP_TEXT, messages.getString(helpText));
    bProps.put("ImagePosition", ImagePosition.Centered);
    XControl button = GuiFactory.createButton(xMCF, xContext, "", null, new Rectangle(0, 0, 16, 16), bProps);
    XButton xButton = UNO.XButton(button);
    XActionListener xButtonOptionsAction = new XActionListener() {
      @Override
      public void actionPerformed(ActionEvent event) {
        try {
          WtSingleDocument document = documents.getCurrentDocument();
          if (document != null) {
            document.setMenuDocId();
          }
          WtOfficeTools.dispatchCmd("service:org.writingtool.WritingTool?" + cmd, xContext);
          setTextToBox();
        } catch (Throwable e1) {
          WtMessageHandler.showError(e1);
        }
      }
      @Override
      public void disposing(EventObject arg0) { }
    };
    xButton.addActionListener(xButtonOptionsAction);
    buttonContainer.addControl(cmd, button);
    XWindow xWindow = UNO.XWindow(button);
    xWindow.setPosSize(BUTTON_MARGIN_LEFT + num * (BUTTON_WIDTH + BUTTON_MARGIN_BETWEEN), 
        BUTTON_MARGIN_TOP, BUTTON_WIDTH, BUTTON_WIDTH, PosSize.POSSIZE);
    return xWindow;
  }

  private boolean hasStatAn() {
    WtSingleDocument document = documents.getCurrentDocument();
    if (document == null) {
      return false;
    }
    if (document.getDocumentType() != DocumentType.WRITER || documents.isBackgroundCheckOff()) {
      return false;
    }
    Language lang = document.getLanguage();
    if (lang == null) {
      return false;
    }
    return WtOfficeTools.hasStatisticalStyleRules(lang);
  }
  
  private void resizeContainer() {
    Rectangle rect = parentWindow.getPosSize();
    contentWindow.setPosSize(rect.X, rect.Y, rect.Width, rect.Height, PosSize.POSSIZE);
    buttonContainer1Window.setPosSize(CONTAINER_MARGIN_LEFT, CONTAINER_MARGIN_TOP, 
        BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT, PosSize.POSSIZE);
    buttonContainer2Window.setPosSize(CONTAINER_MARGIN_LEFT, CONTAINER_MARGIN_TOP + BUTTON_CONTAINER_HEIGHT + CONTAINER_MARGIN_BETWEEN, 
        BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT, PosSize.POSSIZE);
    paragraphLabelWindow.setPosSize(LABEL_LEFT, LABEL_TOP, LABEL_WIDTH, LABEL_HEIGHT, PosSize.POSSIZE);
    int paraBoxX = CONTAINER_MARGIN_LEFT;
    int paraBoxY = CONTAINER_TOP;
    int paraBoxWidth = rect.Width - CONTAINER_MARGIN_LEFT - CONTAINER_MARGIN_RIGHT;
    int paraBoxHeight = aiResultBoxWindow != null ? CONTAINER_HEIGHT : rect.Height - CONTAINER_MARGIN_TOP - CONTAINER_MARGIN_BOTTOM;
    paragraphBoxWindow.setPosSize(paraBoxX, paraBoxY, paraBoxWidth, paraBoxHeight, PosSize.POSSIZE);
    buttonAutoOnWindow.setVisible(documents.isBackgroundCheckOff());
    buttonAutoOffWindow.setVisible(!documents.isBackgroundCheckOff());
    buttonStatAnWindow.setEnable(hasStatAn());
    buttonAiGeneralWindow.setEnable(documents.getConfiguration().useAiSupport() || documents.getConfiguration().useAiTtsSupport());
    buttonAiTranslateTextWindow.setEnable(documents.getConfiguration().useAiSupport() || documents.getConfiguration().useAiTtsSupport());
    buttonAiTextToSpeechWindow.setEnable(documents.getConfiguration().useAiTtsSupport());
    if (aiResultBoxWindow != null) {
      buttonContainerAiWindow.setPosSize(CONTAINER_MARGIN_LEFT, CONTAINER_TOP + CONTAINER_HEIGHT + CONTAINER_MARGIN_BETWEEN, 
          BUTTON_CONTAINER_WIDTH, BUTTON_CONTAINER_HEIGHT, PosSize.POSSIZE);
      aiLabelWindow.setPosSize(LABEL_LEFT, CONTAINER_TOP + CONTAINER_HEIGHT + 2 * CONTAINER_MARGIN_BETWEEN + BUTTON_CONTAINER_HEIGHT, 
          LABEL_WIDTH, LABEL_HEIGHT, PosSize.POSSIZE);
      int aiBoxX = CONTAINER_MARGIN_LEFT;
      int aiBoxY = CONTAINER_TOP + CONTAINER_HEIGHT + 2 * CONTAINER_MARGIN_BETWEEN + LABEL_HEIGHT + BUTTON_CONTAINER_HEIGHT;
      int aiBoxWidth = rect.Width - CONTAINER_MARGIN_LEFT - CONTAINER_MARGIN_RIGHT;
      int aiBoxHeight = CONTAINER_HEIGHT;
      aiResultBoxWindow.setPosSize(aiBoxX, aiBoxY, aiBoxWidth, aiBoxHeight, PosSize.POSSIZE);
    }
  }
  
  private WtProofreadingError[] getErrorsOfParagraph(TextParagraph tPara) {
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
      buttonAutoOnWindow.setVisible(documents.isBackgroundCheckOff());
      buttonAutoOffWindow.setVisible(!documents.isBackgroundCheckOff());
      buttonStatAnWindow.setEnable(hasStatAn());
      buttonAiGeneralWindow.setEnable(documents.getConfiguration().useAiSupport() || documents.getConfiguration().useAiTtsSupport());
      resizeContainer();
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

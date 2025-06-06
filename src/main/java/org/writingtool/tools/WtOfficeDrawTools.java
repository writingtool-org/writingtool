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
package org.writingtool.tools;

import java.awt.Color;
import java.util.ArrayList;
import java.util.List;

import com.sun.star.awt.FontUnderline;
import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPages;
import com.sun.star.drawing.XDrawPagesSupplier;
import com.sun.star.drawing.XDrawView;
import com.sun.star.drawing.XMasterPageTarget;
import com.sun.star.drawing.XMasterPagesSupplier;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.frame.XController;
import com.sun.star.frame.XModel;
import com.sun.star.lang.Locale;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.linguistic2.SingleProofreadingError;
import com.sun.star.presentation.XHandoutMasterSupplier;
import com.sun.star.presentation.XPresentationPage;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.uno.UnoRuntime;

/**
 * Some tools to get information of LibreOffice Impress context
 * @since 1.0
 * @author Fred Kruse
 */
public class WtOfficeDrawTools {
  
  private static boolean debugMode = false;

  /** 
   * get the page count for standard pages
   */
  public static int getDrawPageCount(XComponent xComponent) throws Throwable {
    XDrawPagesSupplier xDrawPagesSupplier = UnoRuntime.queryInterface(XDrawPagesSupplier.class, xComponent);
    XDrawPages xDrawPages = xDrawPagesSupplier.getDrawPages();
    return xDrawPages.getCount();
  }

  /** 
   * get draw page by index
   */
  public static XDrawPage getDrawPageByIndex(XComponent xComponent, int nIndex)
      throws Throwable {
    XDrawPagesSupplier xDrawPagesSupplier = UnoRuntime.queryInterface(XDrawPagesSupplier.class, xComponent);
    XDrawPages xDrawPages = xDrawPagesSupplier.getDrawPages();
    return UnoRuntime.queryInterface(XDrawPage.class, xDrawPages.getByIndex( nIndex ));
  }

  /** 
   * creates and inserts a draw page into the giving position,
   * the method returns the new created page
   */
  public static XDrawPage insertNewDrawPageByIndex(XComponent xComponent, int nIndex) throws Throwable {
    XDrawPagesSupplier xDrawPagesSupplier = UnoRuntime.queryInterface(XDrawPagesSupplier.class, xComponent);
    XDrawPages xDrawPages = xDrawPagesSupplier.getDrawPages();
    return xDrawPages.insertNewByIndex( nIndex );
  }

  /** 
   * get size of the given page
   */
  public static Size getPageSize( XDrawPage xDrawPage ) 
      throws Throwable {
    XPropertySet xPageProperties = UnoRuntime.queryInterface( XPropertySet.class, xDrawPage );
    return new Size(
        ((Integer)xPageProperties.getPropertyValue( "Width" )).intValue(),
        ((Integer)xPageProperties.getPropertyValue( "Height" )).intValue() );
  }

  /** 
   * get the page count for master pages
   */
  public static int getMasterPageCount(XComponent xComponent) {
    XMasterPagesSupplier xMasterPagesSupplier = UnoRuntime.queryInterface(XMasterPagesSupplier.class, xComponent);
    XDrawPages xDrawPages = xMasterPagesSupplier.getMasterPages();
    return xDrawPages.getCount();
  }

  /** 
   * get master page by index
   */
  public static XDrawPage getMasterPageByIndex(XComponent xComponent, int nIndex) throws Throwable {
    XMasterPagesSupplier xMasterPagesSupplier = UnoRuntime.queryInterface(XMasterPagesSupplier.class, xComponent);
    XDrawPages xDrawPages = xMasterPagesSupplier.getMasterPages();
    return UnoRuntime.queryInterface(XDrawPage.class, xDrawPages.getByIndex( nIndex ));
  }

  /** 
   * creates and inserts a new master page into the giving position,
   * the method returns the new created page
   */
  public static XDrawPage insertNewMasterPageByIndex(XComponent xComponent, int nIndex) throws Throwable {
    XMasterPagesSupplier xMasterPagesSupplier = UnoRuntime.queryInterface(XMasterPagesSupplier.class, xComponent);
    XDrawPages xDrawPages = xMasterPagesSupplier.getMasterPages();
    return xDrawPages.insertNewByIndex( nIndex );
  }

  /** 
   * sets given masterpage at the drawpage
   */
  public static void setMasterPage(XDrawPage xDrawPage, XDrawPage xMasterPage) throws Throwable {
    XMasterPageTarget xMasterPageTarget = UnoRuntime.queryInterface(XMasterPageTarget.class, xDrawPage);
    xMasterPageTarget.setMasterPage( xMasterPage );
  }

  /** 
   * test if a Presentation Document is supported.
   * This is important, because only presentation documents
   * have notes and handout pages
   */
  public static boolean isImpressDocument(XComponent xComponent) throws Throwable {
    XServiceInfo xInfo = UnoRuntime.queryInterface(XServiceInfo.class, xComponent);
    if (xInfo == null) {
      return false;
    }
    return xInfo.supportsService("com.sun.star.presentation.PresentationDocument");
  }

  /** 
   * in impress documents each normal draw page has a corresponding notes page
   */
  public static XDrawPage getNotesPage(XDrawPage xDrawPage) throws Throwable {
    XPresentationPage aPresentationPage = UnoRuntime.queryInterface(XPresentationPage.class, xDrawPage);
    return aPresentationPage.getNotesPage();
  }

  /** 
   * in impress each documents has one handout page
   */
  public static XDrawPage getHandoutMasterPage(XComponent xComponent) throws Throwable {
    XHandoutMasterSupplier aHandoutMasterSupplier = UnoRuntime.queryInterface(XHandoutMasterSupplier.class, xComponent);
    return aHandoutMasterSupplier.getHandoutMasterPage();
  }

  /**
   * get shapes of a page
   */
  public static XShapes getShapes(XDrawPage xPage) throws Throwable {
    return UnoRuntime.queryInterface( XShapes.class, xPage );
  }
  
  /**
   * get all paragraphs from Text of a shape
   */
  private static int getAllParagraphsFromText(int nPara, List<String> paragraphs, 
      List<Locale> locales, XText xText) throws Throwable {
    if (xText != null) {
      XTextCursor xTextCursor = xText.createTextCursor();
      xTextCursor.gotoStart(false);
      String sText = xText.getString();
      int kStart = 0;
      int k;
      for (k = 0; k < sText.length(); k++) {
        if (sText.charAt(k) == WtOfficeTools.SINGLE_END_OF_PARAGRAPH.charAt(0)) {
          paragraphs.add(sText.substring(kStart, k));
          nPara++;
          xTextCursor.goRight((short)(k - kStart), true);
          XPropertySet xParaPropSet = UnoRuntime.queryInterface(XPropertySet.class, xTextCursor);
          locales.add((Locale) xParaPropSet.getPropertyValue("CharLocale"));
          xTextCursor.goRight((short)1, false);
          kStart = k + 1;
        }
      }
      if (k > kStart) {
        paragraphs.add(sText.substring(kStart, k));
        nPara++;
        xTextCursor.goRight((short)(k - kStart), true);
        XPropertySet xParaPropSet = UnoRuntime.queryInterface(XPropertySet.class, xTextCursor);
        locales.add((Locale) xParaPropSet.getPropertyValue("CharLocale"));
      }
    }
    return nPara;
  }

  /**
   * get all paragraphs of a impress document
   */
  public static ParagraphContainer getAllParagraphs(XComponent xComponent) {
    List<String> paragraphs = new ArrayList<>();
    List<Locale> locales = new ArrayList<>();
    List<Integer> pageBegins = new ArrayList<>();
    int nPara = 0;
    try {
      int pageCount = WtOfficeDrawTools.getDrawPageCount(xComponent);
      for (int i = 0; i < pageCount; i++) {
        XDrawPage xDrawPage = null;
        for (int n = 0; n < 2; n++) {
          if (n == 0) {
            xDrawPage = WtOfficeDrawTools.getDrawPageByIndex(xComponent, i);
          } else {
            xDrawPage = getNotesPage(xDrawPage);
          }
          XShapes xShapes = WtOfficeDrawTools.getShapes(xDrawPage);
          int nShapes = xShapes.getCount();
          if (nShapes > 0) {
            pageBegins.add(nPara);
          }
          for(int j = 0; j < nShapes; j++) {
            Object oShape = xShapes.getByIndex(j);
            XShape xShape = UnoRuntime.queryInterface(XShape.class, oShape);
            if (xShape != null) {
              XText xText = UnoRuntime.queryInterface(XText.class, xShape);
              nPara = getAllParagraphsFromText(nPara, paragraphs, locales, xText);
            } else {
              WtMessageHandler.printToLogFile("OfficeDrawTools: getAllParagraphs: xShape " + j + " is null");
            }
          }
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return new ParagraphContainer(paragraphs, locales, pageBegins);
  }

  /**
   * find the Paragraph to change in a shape and change the text
   * returns -1 if it was found
   * returns the last number of paragraph otherwise
   */
  private static int changeTextOfParagraphInText(int nParaCount, int nPara, int beginn, int length, String replace, XText xText)
        throws Throwable {
    if (xText != null) {
      XTextCursor xTextCursor = xText.createTextCursor();
      String sText = xText.getString();
      if (nParaCount == nPara) {
        xTextCursor.gotoStart(false);
        xTextCursor.goRight((short)beginn, false);
        xTextCursor.goRight((short)length, true);
        xTextCursor.setString(replace);
        return -1;
      }
      int lastParaEnd = 0;
      for (int k = 0; k < sText.length(); k++) {
        if (sText.charAt(k) == WtOfficeTools.SINGLE_END_OF_PARAGRAPH.charAt(0)) {
          nParaCount++;
          lastParaEnd = k;
          if (nParaCount == nPara) {
            xTextCursor.gotoStart(false);
            xTextCursor.goRight((short)(beginn + k + 1), false);
            xTextCursor.goRight((short)length, true);
            xTextCursor.setString(replace);
            //  Note: The faked change of position is a workaround to trigger the notification of a change
            return -1;
          }
        }
      }
      if (lastParaEnd < sText.length() - 1) {
        nParaCount++;
      }
    }
    return nParaCount;
  }

  /**
   * change the text of a paragraph
   */
  public static void changeTextOfParagraph(int nPara, int beginn, int length, String replace, XComponent xComponent) {
    try {
      int nParaCount = 0;
      int pageCount = WtOfficeDrawTools.getDrawPageCount(xComponent);
      for (int i = 0; i < pageCount; i++) {
        XDrawPage xDrawPage = null;
        for (int n = 0; n < 2; n++) {
          if (n == 0) {
            xDrawPage = WtOfficeDrawTools.getDrawPageByIndex(xComponent, i);
          } else {
            xDrawPage = getNotesPage(xDrawPage);
          }
          XShapes xShapes = WtOfficeDrawTools.getShapes(xDrawPage);
          int nShapes = xShapes.getCount();
          for(int j = 0; j < nShapes; j++) {
            Object oShape = xShapes.getByIndex(j);
            XShape xShape = UnoRuntime.queryInterface(XShape.class, oShape);
            if (xShape != null) {
              XText xText = UnoRuntime.queryInterface(XText.class, xShape);
              nParaCount = changeTextOfParagraphInText(nParaCount, nPara, beginn, length, replace, xText);
              if (nParaCount < 0) {
                //  Note: The faked change of position is a workaround to trigger the notification of a change
                Point p = xShape.getPosition();
                xShape.setPosition(p);
                return;
              }
            } else {
              WtMessageHandler.printToLogFile("OfficeDrawTools: changeTextOfParagraph: xShape " + j + " is null");
            }
          }
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }

  /**
   * find the Paragraph to change in a shape and change the locale
   * returns -1 if it was found
   * returns the last number of paragraph otherwise
   */
  private static int changeLocaleOfParagraphInText(int nParaCount, int nPara, int beginn, int length, Locale locale, 
      XText xText) throws Throwable {
    if (xText != null) {
      XTextCursor xTextCursor = xText.createTextCursor();
      String sText = xText.getString();
      if (nParaCount == nPara) {
        xTextCursor.gotoStart(false);
        xTextCursor.goRight((short)beginn, false);
        xTextCursor.goRight((short)length, true);
        XPropertySet xParaPropSet = UnoRuntime.queryInterface(XPropertySet.class, xTextCursor);
        xParaPropSet.setPropertyValue("CharLocale", locale);
        return -1;
      }
      int lastParaEnd = 0;
      for (int k = 0; k < sText.length(); k++) {
        if (sText.charAt(k) == WtOfficeTools.SINGLE_END_OF_PARAGRAPH.charAt(0)) {
          nParaCount++;
          lastParaEnd = k;
          if (nParaCount == nPara) {
            xTextCursor.gotoStart(false);
            xTextCursor.goRight((short)(beginn + k + 1), false);
            xTextCursor.goRight((short)length, true);
            XPropertySet xParaPropSet = UnoRuntime.queryInterface(XPropertySet.class, xTextCursor);
            xParaPropSet.setPropertyValue("CharLocale", locale);
            return -1;
          }
        }
      }
      if (lastParaEnd < sText.length() - 1) {
        nParaCount++;
      }
    }
    return nParaCount;
  }
  
  /**
   * change the language of a paragraph
   */
  public static void setLanguageOfParagraph(int nPara, int beginn, int length, Locale locale, XComponent xComponent) {
    try {
      int nParaCount = 0;
      int pageCount = WtOfficeDrawTools.getDrawPageCount(xComponent);
      for (int i = 0; i < pageCount; i++) {
        XDrawPage xDrawPage = null;
        for (int n = 0; n < 2; n++) {
          if (n == 0) {
            xDrawPage = WtOfficeDrawTools.getDrawPageByIndex(xComponent, i);
          } else {
            xDrawPage = getNotesPage(xDrawPage);
          }
          XShapes xShapes = WtOfficeDrawTools.getShapes(xDrawPage);
          int nShapes = xShapes.getCount();
          for(int j = 0; j < nShapes; j++) {
            Object oShape = xShapes.getByIndex(j);
            XShape xShape = UnoRuntime.queryInterface(XShape.class, oShape);
            if (xShape != null) {
              XText xText = UnoRuntime.queryInterface(XText.class, xShape);
              nParaCount = changeLocaleOfParagraphInText(nParaCount, nPara, beginn, length, locale, xText);
              if (nParaCount < 0) {
                //  Note: The faked change of position is a workaround to trigger the notification of a change
                Point p = xShape.getPosition();
                xShape.setPosition(p);
                return;
              }
            } else {
              WtMessageHandler.printToLogFile("OfficeDrawTools: setLanguageOfParagraph: xShape " + j + " is null");
            }
          }
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }

  /**
   * find the Paragraph in a shape
   * returns -1 if it was found
   * returns the last number of paragraph otherwise
   */
  private static int findParaInText(int nParaCount, int nPara, XText xText) throws Throwable {
    if (xText != null) {
      String sText = xText.getString();
      if (nParaCount == nPara && sText.length() > 0) {
        return -1;
      }
      int kStart = 0;
      for (int k = 0; k < sText.length(); k++) {
        if (sText.charAt(k) == WtOfficeTools.SINGLE_END_OF_PARAGRAPH.charAt(0)) {
          nParaCount++;
          kStart = k + 1;
          if (nParaCount == nPara) {
            return -1;
          }
        }
      }
      if (kStart < sText.length()) {
        nParaCount++;
      }
    }
    return nParaCount;
  }
  
  /**
   * set the current draw page for a given paragraph
   */
  public static void setCurrentPage(int nPara, XComponent xComponent) {
    try {
      XModel xModel = UnoRuntime.queryInterface(XModel.class, xComponent);
      XController xController = xModel.getCurrentController();
      XDrawView xDrawView = UnoRuntime.queryInterface(XDrawView.class, xController);
      int nParaCount = 0;
      int pageCount = WtOfficeDrawTools.getDrawPageCount(xComponent);
      for (int i = 0; i < pageCount; i++) {
        XDrawPage xDrawPage = null;
        for (int n = 0; n < 2; n++) {
          if (n == 0) {
            xDrawPage = WtOfficeDrawTools.getDrawPageByIndex(xComponent, i);
          } else {
            xDrawPage = getNotesPage(xDrawPage);
          }
          XShapes xShapes = WtOfficeDrawTools.getShapes(xDrawPage);
          int nShapes = xShapes.getCount();
          for(int j = 0; j < nShapes; j++) {
            Object oShape = xShapes.getByIndex(j);
            XShape xShape = UnoRuntime.queryInterface(XShape.class, oShape);
            if (xShape != null) {
              XText xText = UnoRuntime.queryInterface(XText.class, xShape);
              nParaCount = findParaInText(nParaCount, nPara, xText);
              if (nParaCount < 0) {
                xDrawView.setCurrentPage(xDrawPage);
                return;
              }
            } else {
              WtMessageHandler.printToLogFile("OfficeDrawTools: setCurrentPage: xShape " + j + " is null");
            }
          }
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }

  /**
   * get the first paragraph of current draw page
   */
  public static int getParagraphNumberFromCurrentPage(XComponent xComponent) {
    int nParaCount = 0;
    try {
      XModel xModel = UnoRuntime.queryInterface(XModel.class, xComponent);
      XController xController = xModel.getCurrentController();
      XDrawView xDrawView = UnoRuntime.queryInterface(XDrawView.class, xController);
      XDrawPage xCurrentDrawPage = xDrawView.getCurrentPage();
      int pageCount = WtOfficeDrawTools.getDrawPageCount(xComponent);
      for (int i = 0; i < pageCount; i++) {
        XDrawPage xDrawPage = null;
        for (int n = 0; n < 2; n++) {
          if (n == 0) {
            xDrawPage = WtOfficeDrawTools.getDrawPageByIndex(xComponent, i);
          } else {
            xDrawPage = getNotesPage(xDrawPage);
          }
          if (xDrawPage.equals(xCurrentDrawPage)) {
            return nParaCount;
          }
          XShapes xShapes = WtOfficeDrawTools.getShapes(xDrawPage);
          int nShapes = xShapes.getCount();
          for(int j = 0; j < nShapes; j++) {
            Object oShape = xShapes.getByIndex(j);
            XShape xShape = UnoRuntime.queryInterface(XShape.class, oShape);
            if (xShape != null) {
              XText xText = UnoRuntime.queryInterface(XText.class, xShape);
              nParaCount = findParaInText(nParaCount, -1, xText);
            } else {
              WtMessageHandler.printToLogFile("OfficeDrawTools: getParagraphFromCurrentPage: xShape " + j + " is null");
            }
          }
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return nParaCount;
  }

  /**
   * true: if paragraph is in a notes page
   */
  public static boolean isParagraphInNotesPage(int nPara, XComponent xComponent) {
    int nParaCount = 0;
    try {
      int pageCount = WtOfficeDrawTools.getDrawPageCount(xComponent);
      for (int i = 0; i < pageCount; i++) {
        XDrawPage xDrawPage = null;
        for (int n = 0; n < 2; n++) {
          if (n == 0) {
            xDrawPage = WtOfficeDrawTools.getDrawPageByIndex(xComponent, i);
          } else {
            xDrawPage = getNotesPage(xDrawPage);
          }
          XShapes xShapes = WtOfficeDrawTools.getShapes(xDrawPage);
          int nShapes = xShapes.getCount();
          for(int j = 0; j < nShapes; j++) {
            Object oShape = xShapes.getByIndex(j);
            XShape xShape = UnoRuntime.queryInterface(XShape.class, oShape);
            if (xShape != null) {
              XText xText = UnoRuntime.queryInterface(XText.class, xShape);
              nParaCount = findParaInText(nParaCount, nPara, xText);
              if (nParaCount < 0) {
                return (n == 0 ? false : true);
              }
            } else {
              WtMessageHandler.printToLogFile("OfficeDrawTools: isParagraphInNotesPage: xShape " + j + " is null");
            }
          }
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return false;
  }
  
  public static Locale getDocumentLocale(XComponent xComponent) {
    try {
      int pageCount = WtOfficeDrawTools.getDrawPageCount(xComponent);
      for (int i = 0; i < pageCount; i++) {
        XDrawPage xDrawPage = null;
        for (int n = 0; n < 2; n++) {
          if (n == 0) {
            xDrawPage = WtOfficeDrawTools.getDrawPageByIndex(xComponent, i);
          } else {
            xDrawPage = getNotesPage(xDrawPage);
          }
          XShapes xShapes = WtOfficeDrawTools.getShapes(xDrawPage);
          int nShapes = xShapes.getCount();
          for(int j = 0; j < nShapes; j++) {
            Object oShape = xShapes.getByIndex(j);
            XShape xShape = UnoRuntime.queryInterface(XShape.class, oShape);
            if (xShape != null) {
              XText xText = UnoRuntime.queryInterface(XText.class, xShape);
              if (xText != null) {
                XTextCursor xTextCursor = xText.createTextCursor();
                XPropertySet xParaPropSet = UnoRuntime.queryInterface(XPropertySet.class, xTextCursor);
                return ((Locale) xParaPropSet.getPropertyValue("CharLocale"));
              }
            } else {
              WtMessageHandler.printToLogFile("OfficeDrawTools: getDocumentLocale: xShape " + j + " is null");
            }
          }
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return null;
  }
  
  /**
   * mark up the errors in a paragraph
   */
  private static void markupOneParagraph(int nParaBeginn, SingleProofreadingError error, 
      UndoMarkupContainer undoMarkUp, XTextCursor xTextCursor) throws Throwable {
    if (undoMarkUp == null) {
      return;
    }
    xTextCursor.gotoStart(false);
    xTextCursor.goRight((short)nParaBeginn, false);
    if (error != null) {
      xTextCursor.goRight((short)(error.nErrorStart), false);
      undoMarkUp.nStart = error.nErrorStart;
      undoMarkUp.underline = new UndoMarkupContainer.Underline[error.nErrorLength];
      for (int i = 0; i < error.nErrorLength; i++) {
        if (debugMode) {
          WtMessageHandler.printToLogFile("OfficeDrawTools: markupOneParagraph: i: " + i);
        }
        xTextCursor.goRight((short)1, true);
        XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, xTextCursor);
        Object o = xPropertySet.getPropertyValue("CharUnderline");
        short font;
        boolean hasColor;
        int color;
        if (o != null) {
          font = (short) o;
          o = xPropertySet.getPropertyValue("CharUnderlineHasColor");
          if (o != null) {
            hasColor = (boolean) o;
            o = xPropertySet.getPropertyValue("CharUnderlineColor");
            if (o != null) {
              color = (int) o;
            } else {
              color = (int) 0;
            }
          } else {
            hasColor = (boolean) false;
            color = (int) 0;
          }
        } else {
          font = (short) 0;
          hasColor = (boolean) false;
          color = (int) 0;
        }
        undoMarkUp.underline[i] = new UndoMarkupContainer.Underline (font, color, hasColor);
        xPropertySet.setPropertyValue("CharUnderline", FontUnderline.WAVE);
        PropertyValue[] properties = error.aProperties;
        color = -1;
        for (PropertyValue property : properties) {
          if ("LineColor".equals(property.Name)) {
            color = (int) property.Value;
          }
        }
        if (color < 0) {
          color = Color.blue.getRGB() & 0xFFFFFF;
        }
        xPropertySet.setPropertyValue("CharUnderlineColor", color);
        xPropertySet.setPropertyValue("CharUnderlineHasColor", true);
        xTextCursor.goLeft((short)1, false);
        xTextCursor.goRight((short)1, false);
      }
    } else {
      if (undoMarkUp.underline != null) {
        xTextCursor.goRight((short)(undoMarkUp.nStart), false);
        for (int i = 0; i < undoMarkUp.underline.length; i++) {
          xTextCursor.goRight((short)1, true);
          XPropertySet xPropertySet = UnoRuntime.queryInterface(XPropertySet.class, xTextCursor);
          xPropertySet.setPropertyValue("CharUnderlineColor", undoMarkUp.underline[i].color);
          xPropertySet.setPropertyValue("CharUnderlineHasColor", undoMarkUp.underline[i].hasColor);
          xPropertySet.setPropertyValue("CharUnderline", undoMarkUp.underline[i].font);
          xTextCursor.goLeft((short)1, false);
          xTextCursor.goRight((short)1, false);
        }
      }
    }
  }

  /**
   * remove the last mark up
   */
  public static void removeMarkup(UndoMarkupContainer undoMarkup, XComponent xComponent) {
    if (undoMarkup != null) {
      setMarkup(undoMarkup.nPara, null, undoMarkup, xComponent);
    }
  }

  /**
   * mark up an error in a paragraph /remove last mark up if error == null
   */
  public static void setMarkup(int nPara, SingleProofreadingError error, UndoMarkupContainer undoMarkup, XComponent xComponent) {
    if (undoMarkup == null) {
      WtMessageHandler.printToLogFile("OfficeDrawTools: markupParagraph: UndoMarkupContainer is null: return");
      return;
    }
    try {
      int nParaCount = 0;
      int pageCount = WtOfficeDrawTools.getDrawPageCount(xComponent);
      if (error == null) {
        nPara = undoMarkup.nPara;
      }
      for (int i = 0; i < pageCount; i++) {
        XDrawPage xDrawPage = null;
        for (int n = 0; n < 2; n++) {
          if (n == 0) {
            xDrawPage = WtOfficeDrawTools.getDrawPageByIndex(xComponent, i);
          } else {
            xDrawPage = getNotesPage(xDrawPage);
          }
          XShapes xShapes = WtOfficeDrawTools.getShapes(xDrawPage);
          int nShapes = xShapes.getCount();
          for(int j = 0; j < nShapes; j++) {
            Object oShape = xShapes.getByIndex(j);
            XShape xShape = UnoRuntime.queryInterface(XShape.class, oShape);
            if (xShape != null) {
              XText xText = UnoRuntime.queryInterface(XText.class, xShape);
              if (xText != null) {
                XTextCursor xTextCursor = xText.createTextCursor();
                String sText = xText.getString();
                if (nParaCount == nPara) {
                  if (debugMode) {
                    WtMessageHandler.printToLogFile("OfficeDrawTools: markupParagraphs: sText: " + sText);
                    if (error != null) {
                      WtMessageHandler.printToLogFile("OfficeDrawTools: markupParagraphs: error Beginn: " + error.nErrorStart
                          + ", error Length: " + error.nErrorLength);
                    }
                  }
                  if (error != null) {
                    undoMarkup.nPara = nPara;
                  }
                  markupOneParagraph(0, error, undoMarkup, xTextCursor);
                  return;
                }
                int lastParaEnd = 0;
                if (debugMode && nPara < 12) {
                  WtMessageHandler.printToLogFile("OfficeDrawTools: markupParagraphs: sText: " + sText);
                }
                for (int k = 0; k < sText.length(); k++) {
                  if (sText.charAt(k) == WtOfficeTools.SINGLE_END_OF_PARAGRAPH.charAt(0)) {
                    nParaCount++;
                    if (debugMode && nPara < 12) {
                      WtMessageHandler.printToLogFile("OfficeDrawTools: markupParagraphs: nPara: " + nPara + ", nParaCount: " + nParaCount);
                    }
                    lastParaEnd = k;
                    if (nParaCount == nPara) {
                      if (debugMode) {
                        WtMessageHandler.printToLogFile("OfficeDrawTools: markupParagraphs: sText: " + sText);
                        if (error != null) {
                          WtMessageHandler.printToLogFile("OfficeDrawTools: markupParagraphs: error Beginn: " + error.nErrorStart
                            + ", error Length: " + error.nErrorLength);
                        }
                      }
                      if (error != null) {
                        undoMarkup.nPara = nPara;
                      }
                      markupOneParagraph(k + 1, error, undoMarkup, xTextCursor);
                      return;
                    }
                  }
                }
                if (lastParaEnd < sText.length() - 1) {
                  nParaCount++;
                }
                if (debugMode && nPara < 12) {
                  WtMessageHandler.printToLogFile("OfficeDrawTools: markupParagraphs: nPara: " + nPara + ", nParaCount: " + nParaCount);
                }
              }
            } else {
              WtMessageHandler.printToLogFile("OfficeDrawTools: markupParagraph: xShape " + j + " is null");
            }
          }
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }

  public static class ParagraphContainer {
    public List<String> paragraphs;
    public List<Locale> locales;
    public List<Integer> pageBegins;
    
    public ParagraphContainer(List<String> paragraphs, List<Locale> locales, List<Integer> pageBegins) {
      this.paragraphs = paragraphs;
      this.locales = locales;
      this.pageBegins = pageBegins;
    }
  }

  public static class UndoMarkupContainer {
    int nPara;
    int nStart;
    Underline underline[];
    
    public UndoMarkupContainer () {
    }
    
    UndoMarkupContainer (int nPara, int nStart, Underline[] underline) {
      this.nPara = nPara;
      this.nStart = nStart;
      this.underline = underline;
    }
    
    public static class Underline {
      short font;
      int color;
      boolean hasColor;
      
      Underline(short font, int color, boolean hasColor){
        this.font = font;
        this.color = color;
        this.hasColor = hasColor;
      }
    }
  }
  
}



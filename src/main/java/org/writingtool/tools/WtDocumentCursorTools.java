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

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.jetbrains.annotations.Nullable;
import org.writingtool.WtDocumentCache;
import org.writingtool.WtDocumentCache.TextParagraph;

import com.sun.star.beans.Property;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNameContainer;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawPageSupplier;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.lang.XComponent;
import com.sun.star.style.XStyleFamiliesSupplier;
import com.sun.star.text.XFootnote;
import com.sun.star.text.XFootnotesSupplier;
import com.sun.star.text.XMarkingAccess;
import com.sun.star.text.TextMarkupType;
import com.sun.star.text.XEndnotesSupplier;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextFramesSupplier;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextSection;
import com.sun.star.text.XTextTable;
import com.sun.star.text.XTextTablesSupplier;
import com.sun.star.uno.UnoRuntime;

/**
 * Information about Paragraphs of LibreOffice/OpenOffice documents
 * on the basis of the LO/OO text cursor
 * @since 1.0
 * @author Fred Kruse
 */
public class WtDocumentCursorTools {
  
  public static int TEXT_TYPE_NORMAL = -1;
  public static int TEXT_TYPE_AUTOMATIC = -2;
  
  public final static String HeaderFooterTypes[] = { "HeaderText", 
      "HeaderTextRight",
      "HeaderTextLeft",
      "HeaderTextFirst", 
      "FooterText", 
      "FooterTextRight", 
      "FooterTextLeft",
      "FooterTextFirst" 
  };

  private static int isBusy = 0;
  
  private boolean isCheckedSortedTextId = false;
  private boolean hasSortedTextId = false;
  private boolean isDisposed = false;

  private XParagraphCursor xPCursor;
  private XTextCursor xTextCursor;
  private XTextDocument curDoc;
  
  public WtDocumentCursorTools(XComponent xComponent) {
    isBusy++;
    try {
      if (!isDisposed) {
        curDoc = UnoRuntime.queryInterface(XTextDocument.class, xComponent);
        WtOfficeTools.waitForLtDictionary();
        xTextCursor = getCursor();
        xPCursor = createParagraphCursor();
      }
    } finally {
      isBusy--;
    }
  }
  
  /**
   * document is disposed: set all class variables to null
   */
  public void setDisposed() {
    xPCursor = null;
    xTextCursor = null;
    curDoc = null;
    isDisposed = true;
  }
  
  public XTextDocument getTextDocument() {
    return curDoc;
  }

  /** 
   * Returns the text cursor (if any)
   * Returns null if it fails
   */
  @Nullable
  public XTextCursor getCursor() {
    isBusy++;
    try {
      if (curDoc == null) {
        return null;
      }
      XText xText = curDoc.getText();
      if (xText == null) {
        return null;
      } else {
        XTextRange xStart = xText.getStart();
        try {
          return xText.createTextCursorByRange(xStart);
        } catch (Throwable t) {
          return null;           // Return null without message - is needed for documents without main text (e.g. only a table)
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return null;           // Return null as method failed
    } finally {
      isBusy--;
    }
  }

  /** 
   * Returns ParagraphCursor from TextCursor 
   * Returns null if it fails
   */
  @Nullable
  private XParagraphCursor createParagraphCursor() {
    isBusy++;
    try {
      if (xTextCursor == null) {
        return null;
      }
      return UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return null;           // Return null as method failed
    } finally {
      isBusy--;
    }
  }
  
  /** 
   * Returns the TextCursor of the Document
   * Returns null if it fails
   */
  @Nullable
  public XTextCursor getTextCursor() {
    return xTextCursor;
  }
  
  /** 
   * Returns ParagraphCursor from TextCursor 
   * Returns null if it fails
   */
  @Nullable
  public XParagraphCursor getParagraphCursor() {
    return xPCursor;
  }
  
  /** 
   * Returns Number of all Text Paragraphs of Document without footnotes etc.  
   * Returns 0 if it fails
   */
  public int getNumberOfAllTextParagraphs() {
    isBusy++;
    try {
      if (xPCursor == null) {
        return 0;
      }
      xPCursor.gotoStart(false);
      int nPara = 1;
      while (xPCursor.gotoNextParagraph(false)) nPara++;
      return nPara;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return 0;              // Return 0 as method failed
    } finally {
      isBusy--;
    }
  }

  /**
   * Give back a list of positions of deleted characters 
   * or null if there are no
   */
  private static List<Integer> getDeletedCharacters(XParagraphCursor xPCursor, boolean withDeleted) {
    if (xPCursor == null) {
      WtMessageHandler.printToLogFile("DocumentCursorTools: Properties: ParagraphCursor == null");
      return null;
    }
    List<Integer> deletePositions = new ArrayList<Integer>();
    if (!withDeleted) {
      return deletePositions;
    }
    int num = 0;
    try {
      XEnumerationAccess xParaEnumAccess = UnoRuntime.queryInterface(XEnumerationAccess.class, xPCursor);
      if (xParaEnumAccess == null) {
        return null;
      }
      XEnumeration xParaEnum = xParaEnumAccess.createEnumeration();
      if (xParaEnum == null) {
        return null;
      }
      while (xParaEnum.hasMoreElements()) {
        XEnumerationAccess xEnumAccess = null;
        if (xParaEnum.hasMoreElements()) {
          xEnumAccess = UnoRuntime.queryInterface(XEnumerationAccess.class, xParaEnum.nextElement());
        }
        if (xEnumAccess == null) {
          continue;
        }
        XEnumeration xEnum = xEnumAccess.createEnumeration();
        if (xEnum == null) {
          continue;
        }
        boolean isDelete = false;
        while (xEnum.hasMoreElements()) {
          XTextRange xPortion = (XTextRange) UnoRuntime.queryInterface(XTextRange.class, xEnum.nextElement());
          if (xPortion == null) {
            continue;
          }
          XPropertySet xTextPortionPropertySet = UnoRuntime.queryInterface(XPropertySet.class, xPortion);
          if (xTextPortionPropertySet == null) {
            continue;
          }
          String textPortionType = (String) xTextPortionPropertySet.getPropertyValue("TextPortionType");
          if (textPortionType != null && textPortionType.equals("Redline")) {
            String redlineType = (String) xTextPortionPropertySet.getPropertyValue("RedlineType");
            if (redlineType != null && redlineType.equals("Delete")) {
              isDelete = !isDelete;
            }
          } else {
            int portionLen = xPortion.getString().length();
            if (isDelete) {
              for (int i = num; i < num + portionLen; i++) {
                deletePositions.add(i);
              }
            }
            num += portionLen;
          }
        }
      }
    } catch (Throwable e) {
      WtMessageHandler.printException(e);
    }
    if (deletePositions.isEmpty()) {
      return null;
    }
    return deletePositions;
  }

  /** 
   * Returns all Paragraphs of Document without footnotes etc.  
   * Returns null if it fails
   */
  @Nullable
  public
  DocumentText getAllTextParagraphs(boolean withDeleted) {
    isBusy++;
    try {
      List<String> allParas = new ArrayList<>();
      Map<Integer, Integer> headingNumbers = new HashMap<Integer, Integer>();
      List<Integer> automaticTextParagraphs = new ArrayList<Integer>();
      List<Integer> sortedTextIds = null;
      List<List<Integer>> deletedCharacters = new ArrayList<List<Integer>>();
      if (xPCursor == null) {
        return null;
      }
      int paraNum = 0;
      xPCursor.gotoStart(false);
      xPCursor.gotoStartOfParagraph(false);
      xPCursor.gotoEndOfParagraph(true);
      if (sortedTextIds == null && (hasSortedTextId || !isCheckedSortedTextId)) {
        isCheckedSortedTextId = true;
        if (hasSortedTextId || supportOfSortedTextId(xPCursor)) {
          hasSortedTextId = true;
          sortedTextIds = new ArrayList<Integer>();
        }
      }
      allParas.add(new String(xPCursor.getString()));
      deletedCharacters.add(getDeletedCharacters(xPCursor, withDeleted));
      int textType = getTextType();
      if (textType >= 0) {
        headingNumbers.put(paraNum, textType);
      } else if (textType == TEXT_TYPE_AUTOMATIC) {
        headingNumbers.put(paraNum, 0);
        automaticTextParagraphs.add(paraNum);
      }
      if (sortedTextIds != null) {
        sortedTextIds.add(getSortedTextId(xPCursor));
      }
      while (xPCursor.gotoNextParagraph(false)) {
        xPCursor.gotoStartOfParagraph(false);
        xPCursor.gotoEndOfParagraph(true);
        allParas.add(new String(xPCursor.getString()));
        deletedCharacters.add(getDeletedCharacters(xPCursor, withDeleted));
        paraNum++;
        textType = getTextType();
        if (textType >= 0) {
          headingNumbers.put(paraNum, textType);
        } else if (textType == TEXT_TYPE_AUTOMATIC) {
          headingNumbers.put(paraNum, 0);
          automaticTextParagraphs.add(paraNum);
        } 
        if (sortedTextIds != null) {
          sortedTextIds.add(getSortedTextId(xPCursor));
        }
      }
      return new DocumentText(allParas, headingNumbers, automaticTextParagraphs, sortedTextIds, deletedCharacters);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return null;           // Return null as method failed
    } finally {
      isBusy--;
    }
  }

  /**
   * Paragraph is Header or Title
   */
  private int getTextType() {
    String paraStyleName = null;
    XPropertySet xParagraphPropertySet = null;
    try {
      xParagraphPropertySet = UnoRuntime.queryInterface(XPropertySet.class, xPCursor.getStart());
      if (xParagraphPropertySet == null) {
        return TEXT_TYPE_NORMAL;
      }
      Object o = xParagraphPropertySet.getPropertyValue("ParaStyleName");
      if (o != null) {
        paraStyleName = (String) o;
      }
    } catch (Throwable e) {
      WtMessageHandler.printException(e);
      return TEXT_TYPE_NORMAL;
    }
    try {
      XTextSection xTextSection = UnoRuntime.queryInterface(XTextSection.class, xParagraphPropertySet.getPropertyValue("TextSection"));
      if (xTextSection != null) {
        xParagraphPropertySet = UnoRuntime.queryInterface(XPropertySet.class, xTextSection);
        Object o = xParagraphPropertySet.getPropertyValue("IsProtected");
        if (o != null && (boolean) o) {
          return TEXT_TYPE_AUTOMATIC;
        }
      }
    } catch (Throwable e) {
      //  if there is an exception go on with analysis - TextType is not automatic
    }
    if (paraStyleName == null) {
      return TEXT_TYPE_NORMAL;
    }
    if (paraStyleName.equals("Title") || paraStyleName.equals("Subtitle")|| paraStyleName.equals("Heading")) {
      return 0;
    } else if (paraStyleName.startsWith("Heading")) {
      String numberString = paraStyleName.substring(7).trim();
      if (numberString.isEmpty()) { 
        return 0;
      }
      int ret = 0;
      try {
        ret = Integer.parseInt(numberString);
      } catch (Throwable e) {
//        MessageHandler.printToLogFile("DocumentCursorTools: getTextType: paraStyleName: " + paraStyleName);
//        MessageHandler.printException(e);
      }
      return ret;
    } else if (paraStyleName.startsWith("Contents")) {
      return TEXT_TYPE_AUTOMATIC;
    } else {
      return TEXT_TYPE_NORMAL;
    }
  }
  
  /**
   * Print properties to log file for the actual position of cursor
   */
  void printProperties() {
    if (xPCursor == null) {
      WtMessageHandler.printToLogFile("DocumentCursorTools: Properties: ParagraphCursor == null");
      return;
    }
    XPropertySet xParagraphPropertySet = UnoRuntime.queryInterface(XPropertySet.class, xPCursor.getStart());
    Property[] properties = xParagraphPropertySet.getPropertySetInfo().getProperties();
    for (Property property : properties) {
      WtMessageHandler.printToLogFile("DocumentCursorTools: Properties: Name: " + property.Name + ", Type: " + property.Type);
    }
    try {
      WtMessageHandler.printToLogFile("DocumentCursorTools: Properties: SortedTextId: " + xParagraphPropertySet.getPropertyValue("SortedTextId") + "\n");
    } catch (Throwable e) {
      WtMessageHandler.printException(e);
    }
  }
  
  /**
   * Print properties to log file for the actual position of cursor
   */
  private static int getSortedTextId(XParagraphCursor xPCursor) {
    try {
      if (xPCursor == null) {
        WtMessageHandler.printToLogFile("DocumentCursorTools: Properties: ParagraphCursor == null");
        return -1;
      }
      XPropertySet xParagraphPropertySet = UnoRuntime.queryInterface(XPropertySet.class, xPCursor.getStart());
      if (xParagraphPropertySet != null) {
        return (int) xParagraphPropertySet.getPropertyValue("SortedTextId");
      }
    } catch (Throwable e) {
      WtMessageHandler.printException(e);
    }
    return -1;
  }
  
  /**
   * Print properties to log file for the actual position of cursor
   */
  private boolean supportOfSortedTextId(XParagraphCursor xPCursor) {
    try {
      if (xPCursor == null) {
        WtMessageHandler.printToLogFile("DocumentCursorTools: Properties: ParagraphCursor == null");
        return false;
      }
      XPropertySet xParagraphPropertySet = UnoRuntime.queryInterface(XPropertySet.class, xPCursor.getStart());
      if (xParagraphPropertySet != null) {
        return xParagraphPropertySet.getPropertyValue("SortedTextId") != null;
      }
    } catch (Throwable e) {
    }
    return false;
  }
  
  /** 
   * Add all paragraphs of XText to a list of strings
   */
  private List<Integer> addAllParagraphsOfText(XText xText, List<String> sText, 
      List<List<Integer>> deletedCharacters, List<Integer> sortedTextIds, boolean withDeleted) {
    if (xText == null) {
      throw new RuntimeException("XText == null"); 
    }
    if (sText == null) {
      throw new RuntimeException("Text List == null"); 
    }
    if (deletedCharacters == null) {
      throw new RuntimeException("List of deleted Characters == null"); 
    }
    XTextCursor xTextCursor = xText.createTextCursor();
    XParagraphCursor xParagraphCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
    if (xParagraphCursor == null) {
      return sortedTextIds;
    }
    try {
      if (sortedTextIds == null && (hasSortedTextId || !isCheckedSortedTextId)) {
        isCheckedSortedTextId = true;
        if (hasSortedTextId || supportOfSortedTextId(xParagraphCursor)) {
          hasSortedTextId = true;
          sortedTextIds = new ArrayList<Integer>();
        }
      }
    } catch (Exception e) {
      WtMessageHandler.printException(e);
    }
    xParagraphCursor.gotoStart(false);
    do {
      xParagraphCursor.gotoStartOfParagraph(false);
      xParagraphCursor.gotoEndOfParagraph(true);
      sText.add(new String(xParagraphCursor.getString()));
      deletedCharacters.add(getDeletedCharacters(xParagraphCursor, withDeleted));
      if (sortedTextIds != null) {
        sortedTextIds.add(getSortedTextId(xParagraphCursor));
      }
    } while (xParagraphCursor.gotoNextParagraph(false));
    return sortedTextIds;
  }
  
  /** 
   * get all paragraphs as a list of strings
   */
  private List<String> getAllParagraphsOfText(XText xText) {
    List<String> sText = new ArrayList<String>();
    if (xText == null) {
      return sText;
    }
    XTextCursor xTextCursor = xText.createTextCursor();
    XParagraphCursor xParagraphCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
    if (xParagraphCursor == null) {
      return sText;
    }
    xParagraphCursor.gotoStart(false);
    do {
      xParagraphCursor.gotoStartOfParagraph(false);
      xParagraphCursor.gotoEndOfParagraph(true);
      sText.add(new String(xParagraphCursor.getString()));
    } while (xParagraphCursor.gotoNextParagraph(false));
    return sText;
  }
  
  /** 
   * Get the number of all paragraphs of XText
   */
  private static int getNumberOfAllParagraphsOfText(XText xText) {
    if (xText == null) {
      throw new RuntimeException("XText == null"); 
    }
    int num = 0;
    XTextCursor xTextCursor = xText.createTextCursor();
    XParagraphCursor xParagraphCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
    if (xParagraphCursor != null) {
      xParagraphCursor.gotoStart(false);
      do {
        num++;
      } while (xParagraphCursor.gotoNextParagraph(false));
    }
    return num;
  }
  
  /** 
   * Returns all paragraphs of all text frames of a document
   * NOTE: Is currently not used 
   */
  public DocumentText getTextOfAllFrames(boolean withDeleted) {
    isBusy++;
    try {
      List<String> sText = new ArrayList<String>();
      Map<Integer, Integer> headingNumbers = new HashMap<Integer, Integer>();
      List<Integer> hNumbers = new ArrayList<Integer>();
      List<List<Integer>> deletedCharacters = new ArrayList<List<Integer>>();
      List<Integer> sortedTextIds = null;
      XTextFramesSupplier xTextFrameSupplier = UnoRuntime.queryInterface(XTextFramesSupplier.class, curDoc);
      XNameAccess xNamedFrames = xTextFrameSupplier.getTextFrames();
      for (String name : xNamedFrames.getElementNames()) {
        List<String> sTxt = new ArrayList<String>();
        List<List<Integer>> delCharacters = new ArrayList<List<Integer>>();
        Object o = xNamedFrames.getByName(name);
        if (o != null) {
          XText xFrameText = UnoRuntime.queryInterface(XText.class,  o);
          if (xFrameText != null) {
            addAllParagraphsOfText(xFrameText, sTxt, delCharacters, sortedTextIds, withDeleted);
            for (int i = 0; i < hNumbers.size(); i++) {
              hNumbers.set(i, hNumbers.get(i) + sTxt.size());
            }
            hNumbers.add(0, 0);
            sText.addAll(0, sTxt);
            deletedCharacters.addAll(0, delCharacters);
            for (int i = 0; i < hNumbers.size(); i++) {
              headingNumbers.put(hNumbers.get(i), 0);
            }
          }
        }
      }
      return new DocumentText(sText, headingNumbers, new ArrayList<Integer>(), sortedTextIds, deletedCharacters);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return null;           // Return null as method failed
    } finally {
      isBusy--;
    }
  }

  /** 
   * Returns the number of all paragraphs of all text frames of a document
   * NOTE: Is currently not used 
   */
  public int getNumberOfAllFrames() {
    isBusy++;
    try {
      int num = 0;
      if (curDoc != null) {
        XTextFramesSupplier xTextFrameSupplier = UnoRuntime.queryInterface(XTextFramesSupplier.class, curDoc);
        XNameAccess xNamedFrames = xTextFrameSupplier.getTextFrames();
        for (String name : xNamedFrames.getElementNames()) {
          Object o = xNamedFrames.getByName(name);
          if (o != null) {
            XText xFrameText = UnoRuntime.queryInterface(XText.class,  o);
            if (xFrameText != null) {
              num += getNumberOfAllParagraphsOfText(xFrameText);
            }
          }
        }
      }
      return num;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return 0;           // Return 0 as method failed
    } finally {
      isBusy--;
    }
  }

  /** 
   * Returns all paragraphs of all text shapes of a document
   */
  public DocumentText getTextOfAllShapes(boolean withDeleted) {
    isBusy++;
    try {
      List<String> sText = new ArrayList<String>();
      Map<Integer, Integer> headingNumbers = new HashMap<Integer, Integer>();
      List<Integer> sortedTextIds = null;
      List<List<Integer>> deletedCharacters = new ArrayList<List<Integer>>();
      XDrawPageSupplier xDrawPageSupplier = UnoRuntime.queryInterface(XDrawPageSupplier.class, curDoc);
      if (xDrawPageSupplier == null) {
        WtMessageHandler.printToLogFile("XDrawPageSupplier == null");
        return new DocumentText(sText, headingNumbers, new ArrayList<Integer>(), sortedTextIds, deletedCharacters);
      }
      XDrawPage xDrawPage = xDrawPageSupplier.getDrawPage();
      if (xDrawPage == null) {
        WtMessageHandler.printToLogFile("XDrawPage == null");
        return new DocumentText(sText, headingNumbers, new ArrayList<Integer>(), sortedTextIds, deletedCharacters);
      }
      XShapes xShapes = UnoRuntime.queryInterface(XShapes.class, xDrawPage);
      int nShapes = xShapes.getCount();
      for(int j = 0; j < nShapes; j++) {
        Object oShape = xShapes.getByIndex(j);
        XShape xShape = UnoRuntime.queryInterface(XShape.class, oShape);
        if (xShape != null) {
          XText xShapeText = UnoRuntime.queryInterface(XText.class, xShape);
          if (xShapeText != null) {
            sortedTextIds = addAllParagraphsOfText(xShapeText, sText, deletedCharacters, sortedTextIds, withDeleted);
            headingNumbers.put(sText.size(), 0);
          }
        }
      }
      return new DocumentText(sText, headingNumbers, new ArrayList<Integer>(), sortedTextIds, deletedCharacters);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return null;           // Return null as method failed
    } finally {
      isBusy--;
    }
  }
  
  /** 
   * Returns all paragraphs of all text shapes of a document
   */
  public List<String> getTextOfShapes(List<Integer> nPara) {
    isBusy++;
    try {
      List<String> sText = new ArrayList<String>();
      XDrawPageSupplier xDrawPageSupplier = UnoRuntime.queryInterface(XDrawPageSupplier.class, curDoc);
      if (xDrawPageSupplier == null) {
        WtMessageHandler.printToLogFile("XDrawPageSupplier == null");
        return sText;
      }
      XDrawPage xDrawPage = xDrawPageSupplier.getDrawPage();
      if (xDrawPage == null) {
        WtMessageHandler.printToLogFile("XDrawPage == null");
        return sText;
      }
      XShapes xShapes = UnoRuntime.queryInterface(XShapes.class, xDrawPage);
      int nShapes = xShapes.getCount();
      int num = 0;
      for(int j = 0; j < nShapes; j++) {
        Object oShape = xShapes.getByIndex(j);
        XShape xShape = UnoRuntime.queryInterface(XShape.class, oShape);
        if (xShape != null) {
          XText xShapeText = UnoRuntime.queryInterface(XText.class, xShape);
          if (xShapeText != null) {
            List<String> tmpText = getAllParagraphsOfText(xShapeText);
            for(String text : tmpText) {
              if (nPara.contains(num)) {
                sText.add(text);
              }
              num++;
            }
          }
        }
      }
      return sText;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return null;           // Return null as method failed
    } finally {
      isBusy--;
    }
  }
  
  /** 
   * Returns the number of all paragraphs of all text shapes of a document
   */
  public int getNumberOfAllShapes() {
    isBusy++;
    try {
      int num = 0;
      if (curDoc != null) {
        XDrawPageSupplier xDrawPageSupplier = UnoRuntime.queryInterface(XDrawPageSupplier.class, curDoc);
        if (xDrawPageSupplier == null) {
          WtMessageHandler.printToLogFile("XDrawPageSupplier == null");
          return 0;
        }
        XDrawPage xDrawPage = xDrawPageSupplier.getDrawPage();
        if (xDrawPage == null) {
          WtMessageHandler.printToLogFile("XDrawPage == null");
          return 0;
        }
        XShapes xShapes = UnoRuntime.queryInterface(XShapes.class, xDrawPage);
        int nShapes = xShapes.getCount();
        for(int j = 0; j < nShapes; j++) {
          Object oShape = xShapes.getByIndex(j);
          XShape xShape = UnoRuntime.queryInterface(XShape.class, oShape);
          if (xShape != null) {
            XText xShapeText = UnoRuntime.queryInterface(XText.class, xShape);
            if (xShapeText != null) {
              num += getNumberOfAllParagraphsOfText(xShapeText);
            }
          }
        }
      }
      return num;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return 0;           // Return 0 as method failed
    } finally {
      isBusy--;
    }
  }
  
  /** 
   * Returns the Index Access to all tables of a document
   */
  private XIndexAccess getIndexAccessOfAllTables() {
    try {
      if (curDoc == null) {
        return null;
      }
      // Get the TextTablesSupplier interface of the document
      XTextTablesSupplier xTableSupplier = UnoRuntime.queryInterface(XTextTablesSupplier.class, curDoc);
      // Get an XIndexAccess of TextTables
      if (xTableSupplier != null) {
        return UnoRuntime.queryInterface(XIndexAccess.class, xTableSupplier.getTextTables());
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
    }
    return null;           // Return null as method failed
  }
  
  /** 
   * Returns all paragraphs of all cells of all tables of a document
   */
  public DocumentText getTextOfAllTables(boolean withDeleted) {
    isBusy++;
    try {
      List<String> sText = new ArrayList<String>();
      Map<Integer, Integer> headingNumbers = new HashMap<Integer, Integer>();
      List<Integer> sortedTextIds = null;
      List<List<Integer>> deletedCharacters = new ArrayList<List<Integer>>();
      XIndexAccess xTables = getIndexAccessOfAllTables();
      if (xTables != null) {
        // Get all Tables of Document
        for (int i = 0; i < xTables.getCount(); i++) {
          XTextTable xTable = UnoRuntime.queryInterface(XTextTable.class, xTables.getByIndex(i));
          // Get all Cells of Tables
          if (xTable != null) {
            for (String cellName : xTable.getCellNames()) {
              XText xTableText = UnoRuntime.queryInterface(XText.class, xTable.getCellByName(cellName) );
              if (xTableText != null) {
                headingNumbers.put(sText.size(), 0);
                sortedTextIds = addAllParagraphsOfText(xTableText, sText, deletedCharacters, sortedTextIds, withDeleted);
              }
            }
          }
        }
      }
      return new DocumentText(sText, headingNumbers, new ArrayList<Integer>(), sortedTextIds, deletedCharacters);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return null;           // Return null as method failed
    } finally {
      isBusy--;
    }
  }
  
  /** 
   * Returns n paragraphs of tables
   * Note: nPara stores the numbers of textparagraphs
   */
  public List<String> getTextOfTables(List<Integer> nPara) {
    isBusy++;
    try {
      List<String> sText = new ArrayList<String>();
      XIndexAccess xTables = getIndexAccessOfAllTables();
      if (xTables != null) {
        // Get all Tables of Document
        int num = 0;
        for (int i = 0; i < xTables.getCount(); i++) {
          XTextTable xTable = UnoRuntime.queryInterface(XTextTable.class, xTables.getByIndex(i));
          // Get all Cells of Tables
          if (xTable != null) {
            for (String cellName : xTable.getCellNames()) {
              XText xTableText = UnoRuntime.queryInterface(XText.class, xTable.getCellByName(cellName) );
              List<String> tmpText = getAllParagraphsOfText(xTableText);
              for(String text : tmpText) {
                if (nPara.contains(num)) {
                  sText.add(text);
                }
                num++;
              }
            }
          }
        }
      }
      return sText;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return null;           // Return null as method failed
    } finally {
      isBusy--;
    }
  }
  
  /** 
   * Returns the number of paragraphs of all cells of all tables of a document
   */
  public int getNumberOfAllTables() {
    isBusy++;
    try {
      int num = 0;
      XIndexAccess xTables = getIndexAccessOfAllTables();
      if (xTables != null) {
        // Get all Tables of Document
        for (int i = 0; i < xTables.getCount(); i++) {
          XTextTable xTable = UnoRuntime.queryInterface(XTextTable.class, xTables.getByIndex(i));
          // Get all Cells of Tables
          if (xTable != null) {
            for (String cellName : xTable.getCellNames()) {
              XText xTableText = UnoRuntime.queryInterface(XText.class, xTable.getCellByName(cellName) );
              if (xTableText != null) {
                num += getNumberOfAllParagraphsOfText(xTableText);
              }
            }
          }
        }
      }
      return num;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return 0;           // Return 0 as method failed
    } finally {
      isBusy--;
    }
  }
  
  /** 
   * Returns all paragraphs of all footnotes of a document
   */
  public DocumentText getTextOfAllFootnotes(boolean withDeleted) {
    isBusy++;
    List<String> sText = new ArrayList<String>();
    Map<Integer, Integer> headingNumbers = new HashMap<Integer, Integer>();
    List<Integer> sortedTextIds = null;
    List<List<Integer>> deletedCharacters = new ArrayList<List<Integer>>();
    try {
      if (curDoc != null) {
        // Get the XFootnotesSupplier interface of the document
        XFootnotesSupplier xFootnoteSupplier = UnoRuntime.queryInterface(XFootnotesSupplier.class, curDoc );
        // Get an XIndexAccess of Footnotes
        if (xFootnoteSupplier != null) {
          XIndexAccess xFootnotes = UnoRuntime.queryInterface(XIndexAccess.class, xFootnoteSupplier.getFootnotes());
          if (xFootnotes != null) {
            for (int i = 0; i < xFootnotes.getCount(); i++) {
              XFootnote xFootnote = UnoRuntime.queryInterface(XFootnote.class, xFootnotes.getByIndex(i));
              XText xFootnoteText = UnoRuntime.queryInterface(XText.class, xFootnote);
              if (xFootnoteText != null) {
                headingNumbers.put(sText.size(), 0);
                sortedTextIds = addAllParagraphsOfText(xFootnoteText, sText, deletedCharacters, sortedTextIds, withDeleted);
              }
            }
          }
        }
      }
      return new DocumentText(sText, headingNumbers, new ArrayList<Integer>(), sortedTextIds, deletedCharacters);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return null;           // Return null as method failed
    } finally {
      isBusy--;
    }
  }
  
  /** 
   * Returns the number of paragraphs of all footnotes of a document
   */
  public int getNumberOfAllFootnotes() {
    isBusy++;
    try {
      int num = 0;
      if (curDoc != null) {
        // Get the XFootnotesSupplier interface of the document
        XFootnotesSupplier xFootnoteSupplier = UnoRuntime.queryInterface(XFootnotesSupplier.class, curDoc );
        // Get an XIndexAccess of Footnotes
        if (xFootnoteSupplier != null) {
          XIndexAccess xFootnotes = UnoRuntime.queryInterface(XIndexAccess.class, xFootnoteSupplier.getFootnotes());
          if (xFootnotes != null) {
            for (int i = 0; i < xFootnotes.getCount(); i++) {
              XFootnote xFootnote = UnoRuntime.queryInterface(XFootnote.class, xFootnotes.getByIndex(i));
              XText xFootnoteText = UnoRuntime.queryInterface(XText.class, xFootnote);
              if (xFootnoteText != null) {
                num += getNumberOfAllParagraphsOfText(xFootnoteText);
              }
            }
          }
        }
      }
      return num;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return 0;           // Return 0 as method failed
    } finally {
      isBusy--;
    }
  }
  
  /** 
   * Returns all paragraphs of all endnotes of a document
   */
  public DocumentText getTextOfAllEndnotes(boolean withDeleted) {
    isBusy++;
    List<String> sText = new ArrayList<String>();
    Map<Integer, Integer> headingNumbers = new HashMap<Integer, Integer>();
    List<Integer> sortedTextIds = null;
    List<List<Integer>> deletedCharacters = new ArrayList<List<Integer>>();
    try {
      if (curDoc != null) {
        // Get the XEndnotesSupplier interface of the document
        XEndnotesSupplier xEndnoteSupplier = UnoRuntime.queryInterface(XEndnotesSupplier.class, curDoc );
        // Get an XIndexAccess of Endnotes
        if (xEndnoteSupplier != null) {
          XIndexAccess xEndnotes = UnoRuntime.queryInterface(XIndexAccess.class, xEndnoteSupplier.getEndnotes());
          if (xEndnotes != null) {
            for (int i = 0; i < xEndnotes.getCount(); i++) {
              XFootnote xEndnote = UnoRuntime.queryInterface(XFootnote.class, xEndnotes.getByIndex(i));
              XText xFootnoteText = UnoRuntime.queryInterface(XText.class, xEndnote);
              if (xFootnoteText != null) {
                headingNumbers.put(sText.size(), 0);
                sortedTextIds = addAllParagraphsOfText(xFootnoteText, sText, deletedCharacters, sortedTextIds, withDeleted);
              }
            }
          }
        }
      }
      return new DocumentText(sText, headingNumbers, new ArrayList<Integer>(), sortedTextIds, deletedCharacters);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return null;           // Return null as method failed
    } finally {
      isBusy--;
    }
  }
  
  /** 
   * Returns number of paragraphs of all endnotes of a document
   */
  public int getNumberOfAllEndnotes() {
    isBusy++;
    try {
      int num = 0;
      if (curDoc != null) {
        // Get the XEndnotesSupplier interface of the document
        XEndnotesSupplier xEndnoteSupplier = UnoRuntime.queryInterface(XEndnotesSupplier.class, curDoc );
        // Get an XIndexAccess of Endnotes
        if (xEndnoteSupplier != null) {
          XIndexAccess xEndnotes = UnoRuntime.queryInterface(XIndexAccess.class, xEndnoteSupplier.getEndnotes());
          if (xEndnotes != null) {
            for (int i = 0; i < xEndnotes.getCount(); i++) {
              XFootnote xEndnote = UnoRuntime.queryInterface(XFootnote.class, xEndnotes.getByIndex(i));
              XText xEndnoteText = UnoRuntime.queryInterface(XText.class, xEndnote);
              if (xEndnoteText != null) {
                num += getNumberOfAllParagraphsOfText(xEndnoteText);
              }
            }
          }
        }
      }
      return num;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return 0;           // Return 0 as method failed
    } finally {
      isBusy--;
    }
  }
  
  /** 
   * Returns the page property sets of of a document
   */
  private List<XPropertySet> getPagePropertySets() {
    try {
      List<XPropertySet> propertySets = new ArrayList<XPropertySet>();
      if (curDoc == null) {
        return null;
      }
      XStyleFamiliesSupplier xSupplier =  UnoRuntime.queryInterface(XStyleFamiliesSupplier.class, curDoc);
      if (xSupplier == null) {
        return null;
      }
      XNameAccess xNameAccess = xSupplier.getStyleFamilies();
      if (xNameAccess == null) {
        return null;
      }
      XNameContainer pageStyleCon = UnoRuntime.queryInterface(XNameContainer.class, xNameAccess.getByName("PageStyles"));
      if (pageStyleCon == null) {
        return null;
      }
      for (String name : pageStyleCon.getElementNames()) {
        XPropertySet xPageStandardProps = UnoRuntime.queryInterface(XPropertySet.class, pageStyleCon.getByName(name));
        if (xPageStandardProps != null) {
          propertySets.add(UnoRuntime.queryInterface(XPropertySet.class, xPageStandardProps));
        }
      }
      return propertySets;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return null;           // Return null as method failed
    }
  }
  
  /** 
   * Returns all paragraphs of headers and footers of a document
   */
  public DocumentText getTextOfAllHeadersAndFooters(boolean withDeleted) {
    isBusy++;
    try {
      List<String> sText = new ArrayList<String>();
      Map<Integer, Integer> headingNumbers = new HashMap<Integer, Integer>();
      List<Integer> sortedTextIds = null;
      List<List<Integer>> deletedCharacters = new ArrayList<List<Integer>>();
      List<XPropertySet> xPagePropertySets = getPagePropertySets();
      if (xPagePropertySets != null) {
        for (XPropertySet xPagePropertySet : xPagePropertySets) {
          if (xPagePropertySet != null) {
            boolean headerIsOn = WtOfficeTools.getBooleanValue(xPagePropertySet.getPropertyValue("HeaderIsOn"));
            boolean firstIsShared = WtOfficeTools.getBooleanValue(xPagePropertySet.getPropertyValue("FirstIsShared"));
            boolean headerIsShared = WtOfficeTools.getBooleanValue(xPagePropertySet.getPropertyValue("HeaderIsShared"));
            boolean footerIsOn = WtOfficeTools.getBooleanValue(xPagePropertySet.getPropertyValue("FooterIsOn"));
            boolean footerIsShared = WtOfficeTools.getBooleanValue(xPagePropertySet.getPropertyValue("FooterIsShared"));
            for (int i = 0; i < HeaderFooterTypes.length; i++) {
              if ((headerIsOn && ((i == 0 && headerIsShared) 
                  || ((i == 1 || i == 2) && !headerIsShared)
                  || (i == 3 && !firstIsShared)))
                  || (footerIsOn && ((i == 4 && footerIsShared) 
                      || ((i == 5 || i == 6) && !footerIsShared)
                      || (i == 7 && !firstIsShared)))) {
                XText xHeaderText = UnoRuntime.queryInterface(XText.class, xPagePropertySet.getPropertyValue(HeaderFooterTypes[i]));
                if (xHeaderText != null && !xHeaderText.getString().isEmpty()) {
                  if (!headingNumbers.containsKey(sText.size())) {
                    headingNumbers.put(sText.size(), 0);
                  }
                  sortedTextIds = addAllParagraphsOfText(xHeaderText, sText, deletedCharacters, sortedTextIds, withDeleted);
                }
              }
            }
          }
        }
      }
      return new DocumentText(sText, headingNumbers, new ArrayList<Integer>(), sortedTextIds, deletedCharacters);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return null;           // Return null as method failed
    } finally {
      isBusy--;
    }
  }

  /** 
   * Returns the number of paragraphs of headers and footers of a document
   */
  public int getNumberOfAllHeadersAndFooters() {
    isBusy++;
    try {
      int num = 0;
      List<XPropertySet> xPagePropertySets = getPagePropertySets();
      if (xPagePropertySets != null) {
        for (XPropertySet xPagePropertySet : xPagePropertySets) {
          if (xPagePropertySet != null) {
            boolean headerIsOn = WtOfficeTools.getBooleanValue(xPagePropertySet.getPropertyValue("HeaderIsOn"));
            boolean firstIsShared = WtOfficeTools.getBooleanValue(xPagePropertySet.getPropertyValue("FirstIsShared"));
            boolean headerIsShared = WtOfficeTools.getBooleanValue(xPagePropertySet.getPropertyValue("HeaderIsShared"));
            boolean footerIsOn = WtOfficeTools.getBooleanValue(xPagePropertySet.getPropertyValue("FooterIsOn"));
            boolean footerIsShared = WtOfficeTools.getBooleanValue(xPagePropertySet.getPropertyValue("FooterIsShared"));
            for (int i = 0; i < HeaderFooterTypes.length; i++) {
              if ((headerIsOn && ((i == 0 && headerIsShared) 
                  || ((i == 1 || i == 2) && !headerIsShared)
                  || (i == 3 && !firstIsShared)))
                  || (footerIsOn && ((i == 4 && footerIsShared) 
                      || ((i == 5 || i == 6) && !footerIsShared)
                      || (i == 7 && !firstIsShared)))) {
                XText xHeaderText = UnoRuntime.queryInterface(XText.class, xPagePropertySet.getPropertyValue(HeaderFooterTypes[i]));
                if (xHeaderText != null && !xHeaderText.getString().isEmpty()) {
                  num += getNumberOfAllParagraphsOfText(xHeaderText);
                }
              }
            }
          }
        }
      }
      return num;
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return 0;           // Return 0 as method failed
    } finally {
      isBusy--;
    }
  }

  /** 
   * get the paragraph cursor
   */
  public XParagraphCursor getParagraphCursor(TextParagraph textPara) {
    isBusy++;
    try {
      if (textPara == null) {
        return null;
      }
      int type = textPara.type;
      int number = textPara.number;
      int nPara = 0;
      if (type == WtDocumentCache.CURSOR_TYPE_UNKNOWN) {
        return null;
      } else if (type == WtDocumentCache.CURSOR_TYPE_TEXT) {
        if (xPCursor == null) {
          return null;
        }
        xPCursor.gotoStart(false);
        while (nPara < number && xPCursor.gotoNextParagraph(false)) nPara++;
        return xPCursor;
      } else if (type == WtDocumentCache.CURSOR_TYPE_TABLE) {
        XTextTablesSupplier xTableSupplier = UnoRuntime.queryInterface(XTextTablesSupplier.class, curDoc);
        if (xTableSupplier != null) {
          XIndexAccess xTables = UnoRuntime.queryInterface(XIndexAccess.class, xTableSupplier.getTextTables());
          if (xTables != null) {
            for (int i = 0; i < xTables.getCount(); i++) {
              XTextTable xTable = UnoRuntime.queryInterface(XTextTable.class, xTables.getByIndex(i));
              if (xTable != null) {
                for (String cellName : xTable.getCellNames()) {
                  XText xTableText = UnoRuntime.queryInterface(XText.class, xTable.getCellByName(cellName) );
                  if (xTableText != null) {
                    XTextCursor xTextCursor = xTableText.createTextCursor();
                    XParagraphCursor xParagraphCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
                    if (xTableText != null) {
                      xParagraphCursor.gotoStart(false);
                      while (nPara < number && xParagraphCursor.gotoNextParagraph(false)){
                        nPara++;
                      }
                      if (nPara == number) {
                        return xParagraphCursor;
                      }
                      nPara++;
                    }
                  }
                }
              }
            }
          }
        }
      } else if (type == WtDocumentCache.CURSOR_TYPE_FOOTNOTE) {
        XFootnotesSupplier xFootnoteSupplier = UnoRuntime.queryInterface(XFootnotesSupplier.class, curDoc );
        if (xFootnoteSupplier != null) {
          XIndexAccess xFootnotes = UnoRuntime.queryInterface(XIndexAccess.class, xFootnoteSupplier.getFootnotes());
          if (xFootnotes != null) {
            for (int i = 0; i < xFootnotes.getCount(); i++) {
              XFootnote XFootnote = UnoRuntime.queryInterface(XFootnote.class, xFootnotes.getByIndex(i));
              XText xFootnoteText = UnoRuntime.queryInterface(XText.class, XFootnote);
              if (xFootnoteText != null) {
                XTextCursor xTextCursor = xFootnoteText.createTextCursor();
                XParagraphCursor xParagraphCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
                if (xParagraphCursor != null) {
                  xParagraphCursor.gotoStart(false);
                  while (nPara < number && xParagraphCursor.gotoNextParagraph(false)){
                    nPara++;
                  }
                  if (nPara == number) {
                    return xParagraphCursor;
                  }
                  nPara++;
                }
              }
            }
          }
        }
      } else if (type == WtDocumentCache.CURSOR_TYPE_ENDNOTE) {
        XEndnotesSupplier xEndnotesSupplier = UnoRuntime.queryInterface(XEndnotesSupplier.class, curDoc );
        if (xEndnotesSupplier != null) {
          XIndexAccess xEndnotes = UnoRuntime.queryInterface(XIndexAccess.class, xEndnotesSupplier.getEndnotes());
          if (xEndnotes != null) {
            for (int i = 0; i < xEndnotes.getCount(); i++) {
              XFootnote xEndnote = UnoRuntime.queryInterface(XFootnote.class, xEndnotes.getByIndex(i));
              XText xEndnoteText = UnoRuntime.queryInterface(XText.class, xEndnote);
              if (xEndnoteText != null) {
                XTextCursor xTextCursor = xEndnoteText.createTextCursor();
                XParagraphCursor xParagraphCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
                if (xParagraphCursor != null) {
                  xParagraphCursor.gotoStart(false);
                  while (nPara < number && xParagraphCursor.gotoNextParagraph(false)){
                    nPara++;
                  }
                  if (nPara == number) {
                    return xParagraphCursor;
                  }
                  nPara++;
                }
              }
            }
          }
        }
      } else if (type == WtDocumentCache.CURSOR_TYPE_HEADER_FOOTER) {
        List<XPropertySet> xPagePropertySets = getPagePropertySets();
        if (xPagePropertySets != null) {
          XText lastHeaderText = null;
          for (XPropertySet xPagePropertySet : xPagePropertySets) {
            if (xPagePropertySet != null) {
              for (String headerFooter : WtDocumentCursorTools.HeaderFooterTypes) {
                XText xHeaderText = UnoRuntime.queryInterface(XText.class, xPagePropertySet.getPropertyValue(headerFooter));
                if (xHeaderText != null && !xHeaderText.getString().isEmpty() && (lastHeaderText == null || !lastHeaderText.equals(xHeaderText))) {
                  XTextCursor xTextCursor = xHeaderText.createTextCursor();
                  XParagraphCursor xParagraphCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
                  if (xParagraphCursor != null) {
                    xParagraphCursor.gotoStart(false);
                    while (nPara < number && xParagraphCursor.gotoNextParagraph(false)){
                      nPara++;
                    }
                    if (nPara == number) {
                      return xParagraphCursor;
                    }
                    nPara++;
                    lastHeaderText = xHeaderText;
                  }
                }
              }
            }
          }
        }
      } else if (type == WtDocumentCache.CURSOR_TYPE_SHAPE) {
        XDrawPageSupplier xDrawPageSupplier = UnoRuntime.queryInterface(XDrawPageSupplier.class, curDoc);
        if (xDrawPageSupplier != null) {
          XDrawPage xDrawPage = xDrawPageSupplier.getDrawPage();
          if (xDrawPage != null) {
            XShapes xShapes = UnoRuntime.queryInterface(XShapes.class, xDrawPage);
            if (xShapes != null) {
              int nShapes = xShapes.getCount();
              for(int j = 0; j < nShapes; j++) {
                Object oShape = xShapes.getByIndex(j);
                XShape xShape = UnoRuntime.queryInterface(XShape.class, oShape);
                if (xShape != null) {
                  XText xShapeText = UnoRuntime.queryInterface(XText.class, xShape);
                  if (xShapeText != null) {
                    XTextCursor xTextCursor = xShapeText.createTextCursor();
                    XParagraphCursor xParagraphCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
                    if (xParagraphCursor != null) {
                      xParagraphCursor.gotoStart(false);
                      while (nPara < number && xParagraphCursor.gotoNextParagraph(false)){
                        nPara++;
                      }
                      if (nPara == number) {
                        return xParagraphCursor;
                      }
                      nPara++;
                    }
                  }
                }
              }
            }
          }
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
    } finally {
      isBusy--;
    }
    return null;
  }
  
  /** 
   * remove marks of text paragraph
   */
  public void removeMarks(List<TextParagraph> textParas) throws Throwable {
    isBusy++;
    try {
      if (textParas == null) {
        return;
      }
      List<List<Integer>> paras = new ArrayList<List<Integer>>();
      for (int n = 0; n < WtDocumentCache.NUMBER_CURSOR_TYPES; n++) {
        paras.add(new ArrayList<Integer>());
      }
      for (int i = 0; i < textParas.size(); i++) {
        int type = textParas.get(i).type;
        if (type >= 0 && type < WtDocumentCache.NUMBER_CURSOR_TYPES) {
          paras.get(type).add(textParas.get(i).number);
        }
      }
      for (int n = 0; n < WtDocumentCache.NUMBER_CURSOR_TYPES; n++) {
        if (!paras.get(n).isEmpty()) {
          paras.get(n).sort(null);
        }
      }
      for (int type = 0; type < WtDocumentCache.NUMBER_CURSOR_TYPES; type++) {
        if (paras.get(type).size() > 0) {
          int nPara = 0;
          if (type == WtDocumentCache.CURSOR_TYPE_TEXT) {
            if (xPCursor == null) {
              break;
            }
            xPCursor.gotoStart(false);
            for (int i = 0; i < paras.get(type).size(); i++) {
              int number = paras.get(type).get(i);
              while (nPara < number && xPCursor.gotoNextParagraph(false)) {
                nPara++;
              }
              if (xPCursor != null) {
                XMarkingAccess xMarkingAccess = UnoRuntime.queryInterface(XMarkingAccess.class, xPCursor);
                if (xMarkingAccess == null) {
                  WtMessageHandler.printToLogFile("FlatParagraphTools: addMarksToOneParagraph: xMarkingAccess == null");
                } else {
                  xMarkingAccess.invalidateMarkings(TextMarkupType.PROOFREADING);
                }
              }
            }
          } else if (type == WtDocumentCache.CURSOR_TYPE_TABLE) {
            XTextTablesSupplier xTableSupplier = UnoRuntime.queryInterface(XTextTablesSupplier.class, curDoc);
            if (xTableSupplier != null) {
              XIndexAccess xTables = UnoRuntime.queryInterface(XIndexAccess.class, xTableSupplier.getTextTables());
              if (xTables != null) {
                int j = 0;
                int number = paras.get(type).get(j);
                for (int i = 0; i < xTables.getCount() && j < paras.get(type).size(); i++) {
                  XTextTable xTable = UnoRuntime.queryInterface(XTextTable.class, xTables.getByIndex(i));
                  if (xTableSupplier != null) {
                    for (String cellName : xTable.getCellNames()) {
                      XText xTableText = UnoRuntime.queryInterface(XText.class, xTable.getCellByName(cellName) );
                      if (xTableText != null) {
                        XTextCursor xTextCursor = xTableText.createTextCursor();
                        XParagraphCursor xParagraphCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
                        if (xParagraphCursor != null) {
                          xParagraphCursor.gotoStart(false);
                          do {
                            if (nPara == number) {
                              XMarkingAccess xMarkingAccess = UnoRuntime.queryInterface(XMarkingAccess.class, xParagraphCursor);
                              if (xMarkingAccess == null) {
                                WtMessageHandler.printToLogFile("FlatParagraphTools: addMarksToOneParagraph: xMarkingAccess == null");
                              } else {
                                xMarkingAccess.invalidateMarkings(TextMarkupType.PROOFREADING);
                              }
                              j++;
                              if (j < paras.get(type).size()) {
                                number = paras.get(type).get(j);
                              }
                            }
                            nPara++;
                          } while (j < paras.get(type).size() && xParagraphCursor.gotoNextParagraph(false));
                          if (j == paras.get(type).size()) {
                            break;
                          }
                        }
                      }
                    }
                  }
                }
              }
            }
          } else if (type == WtDocumentCache.CURSOR_TYPE_FOOTNOTE) {
            XFootnotesSupplier xFootnoteSupplier = UnoRuntime.queryInterface(XFootnotesSupplier.class, curDoc );
            if (xFootnoteSupplier != null) {
              XIndexAccess xFootnotes = UnoRuntime.queryInterface(XIndexAccess.class, xFootnoteSupplier.getFootnotes());
              if (xFootnotes != null) {
                int j = 0;
                int number = paras.get(type).get(j);
                for (int i = 0; i < xFootnotes.getCount() && j < paras.get(type).size(); i++) {
                  XFootnote XFootnote = UnoRuntime.queryInterface(XFootnote.class, xFootnotes.getByIndex(i));
                  XText xFootnoteText = UnoRuntime.queryInterface(XText.class, XFootnote);
                  if (xFootnoteText != null) {
                    XTextCursor xTextCursor = xFootnoteText.createTextCursor();
                    XParagraphCursor xParagraphCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
                    if (xParagraphCursor != null) {
                      xParagraphCursor.gotoStart(false);
                      do {
                        if (nPara == number) {
                          XMarkingAccess xMarkingAccess = UnoRuntime.queryInterface(XMarkingAccess.class, xParagraphCursor);
                          if (xMarkingAccess == null) {
                            WtMessageHandler.printToLogFile("FlatParagraphTools: addMarksToOneParagraph: xMarkingAccess == null");
                          } else {
                            xMarkingAccess.invalidateMarkings(TextMarkupType.PROOFREADING);
                          }
                          j++;
                          if (j < paras.get(type).size()) {
                            number = paras.get(type).get(j);
                          }
                        }
                        nPara++;
                      } while (j < paras.get(type).size() && xParagraphCursor.gotoNextParagraph(false));
                    }
                  }
                }
              }
            }
          } else if (type == WtDocumentCache.CURSOR_TYPE_ENDNOTE) {
            XEndnotesSupplier xEndnotesSupplier = UnoRuntime.queryInterface(XEndnotesSupplier.class, curDoc );
            if (xEndnotesSupplier != null) {
              XIndexAccess xEndnotes = UnoRuntime.queryInterface(XIndexAccess.class, xEndnotesSupplier.getEndnotes());
              if (xEndnotes != null) {
                int j = 0;
                int number = paras.get(type).get(j);
                for (int i = 0; i < xEndnotes.getCount() && j < paras.get(type).size(); i++) {
                  XFootnote xEndnote = UnoRuntime.queryInterface(XFootnote.class, xEndnotes.getByIndex(i));
                  XText xEndnoteText = UnoRuntime.queryInterface(XText.class, xEndnote);
                  if (xEndnoteText != null) {
                    XTextCursor xTextCursor = xEndnoteText.createTextCursor();
                    XParagraphCursor xParagraphCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
                    if (xParagraphCursor != null) {
                      xParagraphCursor.gotoStart(false);
                      do {
                        if (nPara == number) {
                          XMarkingAccess xMarkingAccess = UnoRuntime.queryInterface(XMarkingAccess.class, xParagraphCursor);
                          if (xMarkingAccess == null) {
                            WtMessageHandler.printToLogFile("FlatParagraphTools: addMarksToOneParagraph: xMarkingAccess == null");
                          } else {
                            xMarkingAccess.invalidateMarkings(TextMarkupType.PROOFREADING);
                          }
                          j++;
                          if (j < paras.get(type).size()) {
                            number = paras.get(type).get(j);
                          }
                        }
                        nPara++;
                      } while (j < paras.get(type).size() && xParagraphCursor.gotoNextParagraph(false));
                    }
                  }
                }
              }
            }
          } else if (type == WtDocumentCache.CURSOR_TYPE_HEADER_FOOTER) {
            List<XPropertySet> xPagePropertySets = getPagePropertySets();
            if (xPagePropertySets != null) {
              XText lastHeaderText = null;
              int j = 0;
              int number = paras.get(type).get(j);
              for (XPropertySet xPagePropertySet : xPagePropertySets) {
                if (xPagePropertySet != null) {
                  for (String headerFooter : WtDocumentCursorTools.HeaderFooterTypes) {
                    XText xHeaderText = UnoRuntime.queryInterface(XText.class, xPagePropertySet.getPropertyValue(headerFooter));
                    if (xHeaderText != null && !xHeaderText.getString().isEmpty() && (lastHeaderText == null || !lastHeaderText.equals(xHeaderText))) {
                      XTextCursor xTextCursor = xHeaderText.createTextCursor();
                      XParagraphCursor xParagraphCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
                      if (xParagraphCursor != null) {
                        xParagraphCursor.gotoStart(false);
                        do {
                          if (nPara == number) {
                            XMarkingAccess xMarkingAccess = UnoRuntime.queryInterface(XMarkingAccess.class, xParagraphCursor);
                            if (xMarkingAccess == null) {
                              WtMessageHandler.printToLogFile("FlatParagraphTools: addMarksToOneParagraph: xMarkingAccess == null");
                            } else {
                              xMarkingAccess.invalidateMarkings(TextMarkupType.PROOFREADING);
                            }
                            j++;
                            if (j < paras.get(type).size()) {
                              number = paras.get(type).get(j);
                            }
                          }
                          nPara++;
                        } while (j < paras.get(type).size() && xParagraphCursor.gotoNextParagraph(false));
                        lastHeaderText = xHeaderText;
                      }
                      if (j == paras.get(type).size()) {
                        break;
                      }
                    }
                  }
                }
                if (j == paras.get(type).size()) {
                  break;
                }
              }
            }
          } else if (type == WtDocumentCache.CURSOR_TYPE_SHAPE) {
            XDrawPageSupplier xDrawPageSupplier = UnoRuntime.queryInterface(XDrawPageSupplier.class, curDoc);
            if (xDrawPageSupplier != null) {
              XDrawPage xDrawPage = xDrawPageSupplier.getDrawPage();
              if (xDrawPage != null) {
                XShapes xShapes = UnoRuntime.queryInterface(XShapes.class, xDrawPage);
                if (xShapes != null) {
                  int nShapes = xShapes.getCount();
                  int j = 0;
                  int number = paras.get(type).get(j);
                  for(int i = 0; i < nShapes && j < paras.get(type).size(); i++) {
                    Object oShape = xShapes.getByIndex(i);
                    XShape xShape = UnoRuntime.queryInterface(XShape.class, oShape);
                    if (xShape != null) {
                      XText xShapeText = UnoRuntime.queryInterface(XText.class, xShape);
                      if (xShapeText != null) {
                        XTextCursor xTextCursor = xShapeText.createTextCursor();
                        XParagraphCursor xParagraphCursor = UnoRuntime.queryInterface(XParagraphCursor.class, xTextCursor);
                        if (xParagraphCursor != null) {
                          xParagraphCursor.gotoStart(false);
                          do {
                            if (nPara == number) {
                              XMarkingAccess xMarkingAccess = UnoRuntime.queryInterface(XMarkingAccess.class, xParagraphCursor);
                              if (xMarkingAccess == null) {
                                WtMessageHandler.printToLogFile("FlatParagraphTools: addMarksToOneParagraph: xMarkingAccess == null");
                              } else {
                                xMarkingAccess.invalidateMarkings(TextMarkupType.PROOFREADING);
                              }
                              j++;
                              if (j < paras.get(type).size()) {
                                number = paras.get(type).get(j);
                              }
                            }
                            nPara++;
                          } while (j < paras.get(type).size() && xParagraphCursor.gotoNextParagraph(false));
                        }
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
    } finally {
      isBusy--;
    }
  }
  
  /**
   * get positions of deleted characters from a text paragraph
   */
  public List<Integer> getDeletedCharactersOfTextParagraph(TextParagraph textPara, boolean withDeleted) {
    isBusy++;
    try {
      if (!withDeleted || textPara == null) {
        return new ArrayList<Integer>();
      }
      XParagraphCursor xPCursor = getParagraphCursor(textPara);
      if (xPCursor == null) {
        return null;
      }
      xPCursor.gotoStartOfParagraph(false);
      xPCursor.gotoEndOfParagraph(true);
      return getDeletedCharacters(xPCursor, withDeleted);
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
      return null;
    } finally {
      isBusy--;
    }
  }
  
  /**
   * is protected character at position in paragraph
   */
  public boolean isProtectedCharacter(TextParagraph textPara, short x) {
    isBusy++;
    try {
      if (textPara == null) {
        return false;
      }
      XParagraphCursor xPCursor = getParagraphCursor(textPara);
      if (xPCursor == null) {
        return false;
      }
      xPCursor.gotoStartOfParagraph(false);
      xPCursor.goRight(x, false);
      XPropertySet xParagraphPropertySet = UnoRuntime.queryInterface(XPropertySet.class, xPCursor.getStart());
      if (xParagraphPropertySet != null) {
        XTextSection xTextSection = UnoRuntime.queryInterface(XTextSection.class, xParagraphPropertySet.getPropertyValue("TextSection"));
        xParagraphPropertySet = UnoRuntime.queryInterface(XPropertySet.class, xTextSection);
        if(xParagraphPropertySet != null && (boolean) xParagraphPropertySet.getPropertyValue("IsProtected")) {
          return true;
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught and printed to log file
    } finally {
      isBusy--;
    }
    return false;
  }
  
  /**
   *  Returns the status of cursor tools
   *  true: If a cursor tool in one or more threads is active
   */
  public static boolean isBusy() {
    return isBusy > 0;
  }
  
  /**
   *  Reset the busy flag
   */
  public static void reset() {
    isBusy = 0;
  }
  
  /**
   * Class to give back the text and the headings under the specified cursor
   */
  public static class DocumentText {
    public List<String> paragraphs;
    public Map<Integer, Integer> headingNumbers;
    public List<Integer> automaticTextParagraphs;
    public List<Integer> sortedTextIds;
    public List<List<Integer>> deletedCharacters;
    
    public DocumentText() {
      this.paragraphs = new ArrayList<String>();
      this.headingNumbers = new HashMap<Integer, Integer>();
      this.automaticTextParagraphs = new ArrayList<Integer>();
      this.sortedTextIds = null;
      this.deletedCharacters = new ArrayList<List<Integer>>();
    }
    
    DocumentText(List<String> paragraphs, Map<Integer, Integer> headingNumbers, List<Integer> automaticTextParagraphs, 
        List<Integer> sortedTextIds, List<List<Integer>> deletedCharacters) {
      this.paragraphs = paragraphs;
      this.headingNumbers = headingNumbers;
      this.automaticTextParagraphs = automaticTextParagraphs;
      this.sortedTextIds = sortedTextIds;
      this.deletedCharacters = deletedCharacters;
    }
  }
  
}
  

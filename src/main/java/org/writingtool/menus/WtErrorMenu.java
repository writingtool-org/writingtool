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
package org.writingtool.menus;

import java.util.ResourceBundle;

import org.jetbrains.annotations.Nullable;
import org.writingtool.WtDocumentCache;
import org.writingtool.WtDocumentsHandler;
import org.writingtool.WtLanguageTool;
import org.writingtool.WtProofreadingError;
import org.writingtool.WtSingleDocument;
import org.writingtool.tools.WtDocumentCursorTools;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;

import com.sun.star.accessibility.XAccessible;
import com.sun.star.accessibility.XAccessibleComponent;
import com.sun.star.accessibility.XAccessibleContext;
import com.sun.star.accessibility.XAccessibleText;
import com.sun.star.awt.MenuEvent;
import com.sun.star.awt.Point;
import com.sun.star.awt.Rectangle;
import com.sun.star.awt.XMenuListener;
import com.sun.star.awt.XPopupMenu;
import com.sun.star.awt.XVclWindowPeer;
import com.sun.star.lang.XComponent;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 * Class for a popup-menu under a text cursor position
 * showing the grammar error on this position
 * @since 24.7
 * @author Fred Kruse
 */
public class WtErrorMenu {
  
  private static final ResourceBundle MESSAGES = WtOfficeTools.getMessageBundle();

  private final XComponentContext xContext;
  private final WtSingleDocument document;

  public WtErrorMenu(WtSingleDocument document, XComponentContext xContext) {
    this.xContext = xContext;
    this.document = document;
  }
  
  /** 
   * Returns the accessible text at caret-position
   * Returns null if it fails
   */
  private static XAccessibleText findAccessibleText(XAccessibleContext xCtx) {
    if (xCtx == null) {
      return null;
    }
    XAccessibleText xText = UnoRuntime.queryInterface(XAccessibleText.class, xCtx);
    if (xText != null) {
      int caretPos = xText.getCaretPosition();
      if (caretPos >= 0) {
        return xText;
      }
    }
    // search children recursively
    long count = xCtx.getAccessibleChildCount();
    for (long i = 0; i < count; i++) {
      try {
        XAccessible xChild = xCtx.getAccessibleChild(i);
        if (xChild == null) {
          continue;
        }
        XAccessibleContext xChildCtx = xChild.getAccessibleContext();
        XAccessibleText result = findAccessibleText(xChildCtx);
        if (result != null) {
          return result;
        }
      } catch (Throwable t) {
        WtMessageHandler.showError(t);
      }
    }
    return null;
  }

  /** 
   * Returns the Caret-Position as Point
   * Returns null if it fails
   */
  @Nullable
  private Point getCaretPosition(XComponent xComponent) {
    try {
      XVclWindowPeer xPeer = WtOfficeTools.getVclWindowPeer(xComponent);
      // Get Accessible-Root of Window
      XAccessible xRootAcc = UnoRuntime.queryInterface(XAccessible.class, xPeer.getProperty("XAccessible"));
      // Get AccessibleContext of Window
      XAccessibleContext xRootCtx = xRootAcc.getAccessibleContext();
      // Find Text-Document-Pane and localize the Cursor-Caret
      XAccessibleText xAccText = findAccessibleText(xRootCtx);
      if (xAccText != null) {
        int caretPos = xAccText.getCaretPosition();
        if (caretPos >= 0) {
          com.sun.star.awt.Rectangle bounds = xAccText.getCharacterBounds(caretPos);
          XAccessibleComponent xComp = UnoRuntime.queryInterface(XAccessibleComponent.class, xRootAcc);
          Point origin = xComp.getLocationOnScreen();
          xComp = UnoRuntime.queryInterface(XAccessibleComponent.class, xAccText);
          Point textOrigin = xComp.getLocationOnScreen();
//          WtMessageHandler.printToLogFile("WtOfficeTools: getViewCursorPosition: origin (X/Y): (" + origin.X + " / " + origin.Y + ")");
//          WtMessageHandler.printToLogFile("WtOfficeTools: getViewCursorPosition: bounds (X/Y): (" + bounds.X + " / " + bounds.Y 
//              + "), Height: " + bounds.Height);
          Point point = new Point();
//          WtMessageHandler.printToLogFile("findAccessibleText: xText != null, caretPos: " + caretPos);
          point.X = textOrigin.X - origin.X + bounds.X;
          point.Y = textOrigin.Y - origin.Y + bounds.Y + bounds.Height;
          return point;
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
    return null;
  }

  public void openPopupMenu(int nFPara, XComponent xComponent, WtProofreadingError error) {
    try {
      XPopupMenu xPopup = WtOfficeTools.getPopupMenu(xContext);
      // insert menu items
      short menuId = 1;
      short posId = 0;
      xPopup.insertItem(menuId, error.aShortComment, (short) 0, posId);
      menuId++;
      posId++;
      xPopup.insertSeparator(posId);
      posId++;
      int numSuggestions = error.aSuggestions.length;
      for (int i = 0; i < numSuggestions; i++) {
        xPopup.insertItem(menuId, error.aSuggestions[i],  (short) 0, posId);
        menuId++;
        posId++;
      }
      xPopup.insertSeparator(posId);
      posId++;
      xPopup.insertItem(menuId, MESSAGES.getString("checkDialogIgnoreButton"),  (short) 0, posId);
      int ignoreId = menuId;
      menuId++;
      posId++;
      xPopup.insertItem(menuId, MESSAGES.getString("loContextMenuIgnorePermanent"),  (short) 0, posId);
      int ignorePermId = menuId;

      xPopup.setDefaultItem((short) 2);
//      xPopup.enableItem((short) 1, false);
//      xPopup.hideDisabledEntries(false);
  
      // Listener for selection
      xPopup.addMenuListener(new XMenuListener() {
        public void itemSelected(MenuEvent e) {
          WtMessageHandler.printToLogFile("MenuId: " + e.MenuId);
          if (e.MenuId == ignoreId) {
            document.ignoreOnce();
          } else if (e.MenuId == ignorePermId) { 
            document.ignorePermanent();
          } else if (e.MenuId > 1) {
            try {
              WtMessageHandler.printToLogFile("MenuId > 1: " + e.MenuId);
              changeTextOfParagraph(nFPara, error.nErrorStart, error.nErrorLength, error.aSuggestions[e.MenuId - 2]);
            } catch (Throwable t) {
              WtMessageHandler.showError(t);
            }
          }
        }
        public void itemActivated(MenuEvent e) {}
        public void itemDeactivated(MenuEvent e) {}
        public void itemHighlighted(MenuEvent e) {}
        public void disposing(com.sun.star.lang.EventObject e) {}
      });

      XVclWindowPeer xPeer = WtOfficeTools.getVclWindowPeer(xComponent);
      com.sun.star.awt.Point screenPos = getCaretPosition(xComponent);
//      WtMessageHandler.printToLogFile("Popup will be executed");
      xPopup.execute(xPeer, new Rectangle(screenPos.X, screenPos.Y, 0, 0), (short) 0);
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }
  
  /**
   * change the text of a paragraph independent of the type of document
   * @throws Throwable 
   */
  private void changeTextOfParagraph(int nFPara, int nStart, int nLength, String replace) throws Throwable {
//    try {
      WtMessageHandler.printToLogFile("WtErrorMenu: changeTextOfParagraph: nFPara: " + nFPara + ", nStart:" + nStart
          + ", nLength:" + nLength + ", replace:" + replace);
      WtDocumentCache docCache = document.getDocumentCache();
      String sPara = docCache.getFlatParagraph(nFPara);
      String sEnd = (nStart + nLength < sPara.length() ? sPara.substring(nStart + nLength) : "");
      sPara = sPara.substring(0, nStart) + replace + sEnd;
      WtDocumentCursorTools docCursor = new WtDocumentCursorTools(document.getXComponent());
      int[] textFieldPositions = docCache.getFlatParagraphFieldPositions(nFPara);
      docCursor.changeTextOfParagraphCorrected(docCache.getNumberOfTextParagraph(nFPara), nStart, nLength, replace, textFieldPositions);
      docCache.setFlatParagraph(nFPara, sPara);
      document.removeResultCache(nFPara, true);
      document.removeIgnoredMatch(nFPara, true);
      document.removePermanentIgnoredMatch(nFPara, true);
      WtDocumentsHandler documents = document.getMultiDocumentsHandler();
      WtLanguageTool lt = documents.getLanguageTool();
      if (documents.getConfiguration().useTextLevelQueue() && !documents.getConfiguration().noBackgroundCheck()) {
        for (int i = 1; i < lt.getNumMinToCheckParas().size(); i++) {
            document.addQueueEntry(nFPara, i, lt.getNumMinToCheckParas().get(i), document.getDocID(), true);
        }
      }
//    } catch (Throwable e) {
//      WtMessageHandler.showError(e);
//    }
  }
  

}

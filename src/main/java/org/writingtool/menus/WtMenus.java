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

import org.writingtool.WtSingleDocument;
import org.writingtool.config.WtConfiguration;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtOfficeTools.DocumentType;

import com.sun.star.uno.XComponentContext;

/**
 * Class of menus adding dynamic components 
 * to header menu and to context menu
 * @since 1.0
 * @author Fred Kruse
 */
public class WtMenus {
  
//  public static final String WtProtocolHandler.WT_MENU_REPLACE_COLON = "__|__";
  public static final String WT_MENU_REPLACE_COLON = ":";

  private static boolean debugMode;   //  should be false except for testing
  
  private WtHeadMenu wtHeadMenu = null;
  private WtContextMenu wtContextMenu = null;

  public WtMenus(XComponentContext xContext, WtSingleDocument document, WtConfiguration config) {
    try {
      debugMode = WtOfficeTools.DEBUG_MODE_LM;
      if (document.getDocumentType() == DocumentType.WRITER) {
        wtHeadMenu = new WtHeadMenu(document, config, xContext);
      }
      wtContextMenu = new WtContextMenu(document, config);
      if (debugMode) {
        WtMessageHandler.printToLogFile("WritingToolMenus initialised");
      }
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
    }
  }
  
  public void setConfigValues(WtConfiguration config) {
    if (wtHeadMenu != null) {
      wtHeadMenu.setConfigValues(config);
    }
    wtContextMenu.setConfigValues(config);
  }
  
  public void removeListener() {
    if (wtHeadMenu != null) {
      wtHeadMenu.removeListener();
    }
  }
  
  public void addListener() {
    if (wtHeadMenu != null) {
      wtHeadMenu.addListener();
    }
  }
  
  public static String replaceColon (String str) {
    return str == null ? null : str.replace(":", WT_MENU_REPLACE_COLON);
  }
  


}

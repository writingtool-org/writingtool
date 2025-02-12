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

import com.sun.star.awt.XWindow;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.XComponent;
import com.sun.star.lib.uno.helper.ComponentBase;
import com.sun.star.ui.UIElementType;
import com.sun.star.ui.XToolPanel;
import com.sun.star.ui.XUIElement;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 *  WT sidebar panel
 * @since 1.3
 * @author Fred Kruse
 */
public class WtSidebarPanel extends ComponentBase implements XUIElement {

  private final String resourceUrl;
  private XToolPanel panel;

  public WtSidebarPanel(XComponentContext context, XWindow parentWindow, String resourceUrl) {
    this.resourceUrl = resourceUrl;
    panel = new WtSidebarContent(context, parentWindow);
  }

  @Override
  public XFrame getFrame() {
    return null;
  }

  @Override
  public Object getRealInterface() {
    return panel;
  }

  @Override
  public String getResourceURL() {
    return resourceUrl;
  }

  @Override
  public short getType() {
    return UIElementType.TOOLPANEL;
  }
  
  @Override
  public void dispose()
  {
    XComponent xPanelComponent = UnoRuntime.queryInterface(XComponent.class, panel);
    xPanelComponent.dispose();
  }

}

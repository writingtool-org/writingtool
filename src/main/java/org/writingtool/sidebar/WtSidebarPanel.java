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
import com.sun.star.uno.XComponentContext;

import org.libreoffice.ext.unohelper.dialog.adapter.AbstractSidebarPanel;

/**
 *  WT sidebar panel
 * @since 1.3
 * @author Fred Kruse
 */
public class WtSidebarPanel extends AbstractSidebarPanel {

    public WtSidebarPanel(XComponentContext context, XWindow parentWindow, String resourceUrl)
    {
        super(resourceUrl);
        panel = new WtSidebarContent(context, parentWindow);
    }
}

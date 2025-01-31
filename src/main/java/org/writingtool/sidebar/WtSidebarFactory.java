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
import com.sun.star.beans.PropertyValue;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.ui.XUIElement;
import com.sun.star.uno.XComponentContext;

import org.libreoffice.ext.unohelper.dialog.adapter.AbstractSidebarFactory;
import org.libreoffice.ext.unohelper.common.UNO;

/**
 * Factory for the the WritingTool sidebar.
 * 
 * @since 1.3
 * @author Fred Kruse
 */
public class WtSidebarFactory extends AbstractSidebarFactory {

    public static final String SERVICE_NAME = "org.writingtool.sidebar.WtSidebarFactory";

    public WtSidebarFactory(XComponentContext context) {
      super(SERVICE_NAME, context);
      System.err.println("Construct SidebarFactory");
    }
    
    public static String[] getServiceNames() {
      return new String[] { SERVICE_NAME };
    }

    @Override
    public XUIElement createUIElement(String resourceUrl, PropertyValue[] arguments) throws NoSuchElementException {
        if (!resourceUrl.startsWith("private:resource/toolpanel/WtSidebarFactory"))
        {
            throw new NoSuchElementException(resourceUrl, this);
        }

        XWindow parentWindow = null;
        for (int i = 0; i < arguments.length; i++)
        {
            if (arguments[i].Name.equals("ParentWindow"))
            {
                parentWindow = UNO.XWindow(arguments[i].Value);
                break;
            }
        }

        return new WtSidebarPanel(context, parentWindow, resourceUrl);
    }
}

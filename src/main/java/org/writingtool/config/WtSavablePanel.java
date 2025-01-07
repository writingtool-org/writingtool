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
package org.writingtool.config;

import org.writingtool.dialogs.WtConfigurationDialog;

/**
 * Interface for JPanel that can persist its state.
 * <P>
 * See {@link WtConfigurationDialog#addExtraPanel}
 *
 * @author Panagiotis Minos
 * @since 1.0
 */
public interface WtSavablePanel {

    /**
     * Called when {@link WtConfigurationDialog} is about to be shown.
     */
    public void componentShowing();
    
    /**
     * Invoke the save operation.
     */
    public void save();
}

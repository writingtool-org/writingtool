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
package org.writingtool.dialogs;

import java.awt.*;
import java.awt.event.*;
import javax.swing.*;
import javax.swing.text.JTextComponent;

public class WtTextContextMenu extends JPopupMenu implements ActionListener {
  private static final long serialVersionUID = 1L;
  public static final WtTextContextMenu INSTANCE = new WtTextContextMenu();
  private final JMenuItem itemCut;
  private final JMenuItem itemCopy;
  private final JMenuItem itemPaste;
  private final JMenuItem itemDelete;
  private final JMenuItem itemSelectAll;

  private WtTextContextMenu() {
    itemCut = newItem("Cut", 'T');
    itemCopy = newItem("Copy", 'C');
    itemPaste = newItem("Paste", 'P');
    itemDelete = newItem("Delete", 'D');
    addSeparator();
    itemSelectAll = newItem("Select All", 'A');
  }

  private JMenuItem newItem(String text, char mnemonic) {
    JMenuItem item = new JMenuItem(text, mnemonic);
    item.addActionListener(this);
    return add(item);
  }

  @Override
  public void show(Component invoker, int x, int y) {
    JTextComponent tc = (JTextComponent)invoker;
    boolean changeable = tc.isEditable() && tc.isEnabled();
    itemCut.setVisible(changeable);
    itemPaste.setVisible(changeable);
    itemDelete.setVisible(changeable);
    super.show(invoker, x, y);
  }

  @Override
  public void actionPerformed(ActionEvent e) {
    JTextComponent tc = (JTextComponent)getInvoker();
    tc.requestFocus();
    boolean haveSelection = tc.getSelectionStart() != tc.getSelectionEnd();
    if (e.getSource() == itemCut) {
      if (!haveSelection) tc.selectAll();
      tc.cut();
    } else if (e.getSource() == itemCopy) {
      if (!haveSelection) tc.selectAll();
      tc.copy();
    } else if (e.getSource() == itemPaste) {
      tc.paste();
    } else if (e.getSource() == itemDelete) {
      if (!haveSelection) tc.selectAll();
      tc.replaceSelection("");
    } else if (e.getSource() == itemSelectAll) {
      tc.selectAll();
    }
  }
}


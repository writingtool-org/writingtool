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

import java.awt.Dimension;
import java.awt.MouseInfo;
import java.awt.Point;
import java.awt.event.KeyEvent;
import java.awt.event.KeyListener;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import javax.swing.JCheckBox;
import javax.swing.JTree;
import javax.swing.event.TreeExpansionEvent;
import javax.swing.event.TreeWillExpandListener;
import javax.swing.tree.DefaultTreeModel;
import javax.swing.tree.ExpandVetoException;
import javax.swing.tree.TreePath;

/**
 *
 * @author Panagiotis Minos
 * @since 1.0
 */
public class WtTreeListener implements KeyListener, MouseListener, TreeWillExpandListener {

  public static void install(JTree tree) {
    WtTreeListener listener = new WtTreeListener(tree);
    tree.addMouseListener(listener);
    tree.addKeyListener(listener);
    tree.addTreeWillExpandListener(listener);
  }

  private static final Dimension checkBoxDimension = new JCheckBox().getPreferredSize();

  private final JTree tree;

  private WtTreeListener(JTree tree) {
    this.tree = tree;
  }

  @Override
  public void keyTyped(KeyEvent e) {
  }

  @Override
  public void keyPressed(KeyEvent e) {
    if (e.getKeyCode() == KeyEvent.VK_SPACE) {
      TreePath[] paths = tree.getSelectionPaths();
      if (paths != null) {
        for (TreePath path : paths) {
          handle(path);
        }
      }
    }
  }

  @Override
  public void keyReleased(KeyEvent e) {
  }

  @Override
  public void mouseClicked(MouseEvent e) {
  }

  @Override
  public void mousePressed(MouseEvent e) {
    int x = e.getX();
    int y = e.getY();
    TreePath path = tree.getPathForLocation(x, y);
    if (isOverCheckBox(x, y, path)) {
      handle(path);
    }
  }

  @Override
  public void mouseReleased(MouseEvent e) {
  }

  @Override
  public void mouseEntered(MouseEvent e) {
  }

  @Override
  public void mouseExited(MouseEvent e) {
  }

  private void handle(TreePath path) {
    if ((path != null) && (path.getPathCount() > 0)) {
      if (path.getLastPathComponent() instanceof WtCategoryNode) {
        DefaultTreeModel model = (DefaultTreeModel) tree.getModel();

        WtCategoryNode node = (WtCategoryNode) path.getLastPathComponent();
        node.setEnabled(!node.isEnabled());
        model.nodeChanged(node);

        for (int i = 0; i < node.getChildCount(); i++) {
          WtRuleNode child = (WtRuleNode) node.getChildAt(i);
          if (child.isEnabled() != node.isEnabled()) {
            child.setEnabled(node.isEnabled());
            model.nodeChanged(child);
          }
        }
      }
      if (path.getLastPathComponent() instanceof WtRuleNode) {
        DefaultTreeModel model = (DefaultTreeModel) tree.getModel();

        WtRuleNode node = (WtRuleNode) path.getLastPathComponent();
        node.setEnabled(!node.isEnabled());
        model.nodeChanged(node);

        if (node.isEnabled()) {
          WtCategoryNode parent = (WtCategoryNode) node.getParent();
          parent.setEnabled(true);
        }
        model.nodeChanged(node.getParent());
      }
    }
  }

  private boolean isOverCheckBox(int x, int y, TreePath path) {
    if ((path == null) || (path.getPathCount() == 0)) {
        return false;
    }
    if (!isValidNode(path.getLastPathComponent())) {
      return false;
    }
    //checkbox is east
    //int offset = tree.getPathBounds(path).x + tree.getPathBounds(path).width - checkBoxDimension.width;
    //if (x < offset) {

    //checkbox is west
    int offset = tree.getPathBounds(path).x + checkBoxDimension.width;
    if (x > offset) {
      return false;
    }
    return true;
  }

  private boolean isValidNode(Object c) {
    return ((c instanceof WtCategoryNode) || (c instanceof WtRuleNode));
  }

  @Override
  public void treeWillExpand(TreeExpansionEvent e) throws ExpandVetoException {
    Point cursorPosition = MouseInfo.getPointerInfo().getLocation();
    Point treePosition = tree.getLocationOnScreen();
    int x = (int) (cursorPosition.getX() - treePosition.getX());
    int y = (int) (cursorPosition.getY() - treePosition.getY());
    TreePath path = tree.getPathForLocation(x, y);
    if (isOverCheckBox(x, y, path)) {
      throw new ExpandVetoException(e);
    }
  }

  @Override
  public void treeWillCollapse(TreeExpansionEvent e) throws ExpandVetoException {
    treeWillExpand(e);
  }
}

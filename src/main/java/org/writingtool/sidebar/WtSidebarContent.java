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

import java.awt.SystemColor;
import java.util.ResourceBundle;
import java.util.SortedMap;
import java.util.TreeMap;

import com.sun.star.accessibility.XAccessible;
import com.sun.star.awt.ImagePosition;
import com.sun.star.awt.Rectangle;
import com.sun.star.awt.WindowEvent;
import com.sun.star.awt.XButton;
import com.sun.star.awt.XControl;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XFixedText;
import com.sun.star.awt.XTextComponent;
import com.sun.star.awt.XToolkit;
import com.sun.star.awt.XWindow;
import com.sun.star.awt.XWindowPeer;
import com.sun.star.awt.tab.XTabPageContainer;
import com.sun.star.awt.tab.XTabPageContainerModel;
import com.sun.star.lang.DisposedException;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lib.uno.helper.ComponentBase;
import com.sun.star.text.XTextDocument;
import com.sun.star.ui.LayoutSize;
import com.sun.star.ui.XSidebarPanel;
import com.sun.star.ui.XToolPanel;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import org.libreoffice.ext.unohelper.common.UNO;
import org.libreoffice.ext.unohelper.common.UnoHelperException;
import org.libreoffice.ext.unohelper.dialog.adapter.AbstractActionListener;
import org.libreoffice.ext.unohelper.dialog.adapter.AbstractWindowListener;
import org.libreoffice.ext.unohelper.ui.GuiFactory;
import org.libreoffice.ext.unohelper.ui.layout.HorizontalLayout;
import org.libreoffice.ext.unohelper.ui.layout.Layout;
import org.libreoffice.ext.unohelper.ui.layout.VerticalLayout;
import org.libreoffice.ext.unohelper.util.UnoProperty;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;

/**
 * Create the window for the WT sidebar panel
 * 
 * @since 1.3
 * @author Fred Kruse
 */
public class WtSidebarContent extends ComponentBase implements XToolPanel, XSidebarPanel {

  private static final ResourceBundle messages = WtOfficeTools.getMessageBundle();

  private XMultiComponentFactory xMCF;  // The component factory
  private XControlContainer controlContainer;  //  The container of the controls
  private Layout layout;  //  The layout of the controls
  private Layout layout2;  //  The layout of the controls

  public WtSidebarContent(XComponentContext context, XWindow parentWindow) {
    AbstractWindowListener windowAdapter = new AbstractWindowListener() {
      @Override
      public void windowResized(WindowEvent e) {
        layout.layout(parentWindow.getPosSize());
      }
    };
    parentWindow.addWindowListener(windowAdapter);
    layout = new VerticalLayout(5, 5, 5, 5, 5);

    xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class, context.getServiceManager());
    XWindowPeer parentWindowPeer = UNO.XWindowPeer(parentWindow);

    if (parentWindowPeer == null) {
        return;
    }

    XToolkit parentToolkit = parentWindowPeer.getToolkit();
    controlContainer = UNO
            .XControlContainer(GuiFactory.createControlContainer(xMCF, context, new Rectangle(0, 0, 0, 0), null));
    XControl xContainer = UNO.XControl(controlContainer);
    xContainer.createPeer(parentToolkit, parentWindowPeer);
/*
    controlContainer2 = UNO
        .XControlContainer(GuiFactory.createControlContainer(xMCF, context, new Rectangle(0, 0, 0, 0), null));
    UNO.XControl(controlContainer2).createPeer(parentToolkit, xContainer.getPeer());
/*    
    XControl tabControl = GuiFactory.createTabPageContainer(xMCF, context);
    XTabPageContainer tabContainer = UNO.XTabPageContainer(tabControl);
    XTabPageContainerModel tabModel = UNO.XTabPageContainerModel(tabControl.getModel());
    GuiFactory.createTab(xMCF, context, tabModel, "Test1", (short) 1, 100);
    controlContainer.addControl("tabs", UNO.XControl(tabContainer));
    layout.addControl(UNO.XControl(tabContainer));
*/
    // Add text field
    SortedMap<String, Object> props = new TreeMap<>();
    props.put("TextColor", SystemColor.textInactiveText.getRGB() & ~0xFF000000);
    props.put("Autocomplete", false);
    props.put("HideInactiveSelection", true);
    XFixedText searchBox = UNO.XFixedText(
        GuiFactory.createLabel(xMCF, context, WtOfficeTools.getFormatedTextVersionInformation(), new Rectangle(0, 0, 100, 32), props));
//    XTextComponent searchBox = UNO.XTextComponent(
//            GuiFactory.createTextfield(xMCF, context, WtOfficeTools.getFormatedLicenseInformation(), new Rectangle(0, 0, 100, 32), props, null));
    controlContainer.addControl("searchbox", UNO.XControl(searchBox));
    layout.addControl(UNO.XControl(searchBox));

    // Add button
    XControl button = GuiFactory.createButton(xMCF, context, "Insert", null, new Rectangle(0, 0, 100, 32), null);
    XButton xbutton = UNO.XButton(button);
    AbstractActionListener xButtonAction = event -> {
      String text = searchBox.getText();
      try {
        UNO.init(xMCF);
        XTextDocument doc = UNO.getCurrentTextDocument();
        doc.getText().setString(text);
      } catch (UnoHelperException e1) {
        WtMessageHandler.printException(e1);
      }
    };
    xbutton.addActionListener(xButtonAction);
    controlContainer.addControl("button1", button);
    layout.addControl(button);

    XControl xButtonContainer = GuiFactory.createControlContainer(xMCF, context, new Rectangle(0, 0, 100, 32), null);
    controlContainer.addControl("buttonContainer", xButtonContainer);
    layout.addControl(xButtonContainer);
    XControlContainer buttonContainer = UNO.XControlContainer(xButtonContainer);
    XWindow buttonContainerWindow = UNO.XWindow(xButtonContainer);
    AbstractWindowListener buttonContainerAdapter = new AbstractWindowListener() {
      @Override
      public void windowResized(WindowEvent e) {
        layout2.layout(buttonContainerWindow.getPosSize());
      }
    };
    buttonContainerWindow.addWindowListener(buttonContainerAdapter);
    
    // Add button
    SortedMap<String, Object> bProps = new TreeMap<>();
    bProps.put(UnoProperty.IMAGE_URL, "vnd.sun.star.extension://org.openoffice.writingtool.oxt/images/WTNextBig.png");
    bProps.put("ImagePosition", ImagePosition.LeftCenter);
    XControl button1 = GuiFactory.createButton(xMCF, context, "Next Error", null, new Rectangle(0, 0, 15, 15), bProps);
//    XButton xbutton1 = UNO.XButton(button1);
/*    
    AbstractActionListener xButtonAction = event -> {
      String text = searchBox.getText();
      try {
        UNO.init(xMCF);
        XTextDocument doc = UNO.getCurrentTextDocument();
        doc.getText().setString(text);
      } catch (UnoHelperException e1) {
        WtMessageHandler.printException(e1);
      }
    };
    xbutton.addActionListener(xButtonAction);
*/
    
    layout2 = new HorizontalLayout(5, 5, 5, 5, 5);
    controlContainer.addControl("button2", button1);
    layout.addControl(button1);
  }

  @Override
  public XAccessible createAccessible(XAccessible a) {
    return UNO.XAccessible(getWindow());
  }

  @Override
  public XWindow getWindow() {
    if (controlContainer == null) {
      throw new DisposedException("", this);
    }
    return UNO.XWindow(controlContainer);
  }

  @Override
  public LayoutSize getHeightForWidth(int width) {
    int height = layout.getHeightForWidth(width);
    return new LayoutSize(height, height, height);
  }

  @Override
  public int getMinimalWidth() {
    return 300;
  }
}

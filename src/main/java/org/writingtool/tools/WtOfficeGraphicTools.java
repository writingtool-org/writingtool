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

import java.io.File;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.ElementExistException;
import com.sun.star.container.XNameContainer;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawView;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XController;
import com.sun.star.frame.XModel;
import com.sun.star.graphic.XGraphic;
import com.sun.star.graphic.XGraphicProvider;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.text.XText;
import com.sun.star.text.XTextContent;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextDocument;
import com.sun.star.text.XTextRange;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 * Some tools to handle graphics in LibreOffice/OpenOffice document context
 * @since 1.0
 * @author Fred Kruse
 */
public class WtOfficeGraphicTools {
  
  private final static int SIZE_FACTOR = 20;
  private final static int SIZE_Del = 8;
  private final static String BITMAP_NAME_PREFIX = "WtAiImage";

  public static void insertGraphic(String strUrl, int size, XComponent xComp, XComponentContext xContext) {
    try {
      XTextDocument xTextDoc = UnoRuntime.queryInterface(XTextDocument.class, xComp);
      XMultiServiceFactory xMSFDoc = UnoRuntime.queryInterface(XMultiServiceFactory.class, xTextDoc);

      Object oGraphic = null;
      try {
          // Creating the service GraphicObject
          oGraphic =xMSFDoc.createInstance("com.sun.star.text.TextGraphicObject");
      } catch (Exception e) {
          WtMessageHandler.printToLogFile("Could not create instance");
          WtMessageHandler.printException(e);
      }

      // Getting the text
      XText xText = xTextDoc.getText();

      // Getting the cursor on the document (current position)
//      com.sun.star.text.XTextCursor xTextCursor = xText.createTextCursor();
      WtViewCursorTools vCursor = new WtViewCursorTools(xComp);
      XTextCursor xTextCursor = vCursor.getTextCursorBeginn();

      // Querying for the interface XTextContent on the GraphicObject
      XTextContent xTextContent = UnoRuntime.queryInterface(XTextContent.class, oGraphic);

      // Printing information to the log file
      WtMessageHandler.printToLogFile("inserting graphic");
      try {
        // Inserting the content
        xText.insertTextContent(xTextCursor, xTextContent, true);
      } catch (Exception e) {
        WtMessageHandler.printToLogFile("Could not insert Content");
        WtMessageHandler.printException(e);
      }

      // Printing information to the log file
      WtMessageHandler.printToLogFile("adding graphic");

      // Querying for the interface XPropertySet on GraphicObject
      XPropertySet xPropSet = UnoRuntime.queryInterface(XPropertySet.class, oGraphic);
      try {
        // Creating a string for the graphic url
        java.io.File sourceFile = new java.io.File(strUrl);
        StringBuffer sUrl = new StringBuffer("file:///");
        sUrl.append(sourceFile.getCanonicalPath().replace('\\', '/'));
        WtMessageHandler.printToLogFile("insert graphic \"" + sUrl + "\"");
        
        XGraphic xGraphic = getGraphic(sUrl.toString(), xContext);

        // Setting the anchor type
        xPropSet.setPropertyValue("AnchorType",
                   com.sun.star.text.TextContentAnchorType.AT_PARAGRAPH );

        // Setting the graphic url
        xPropSet.setPropertyValue( "Graphic", xGraphic );

        // Setting the horizontal position
//        xPropSet.setPropertyValue( "HoriOrientPosition", Integer.valueOf( 5500 ) );

        // Setting the vertical position
//        xPropSet.setPropertyValue( "VertOrientPosition", Integer.valueOf( 4200 ) );

        // Setting the width
        xPropSet.setPropertyValue("Width", size * SIZE_FACTOR);

        // Setting the height
        xPropSet.setPropertyValue("Height", size * SIZE_FACTOR);
      } catch (Exception e) {
        WtMessageHandler.printToLogFile("Couldn't set property 'GraphicURL'");
        WtMessageHandler.printException(e);
      }
    }
    catch( Exception e ) {
      WtMessageHandler.printException(e);
    }
  }
  
  private static XGraphic getGraphic(String sUrl, XComponentContext xContext) {
    try {
      XMultiComponentFactory xMCF = xContext.getServiceManager();
      XGraphicProvider xGraphicProvider = UnoRuntime.queryInterface(XGraphicProvider.class,
          xMCF.createInstanceWithContext("com.sun.star.graphic.GraphicProvider", xContext));
    
      PropertyValue[] aMediaProps = new PropertyValue[] { new PropertyValue() };
      aMediaProps[0].Name = "URL";
      aMediaProps[0].Value = sUrl.toString();
      
      return UnoRuntime.queryInterface(XGraphic.class, xGraphicProvider.queryGraphic(aMediaProps));
    } catch (Throwable t) {
      WtMessageHandler.showError(t);
      return null;
    }
  }
  
  //  For Impress
  
  public static void insertGraphicInImpress(String strUrl, int size, XComponent xComponent, XComponentContext xContext) {
    try {
      XModel xModel = UnoRuntime.queryInterface(XModel.class, xComponent);
      XController xController = xModel.getCurrentController();
      XDrawView xDrawView = UnoRuntime.queryInterface(XDrawView.class, xController);
      XDrawPage xCurrentDrawPage = xDrawView.getCurrentPage();
      drawImage(xCurrentDrawPage, strUrl, 0, 0, size / SIZE_Del, size / SIZE_Del, xComponent, xContext);    
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
  }
  
  private static XShape drawImage(XDrawPage slide, String imFnm, int x, int y, int width, int height,
      XComponent xComponent, XComponentContext xContext) {
    // units in mm's
    WtMessageHandler.printToLogFile("Adding picture \"" + imFnm + "\"");
    XShape imShape = addShape(slide, "GraphicObjectShape", 
                                x, y, width, height, xComponent);
    setImage(imShape, imFnm, xComponent);
    setLineStyle(imShape, LineStyle.NONE);
            // so no border around the image
    return imShape;
  }  // end of drawImage()

  private static XShape addShape(XDrawPage slide, String shapeType, 
      int x, int y, int width, int height, XComponent xComponent) { 
    warnsPosition(slide, x, y);
    XShape shape = makeShape(shapeType, x, y, width, height, xComponent);
    if (shape != null) { 
      slide.add(shape);
    } else {
      WtMessageHandler.printToLogFile("shape == null: not added to slide!");
    }
    return shape;
  }  // end of addShape()
  
  private static XShape makeShape(String shapeType, int x, int y, int width, int height, XComponent xComponent) {
    // parameters are in mm units 
    XShape shape = null;
    try {
      XMultiServiceFactory msf = UnoRuntime.queryInterface(XMultiServiceFactory.class, xComponent);
      Object o = msf.createInstance("com.sun.star.drawing." + shapeType);
      shape = UnoRuntime.queryInterface(XShape.class, o);
      shape.setPosition( new Point(x*100,y*100) );
      shape.setSize( new Size(width*100, height*100) );
    } catch(Exception e) {
      WtMessageHandler.showError(e);
    }
    return shape;
  }  // end of makeShape()
  
  private static void warnsPosition(XDrawPage slide, int x, int y)
  // warns if (x, y) is not on the page
  {
    Size slideSize = getSlideSize(slide);
    if (slideSize == null) {
      WtMessageHandler.printToLogFile("No slide size found");
      return;
    }
    int slideWidth = slideSize.Width;
    int slideHeight = slideSize.Height;

    if (x < 0) {
      WtMessageHandler.printToLogFile("x < 0");
    } else if (x > slideWidth-1) {
      WtMessageHandler.printToLogFile("x position off right hand side of the slide");
    }
    if (y < 0) {
      System.out.println("y < 0");
    } else if (y > slideHeight-1) {
      WtMessageHandler.printToLogFile("y position off bottom of the slide");
    }
  }  // end of warnsPosition()

  private static Size getSlideSize(XDrawPage xDrawPage)
  // get size of the given slide page (in mm units)
  {
    try {
      XPropertySet props = UnoRuntime.queryInterface(XPropertySet.class, xDrawPage);
      if (props == null) {
         System.out.println("No slide properties found");
         return null;
       }
      int width = (Integer)props.getPropertyValue("Width");
      int height = (Integer)props.getPropertyValue("Height");
      return new Size(width/100, height/100);
    }
    catch(Exception e)
    {  System.out.println("Could not get page dimensions");
       return null;
    }
  }  // end of getSlideSize()

  private static void setLineStyle(XShape shape, LineStyle style) {
    try {
      XPropertySet propSet = UnoRuntime.queryInterface(XPropertySet.class, shape);
      propSet.setPropertyValue("LineStyle", style);
    } catch(Throwable e) {
      WtMessageHandler.printException(e);
    }
  }  

  private static void setImage(XShape shape, String imFnm, XComponent xComponent) {
    try {
      Object bitmap = getBitmap(imFnm, xComponent);
      XPropertySet propSet = UnoRuntime.queryInterface(XPropertySet.class, shape);
      WtMessageHandler.printToLogFile("Set Image: Url: " + imFnm + ", bitmap " + (bitmap == null ? "==" : "!=") + " null");
      propSet.setPropertyValue("GraphicURL", bitmap);
//      XGraphic xGraphic = loadGraphicFile(imFnm, xcc);
//      propSet.setPropertyValue("Graphic", xGraphic);
    } catch (Throwable e) {
      WtMessageHandler.printException(e);
    }
  }  // end of setImage()
  
  private static Object getBitmap(String fnm, XComponent xComponent)
  // load the graphic as a bitmap, and return it as a string
  {
    try {
      XMultiServiceFactory msf = UnoRuntime.queryInterface(XMultiServiceFactory.class, xComponent);
      Object o;
      o = msf.createInstance("com.sun.star.drawing.BitmapTable");
      XNameContainer bitmapContainer =  UnoRuntime.queryInterface(XNameContainer.class, o);

      // insert image into container
      if (!isOpenable(fnm)) return null;
      String picURL = fnmToURL(fnm);
      if (picURL == null) return null;
      int iNum = 0;
      boolean bSet = true;
      String imgName = null;
      while (bSet) {
        iNum++;
        imgName = BITMAP_NAME_PREFIX + iNum;
        try {
          bitmapContainer.insertByName(imgName, picURL);
          bSet = false;
        } catch (ElementExistException e) {
        }
      }
      return bitmapContainer.getByName(imgName);
    }
    catch(Exception e) {
      WtMessageHandler.showError(e);
      return null;
    }
  }  // end of getBitmap()

  private static boolean isOpenable(String fnm)
  // convert a file path to URL format
  {
     File f = new File(fnm);
     if (!f.exists()) {
       WtMessageHandler.showMessage(fnm + " does not exist");
       return false;
     }
     if (!f.isFile()) {
       WtMessageHandler.showMessage(fnm + " is not a file");
       return false;
     }
     if (!f.canRead()) {
       WtMessageHandler.showMessage(fnm + " is not readable");
       return false;
     }
     return true;
  } // end of isOpenable()

  private static String fnmToURL(String fnm)
  // convert a file path to URL format
  {
     try {
       StringBuffer sb = null;
       String path = new File(fnm).getCanonicalPath();
       sb = new StringBuffer("file:///");
       sb.append(path.replace('\\', '/'));
       return sb.toString();
     }
     catch (java.io.IOException e) {
       WtMessageHandler.printException(e);
       return null;
     }
  }

  public static void insertDrawText(String txt, int width, int height, XComponent xComponent) {
    try {
      XModel xModel = UnoRuntime.queryInterface(XModel.class, xComponent);
      XController xController = xModel.getCurrentController();
      XDrawView xDrawView = UnoRuntime.queryInterface(XDrawView.class, xController);
      XDrawPage xCurrentDrawPage = xDrawView.getCurrentPage();
      XShape shape = addShape(xCurrentDrawPage, "TextShape", 0, 0, width, height, xComponent);
      addText(shape, txt);
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
  }
  
  private static void addText(XShape shape, String txt)
  {
    XText xText = UnoRuntime.queryInterface(XText.class, shape);
    XTextCursor cursor = xText.createTextCursor();
    cursor.gotoEnd(false);
    XTextRange range = UnoRuntime.queryInterface(XTextRange.class, cursor);
    range.setString(txt);
  }  // end of addText()







  
  

}

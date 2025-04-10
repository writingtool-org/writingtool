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

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.URL;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;

import org.languagetool.JLanguageTool;

import com.sun.star.beans.PropertyValue;
import com.sun.star.container.XNameAccess;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 * class to get information about the application and the environment
 * @since 1.1
 * @author Fred Kruse
 */
public class WtVersionInfo {
  public static String wtName = null;
  public static String wtVersion;
  public static String wtBuildDate;
  public static String ltName;
  public static String ltVersion;
  public static String ltBuildDate;
  public static String ltShortGitId;
  public static String ooName;
  public static String ooVersion;
  public static String ooExtension;
  public static String ooVendor;
  public static String ooLocale;
  public static String osArch;
  public static String javaVersion;
  public static String javaVendor;
  public static IOException ioEx;
  public static Throwable thEx;
  
  public static void init(XComponentContext xContext) {
    setWtInfo();
    setLtInfo();
    setOfficeProductInfo(xContext);
    setJavaInfo();
  }
  
  /**
   * Get information about WritingTool
   */
  public static String getWtInformation () {
    String txt = wtVersion;
    if (wtVersion != null && wtVersion.contains("SNAPSHOT")) {
      txt += " - " + wtBuildDate;
    }
    return txt;
  }

  /**
   * Get name and information about WritingTool
   */
  public static String getWtNameWithInformation () {
    return wtName + " " + getWtInformation();
  }
  
  /**
   * get LanguageTool version
   * NOTE: The recommended method doesn't work in WritingTool
   */
  public static String ltVersion() {
    return ltVersion;
  }

  /**
   * get LanguageTool build date
   * NOTE: The recommended method doesn't work in WritingTool
   */
  public static String ltBuildDate() {
    return ltBuildDate;
  }

  /**
   * get LanguageTool short git id
   * NOTE: The recommended method doesn't work in WritingTool
   */
  public static String ltShortGitId() {
    return ltShortGitId;
  }

  /**
   * Get information about LanguageTool
   */
  public static String getLtInformation () {
    String txt = ltVersion();
    if (txt.contains("SNAPSHOT")) {
      txt += " - " + ltBuildDate() + ", " + ltShortGitId();
    }
    return txt;
  }

  /**
   * Set WT version and build date from version file
   */
  private static void setWtInfo() {
    WtOfficeTools wtTools = new WtOfficeTools();
    URL resource = wtTools.getClass().getResource(wtTools.getClass().getSimpleName() + ".class");
    if (resource != null) {
      try {
        String dir = resource.getPath();
        dir = dir.substring(5, dir.indexOf("!"));  // get jar file
        dir = dir.substring(0, dir.lastIndexOf("/"));  // get directory
        dir = URLDecoder.decode(dir, "UTF-8");  // decode URL
        File vFile = new File(dir, "version.txt");
        try (InputStream stream = new FileInputStream(vFile);
            InputStreamReader reader = new InputStreamReader(stream, StandardCharsets.UTF_8);
            BufferedReader br = new BufferedReader(reader)
           ) {
          String line;
          while ((line = br.readLine()) != null) {
            String lines[] = line.split("=");
            if(lines.length > 1) {
              if ("version".equals(lines[0].trim())) {
                wtVersion = new String(lines[1].trim());
              } else  if ("build.date".equals(lines[0].trim())) {
                wtBuildDate = new String(lines[1].trim());
            }
            }
          }
        } catch (IOException e) {
          ioEx = e;
        }
      } catch (Throwable e1) {
        thEx = e1;
      }
    }
    wtName = WtOfficeTools.WT_NAME;
  }
  
  /**
   * Set LT Information
   */
  @SuppressWarnings("deprecation")
  private static void setLtInfo() {
    ltName = "LanguageTool";
    ltVersion = JLanguageTool.VERSION;
    ltBuildDate = JLanguageTool.BUILD_DATE;
    ltShortGitId = JLanguageTool.GIT_SHORT_ID;
  }

  /**
   * Set information of LO/OO office product from system
   */
  private static void setOfficeProductInfo(XComponentContext xContext) {
    try {
      if (xContext == null) {
        return;
      }
      XMultiServiceFactory xMSF = UnoRuntime.queryInterface(XMultiServiceFactory.class, xContext.getServiceManager());
      if (xMSF == null) {
        WtMessageHandler.printToLogFile("XMultiServiceFactory == null");
        return;
      }
      Object oConfigProvider = xMSF.createInstance("com.sun.star.configuration.ConfigurationProvider");
      XMultiServiceFactory confMsf = UnoRuntime.queryInterface(XMultiServiceFactory.class, oConfigProvider);
      if (oConfigProvider == null) {
        WtMessageHandler.printToLogFile("oConfigProvider == null");
        return;
      }

      final String sView = "com.sun.star.configuration.ConfigurationAccess";

      Object args[] = new Object[1];
      PropertyValue aPathArgument = new PropertyValue();
      aPathArgument.Name = "nodepath";
      aPathArgument.Value = "org.openoffice.Setup/Product";
      args[0] = aPathArgument;
      Object oConfigAccess =  confMsf.createInstanceWithArguments(sView, args);
      XNameAccess xName = UnoRuntime.queryInterface(XNameAccess.class, oConfigAccess);
      
      aPathArgument.Value = "org.openoffice.Setup/L10N";
      Object oConfigAccess1 =  confMsf.createInstanceWithArguments(sView, args);
      XNameAccess xName1 = UnoRuntime.queryInterface(XNameAccess.class, oConfigAccess1);
      if (xName != null) {
        ooName = (String) xName.getByName("ooName");
        ooVersion = (String) xName.getByName("ooSetupVersion");
        ooExtension = (String) xName.getByName("ooSetupExtension");
        ooVendor = (String) xName.getByName("ooVendor");
      }
      if (xName1 != null) {
        ooLocale = (String) xName1.getByName("ooLocale");
      }
      osArch = System.getProperty("os.arch");
    } catch (Throwable t) {
      WtMessageHandler.printException(t);     // all Exceptions thrown by UnoRuntime.queryInterface are caught
    }
  }
  
  private static void setJavaInfo() {
    javaVersion = System.getProperty("java.version");
    javaVendor = System.getProperty("java.vm.vendor");
  }
  

}

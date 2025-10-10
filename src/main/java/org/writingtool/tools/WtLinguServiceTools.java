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

import java.util.ArrayList;
import java.util.List;

import org.writingtool.WtDocumentsHandler;
import org.writingtool.WtSpellChecker;

import com.sun.star.beans.XFastPropertySet;
import com.sun.star.beans.XPropertySet;
import com.sun.star.lang.Locale;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.linguistic2.XLinguServiceManager;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

/**
 * Class to handle general tools for linguistic services of LibreOffice/OpenOffice
 * @since 26.1
 * @author Fred Kruse
 */
public class WtLinguServiceTools {
  
  //  fast handles defined in LibeOffice source: unotools/source/config/lingucfg.cxx
  private static int FH_IS_SPELL_AUTO = 8;
  private static int FH_IS_GRAMMAR_AUTO = 29;

  /** 
   * Get the LinguServiceManager to be used for example 
   * to access spell checker, grammar checker, thesaurus and hyphenator
   */
  public static XLinguServiceManager getLinguSvcMgr(XComponentContext xContext) {
    try {
      XMultiComponentFactory xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class,
          xContext.getServiceManager());
      if (xMCF == null) {
        WtMessageHandler.printToLogFile("WtLinguServiceTools: getLinguSvcMgr: XMultiComponentFactory == null");
        return null;
      }
      // retrieve Office's remote component context as a property
      XPropertySet props = UnoRuntime.queryInterface(XPropertySet.class, xMCF);
      if (props == null) {
        WtMessageHandler.printToLogFile("WtLinguServiceTools: getLinguSvcMgr: XPropertySet == null");
        return null;
      }
      Object defaultContext = props.getPropertyValue("DefaultContext");
      // get the remote interface XComponentContext
      XComponentContext xComponentContext = UnoRuntime.queryInterface(XComponentContext.class, defaultContext);
      if (xComponentContext == null) {
        WtMessageHandler.printToLogFile("WtLinguServiceTools: getLinguSvcMgr: XComponentContext == null");
        return null;
      }
      Object o = xMCF.createInstanceWithContext("com.sun.star.linguistic2.LinguServiceManager", xComponentContext);     
      // create service component using the specified component context
      XLinguServiceManager mxLinguSvcMgr = UnoRuntime.queryInterface(XLinguServiceManager.class, o);
      if (mxLinguSvcMgr == null) {
        WtMessageHandler.printToLogFile("WtLinguServiceTools: getLinguSvcMgr: XLinguServiceManager2 == null");
        return null;
      }
      return mxLinguSvcMgr;
    } catch (Throwable t) {
      // If anything goes wrong, give the user a stack trace
      WtMessageHandler.printException(t);
    }
    return null;
  }
  
  /** 
   * Get the LinguServiceManager to be used for example 
   * to access spell checker, grammar checker, thesaurus and hyphenator
   */
  public static void getLinguProperties(XComponentContext xContext) {
    try {
      XMultiComponentFactory xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class,
          xContext.getServiceManager());
      if (xMCF == null) {
        WtMessageHandler.printToLogFile("WtLinguServiceTools: getLinguSvcMgr: XMultiComponentFactory == null");
        return;
      }
      Object o = xMCF.createInstanceWithContext("com.sun.star.linguistic2.LinguProperties", xContext);     
      WtMessageHandler.printToLogFile("WtLinguServiceTools: printPropertySet: XLinguServiceManager");
      WtOfficeTools.printPropertySet(o);
      o = xMCF.createInstanceWithContext("com.sun.star.linguistic2.Proofreader", xContext);     
      WtMessageHandler.printToLogFile("WtLinguServiceTools: printPropertySet: XProofreader");
      if (o == null) {
        WtMessageHandler.printToLogFile("WtLinguServiceTools: getLinguSvcMgr: XProofreader == null");
        return;
      }
      WtOfficeTools.printPropertySet(o);
/*
      XLinguServiceManager linguSvcMgr = getLinguSvcMgr(xContext);
      WtMessageHandler.printToLogFile("WtLinguServiceTools: printPropertySet: XLinguServiceManager");
      WtOfficeTools.printPropertySet(linguSvcMgr);
      XSpellChecker xSpellChecker = linguSvcMgr.getSpellChecker();
      WtOfficeTools.printPropertySet(xSpellChecker);
*/
    } catch (Throwable t) {
      // If anything goes wrong, give the user a stack trace
      WtMessageHandler.printException(t);
    }
  }

  /**
   * is automatic spell active
   */
  public static boolean isSpellAuto(XComponentContext xContext) {
    return isSettingHandle(FH_IS_SPELL_AUTO, xContext);
  }

  /**
   * is automatic grammar check active
   */
  public static boolean isGrammarAuto(XComponentContext xContext) {
    return isSettingHandle(FH_IS_GRAMMAR_AUTO, xContext);
  }

  /**
   * is a fast handle set to true
   */
  public static boolean isSettingHandle(int handle, XComponentContext xContext) {
    try {
      XMultiComponentFactory xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class,
          xContext.getServiceManager());
      if (xMCF == null) {
        WtMessageHandler.printToLogFile("WtLinguServiceTools: isSpellAuto: XMultiComponentFactory == null");
        return false;
      }
      Object o = xMCF.createInstanceWithContext("com.sun.star.linguistic2.LinguProperties", xContext);     
      XFastPropertySet fProps = UnoRuntime.queryInterface(XFastPropertySet.class, o);
      if (fProps == null) {
        WtMessageHandler.printToLogFile("WtLinguServiceTools: isSpellAuto: XPropertySet == null");
        return false;
      }
      return (boolean) fProps.getFastPropertyValue(handle);
    } catch (Throwable t) {
      // If anything goes wrong, give the user a stack trace
      WtMessageHandler.printException(t);
    }
    return false;
  }
  
  /**
   * set value of automatic spell active
   */
  public static void setSpellAuto(boolean value, XComponentContext xContext) {
    setSettingHandle(FH_IS_SPELL_AUTO, value, xContext);
  }

  /**
   * set value of automatic grammar check active
   */
  public static void setGrammarAuto(boolean value, XComponentContext xContext) {
    setSettingHandle(FH_IS_GRAMMAR_AUTO, value, xContext);
  }

  /**
   * set a fast handle (only boolean values)
   */
  public static void setSettingHandle(int handle, boolean value, XComponentContext xContext) {
    try {
      XMultiComponentFactory xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class,
          xContext.getServiceManager());
      if (xMCF == null) {
        WtMessageHandler.printToLogFile("WtLinguServiceTools: isSpellAuto: XMultiComponentFactory == null");
        return;
      }
      Object o = xMCF.createInstanceWithContext("com.sun.star.linguistic2.LinguProperties", xContext);     
      XFastPropertySet fProps = UnoRuntime.queryInterface(XFastPropertySet.class, o);
      if (fProps == null) {
        WtMessageHandler.printToLogFile("WtLinguServiceTools: isSpellAuto: XPropertySet == null");
        return;
      }
      fProps.setFastPropertyValue(handle, value);
    } catch (Throwable t) {
      // If anything goes wrong, give the user a stack trace
      WtMessageHandler.printException(t);
    }
  }
  
/*
  public static boolean isSetting(String property, XComponentContext xContext) {
    try {
      XMultiComponentFactory xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class,
          xContext.getServiceManager());
      if (xMCF == null) {
        WtMessageHandler.printToLogFile("WtLinguServiceTools: isSpellAuto: XMultiComponentFactory == null");
        return false;
      }
      Object o = xMCF.createInstanceWithContext("com.sun.star.linguistic2.LinguProperties", xContext);     
      XPropertySet props = UnoRuntime.queryInterface(XPropertySet.class, o);
      if (props == null) {
        WtMessageHandler.printToLogFile("WtLinguServiceTools: isSpellAuto: XPropertySet == null");
        return false;
      }
      return (boolean) props.getPropertyValue(property);
    } catch (Throwable t) {
      // If anything goes wrong, give the user a stack trace
      WtMessageHandler.printException(t);
    }
    return false;
  }

  public static void setSpellAuto(boolean spellAuto, XComponentContext xContext) {
    try {
      XMultiComponentFactory xMCF = UnoRuntime.queryInterface(XMultiComponentFactory.class,
          xContext.getServiceManager());
      if (xMCF == null) {
        WtMessageHandler.printToLogFile("WtLinguServiceTools: isSpellAuto: XMultiComponentFactory == null");
        return;
      }
      Object o = xMCF.createInstanceWithContext("com.sun.star.linguistic2.LinguProperties", xContext);     
      XPropertySet props = UnoRuntime.queryInterface(XPropertySet.class, o);
      if (props == null) {
        WtMessageHandler.printToLogFile("WtLinguServiceTools: isSpellAuto: XPropertySet == null");
        return;
      }
      props.setPropertyValue("IsSpellAuto", spellAuto);
    } catch (Throwable t) {
      // If anything goes wrong, give the user a stack trace
      WtMessageHandler.printException(t);
    }
  }
*/
  public static boolean isWtGrammarServiceActive(XComponentContext xContext, Locale locale) {
    if (xContext != null) {
      XLinguServiceManager mxLinguSvcMgr = getLinguSvcMgr(xContext); 
      if (mxLinguSvcMgr == null) {
        WtMessageHandler.printToLogFile("LinguisticServices: setLtAsGrammarService: XLinguServiceManager == null");
        return false;
      }
      String[] configuredServices = mxLinguSvcMgr.getConfiguredServices("com.sun.star.linguistic2.Proofreader", locale);
      for (String service : configuredServices) {
        if (service.equals(WtOfficeTools.WT_SERVICE_NAME)) {
          return true;
        }
      }
    }
    return false;
  }
  
  public static boolean isWtSpellServiceActive(XComponentContext xContext, Locale locale) {
    if (xContext != null) {
      XLinguServiceManager mxLinguSvcMgr = getLinguSvcMgr(xContext); 
      if (mxLinguSvcMgr == null) {
        WtMessageHandler.printToLogFile("LinguisticServices: setLtAsGrammarService: XLinguServiceManager == null");
        return false;
      }
      String[] configuredServices = mxLinguSvcMgr.getConfiguredServices("com.sun.star.linguistic2.SpellChecker", locale);
      for (String service : configuredServices) {
        if (service.equals(WtSpellChecker.IMPLEMENTATION_NAME)) {
          return true;
        }
      }
    }
    return false;
  }
  
  /**
   * Set WT as grammar checker for a specific language
   * is normally used deactivate lightproof 
   */
  public static boolean setWtAsGrammarService(XComponentContext xContext, Locale locale) {
    if (xContext != null) {
      XLinguServiceManager mxLinguSvcMgr = getLinguSvcMgr(xContext); 
      if (mxLinguSvcMgr == null) {
        WtMessageHandler.printToLogFile("LinguisticServices: setLtAsGrammarService: XLinguServiceManager == null");
        return false;
      }
      Locale[] locales = WtDocumentsHandler.getLocales();
      for (Locale loc : locales) {
        if (WtOfficeTools.isEqualLocale(locale, loc)) {
          String[] serviceNames = mxLinguSvcMgr.getConfiguredServices("com.sun.star.linguistic2.Proofreader", locale);
/*
          if (serviceNames.length == 0) {
            WtMessageHandler.printToLogFile("LinguisticServices: setLtAsGrammarService: No configured Service for: " + WtOfficeTools.localeToString(locale));
          } else {
            for (String service : serviceNames) {
              WtMessageHandler.printToLogFile("Configured Linguistic Service: " + service + ", " + WtOfficeTools.localeToString(locale));
            }
          }
*/
          if (serviceNames.length != 1 || !serviceNames[0].equals(WtOfficeTools.WT_SERVICE_NAME)) {
/*
            String[] aServiceNames = mxLinguSvcMgr.getAvailableServices("com.sun.star.linguistic2.Proofreader", locale);
            for (String service : aServiceNames) {
              WtMessageHandler.printToLogFile("Available Linguistic Service: " + service + ", " + WtOfficeTools.localeToString(locale));
            }
*/
            String[] configuredServices = new String[1];
            configuredServices[0] = new String(WtOfficeTools.WT_SERVICE_NAME);
            mxLinguSvcMgr.setConfiguredServices("com.sun.star.linguistic2.Proofreader", locale, configuredServices);
            WtMessageHandler.printToLogFile("WT set as configured Service for Language: " + WtOfficeTools.localeToString(locale));
          }
          return true;
        }
      }
      WtMessageHandler.printToLogFile("WT doesn't support language: " + WtOfficeTools.localeToString(locale));
    }
    return false;
  }

  /**
   * Set LT as grammar checker for all supported languages
   * is normally used deactivate lightproof 
   */
  public static boolean setWtAsGrammarService(XComponentContext xContext) {
    if (xContext != null) {
      return setWtAsGrammarService(getLinguSvcMgr(xContext));
    } else {
      return false;
    }
  }

  /**
   * Set LT as grammar checker for all supported languages
   * is normally used deactivate lightproof 
   */
  private static boolean setWtAsGrammarService(XLinguServiceManager mxLinguSvcMgr) {
    int num = 0;
    try {
      if (mxLinguSvcMgr == null) {
        WtMessageHandler.printToLogFile("LinguisticServices: setLtAsGrammarService: XLinguServiceManager == null");
        return false;
      }
      Locale[] locales = WtDocumentsHandler.getLocales();
      for (Locale locale : locales) {
        String[] serviceNames = mxLinguSvcMgr.getConfiguredServices("com.sun.star.linguistic2.Proofreader", locale);
        if (serviceNames.length != 1 || !serviceNames[0].equals(WtOfficeTools.WT_SERVICE_NAME)) {
  /*
          String[] aServiceNames = mxLinguSvcMgr.getAvailableServices("com.sun.star.linguistic2.Proofreader", locale);
          for (String service : aServiceNames) {
            WtMessageHandler.printToLogFile("Available Linguistic Service: " + service + ", " + WtOfficeTools.localeToString(locale));
          }
  */
          String[] configuredServices = new String[1];
          configuredServices[0] = new String(WtOfficeTools.WT_SERVICE_NAME);
          mxLinguSvcMgr.setConfiguredServices("com.sun.star.linguistic2.Proofreader", locale, configuredServices);
          num++;
        }
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t); // If anything goes wrong, give the user a stack trace
    }
    WtMessageHandler.printToLogFile("WT set as configured Service for " + num + " Languages");
    return true;
  }

  /**
   * ActivateLT as spell checker for all supported languages
   * is normally used
   */
  public static boolean setWtAsSpellService(XComponentContext xContext) {
    return setWtAsSpellService(xContext, true);
  }
  /**
   * Activate / deactivate LT as spell checker for all supported languages
   * is normally used
   */
  public static boolean setWtAsSpellService(XComponentContext xContext, boolean activate) {
    int num = 0;
    try {
      if (xContext == null) {
        return false;
      }
      XLinguServiceManager mxLinguSvcMgr = getLinguSvcMgr(xContext);
      if (mxLinguSvcMgr == null) {
        WtMessageHandler.printToLogFile("LinguisticServices: setLtAsSpellService: XLinguServiceManager == null");
        return false;
      }
      Locale[] locales = WtDocumentsHandler.getLocales();
      WtMessageHandler.printToLogFile("LinguisticServices: setLtAsSpellService: Number locales: " + locales.length);
      for (Locale locale : locales) {
    /*
          String[] serviceNames = mxLinguSvcMgr.getAvailableServices("com.sun.star.linguistic2.SpellChecker", locale);
          WtMessageHandler.printToLogFile("Available Linguistic Service: NUmber: " + serviceNames.length + ", " + WtOfficeTools.localeToString(locale));
          for (String service : serviceNames) {
            WtMessageHandler.printToLogFile("Available Linguistic Service: " + service + ", " + WtOfficeTools.localeToString(locale));
          }
    */
        String[] serviceNames = mxLinguSvcMgr.getConfiguredServices("com.sun.star.linguistic2.SpellChecker", locale);
        List<String> serviceList = new ArrayList<>();
        if (!activate) {
          for (String serviceName : serviceNames) {
            if(!WtSpellChecker.IMPLEMENTATION_NAME.equals(serviceName)) {
              serviceList.add(serviceName);
            } else {
              num++;
            }
          }
        } else {
          boolean add = true;
          for (String serviceName : serviceNames) {
            if(WtSpellChecker.IMPLEMENTATION_NAME.equals(serviceName)) {
              add = false;
            }
            serviceList.add(serviceName);
          }
          if (add) {
            serviceList.add(WtSpellChecker.IMPLEMENTATION_NAME);
            num++;
          }
        }
        serviceNames = serviceList.toArray(new String[serviceList.size()]);
        mxLinguSvcMgr.setConfiguredServices("com.sun.star.linguistic2.SpellChecker", locale, serviceNames);
    /*      
          if (serviceNames.length != 1 || !serviceNames[0].equals(OfficeTools.WT_SERVICE_NAME)) {
            String[] aServiceNames = mxLinguSvcMgr.getAvailableServices("com.sun.star.linguistic2.Proofreader", locale);
            for (String service : aServiceNames) {
              MessageHandler.printToLogFile("Available Linguistic Service: " + service + ", " + OfficeTools.localeToString(locale));
            }
            String[] configuredServices = new String[1];
            configuredServices[0] = new String(OfficeTools.WT_SERVICE_NAME);
            mxLinguSvcMgr.setConfiguredServices("com.sun.star.linguistic2.Proofreader", locale, configuredServices);
            MessageHandler.printToLogFile("LT set as configured Service for Language: " + OfficeTools.localeToString(locale));
          }
    */
      }
    } catch (Throwable t) {
      WtMessageHandler.printException(t); // If anything goes wrong, give the user a stack trace
    }
    WtMessageHandler.printToLogFile("WT spell service (" + WtSpellChecker.IMPLEMENTATION_NAME + ")(" + num + " Locales) " 
              + (!activate ? "deactivated" : "activated"));
    return true;
  }
}

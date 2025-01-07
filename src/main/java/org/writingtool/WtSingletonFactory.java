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
package org.writingtool;

import com.sun.star.lang.XServiceInfo;
import com.sun.star.lang.XSingleComponentFactory;
import com.sun.star.uno.XComponentContext;

/**
 * This class is a factory that creates only a single instance,
 * or a singleton, of the Main class. Used for performance 
 * reasons and to allow various parts of code to interact.
 *
 * @author Fred Kruse
 */
public class WtSingletonFactory implements XSingleComponentFactory, XServiceInfo {

//  private transient org.languagetool.openoffice.Main instance;
  private static org.writingtool.WritingTool instance = null;
//  private static org.languagetool.openoffice.LanguageToolSpellChecker spellInstance = null;
  private boolean isSpellChecker = false;
  
  WtSingletonFactory(boolean isSpChecker) {
    isSpellChecker = isSpChecker;
  }

  @Override
  public final Object createInstanceWithArgumentsAndContext(Object[] arguments, 
      XComponentContext xContext) {
    return createInstanceWithContext(xContext);
  }

  @Override
  public final Object createInstanceWithContext(XComponentContext xContext) {
    if (isSpellChecker) {
      return new org.writingtool.WtSpellChecker(xContext);
//      if (spellInstance == null) {     
//        spellInstance = new org.languagetool.openoffice.LanguageToolSpellChecker(xContext);
//      }
//      return spellInstance;
    } else {
      if (instance == null) {     
        instance = new org.writingtool.WritingTool(xContext);      
      } else {  
        instance.changeContext(xContext);      
      }
      return instance;
    }
  }  

  @Override
  public final String getImplementationName() {
    return isSpellChecker ? WtSpellChecker.class.getName() : WritingTool.class.getName();
  }

  @Override
  public final boolean supportsService(String serviceName) {
    for (String s : getSupportedServiceNames()) {
      if (s.equals(serviceName)) {
        return true;
      }
    }
    return false;
  }

  @Override
  public final String[] getSupportedServiceNames() {
    return isSpellChecker ? WtSpellChecker.getServiceNames() : WritingTool.getServiceNames();
  }
}

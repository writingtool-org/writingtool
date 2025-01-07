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

import java.util.Enumeration;
import java.util.MissingResourceException;
import java.util.ResourceBundle;

/**
 * A resource bundle that uses the bundles in the following order
 * 1. WT resource bundle
 * 2. LT resource bundle
 * 3. LT fallback bundle
 * 
 * @since 1.0
 * @author Fred Kruse
 */
public class WtResourceBundle extends ResourceBundle {

  private final ResourceBundle wtBundle;
  private final ResourceBundle wtFallbackBundle;
  private final ResourceBundle bundle;
  private final ResourceBundle fallbackBundle;
  
  public WtResourceBundle(ResourceBundle wtBundle, ResourceBundle wtFallbackBundle, 
      ResourceBundle bundle, ResourceBundle fallbackBundle) {
    this.wtBundle = wtBundle;
    this.wtFallbackBundle = wtFallbackBundle;
    this.bundle = bundle;
    this.fallbackBundle = fallbackBundle;
  }

  @Override
  public Object handleGetObject(String key) {
    String s = null;
    if (wtBundle != null) {
      try {
        s = wtBundle.getString(key);
      } catch (MissingResourceException e) {
      }
    }
    if (s == null || s.trim().isEmpty()) {
      if (bundle != null) {
        try {
          s = bundle.getString(key);
        } catch (MissingResourceException e) {
        }
      }
      if (s == null || s.trim().isEmpty()) {
        try {
          s = wtFallbackBundle.getString(key);
        } catch (MissingResourceException e) {
          return fallbackBundle.getString(key);
        }
      }
      return s;
    }
    return s;
  }

  @Override
  public Enumeration<String> getKeys() {
    return bundle.getKeys();
  }

}

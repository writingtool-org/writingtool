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
package org.writingtool.languagedetectors;

/**
 * Helps to detect the language of strings by the Unicode range used by the characters.
 * @since 1.0
 */
public abstract class WtUnicodeLanguageDetector {

  private static final int MAX_CHECK_LENGTH = 100;

  protected abstract boolean isInAlphabet(int numericValue);

  public boolean isThisLanguage(String str) {
    int maxCheckLength = Math.min(str.length(), MAX_CHECK_LENGTH);
    for (int i = 0; i < maxCheckLength; i++) {
      int numericValue = str.charAt(i);
      if (isInAlphabet(numericValue)) {
        return true;
      }
    }
    return false;
  }

}

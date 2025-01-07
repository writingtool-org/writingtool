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

import org.junit.Test;
import org.writingtool.languagedetectors.WtTamilDetector;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;


public class TamilDetectorTest {

  @Test
  public void testIsThisLanguage() {
    WtTamilDetector detector = new WtTamilDetector();

    assertTrue(detector.isThisLanguage("இந்த"));
    assertTrue(detector.isThisLanguage("இ"));
    assertTrue(detector.isThisLanguage("\"லேங்குவேஜ்"));

    assertFalse(detector.isThisLanguage("Hallo"));
    assertFalse(detector.isThisLanguage("öäü"));

    assertFalse(detector.isThisLanguage(""));
    try {
      assertFalse(detector.isThisLanguage(null));
      fail();
    } catch (NullPointerException ignored) {}
  }

}

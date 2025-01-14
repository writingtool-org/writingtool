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

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

import java.util.Arrays;

import org.junit.Test;

public class ErrorSortTest {
  
  @Test
  public void testIsCorrectSortTest() {
    
    WtProofreadingError[] aError = new WtProofreadingError[11];
    aError[9] = createError("A1", 5, 3, 0);
    aError[8] = createError("B1", 17, 11, 0);
    aError[3] = createError("C1", 14, 5, 0);
    aError[10] = createError("C2", 14, 7, 0);
    aError[4] = createError("D1", 9, 10, 0);
    aError[2] = createError("E1", 30, 3, 0);
    aError[5] = createError("F1", 35, 4, 2);
    aError[6] = createError("F2", 35, 4, 2);
    aError[1] = createError("F3", 35, 4, 1);
    aError[0] = createError("F4", 35, 4, 0);
    aError[7] = createError("F5", 35, 4, 0);
    Arrays.sort(aError, new WtErrorPositionComparator());
    assertEquals("A1", aError[0].aRuleIdentifier);
    assertEquals("B1", aError[1].aRuleIdentifier);
    assertEquals("C1", aError[2].aRuleIdentifier);
    assertEquals("C2", aError[3].aRuleIdentifier);
    assertEquals("D1", aError[4].aRuleIdentifier);
    assertEquals("E1", aError[5].aRuleIdentifier);
    assertEquals("F1", aError[6].aRuleIdentifier);
    assertEquals("F2", aError[7].aRuleIdentifier);
    assertEquals("F3", aError[8].aRuleIdentifier);
    assertEquals("F4", aError[9].aRuleIdentifier);
    assertEquals("F5", aError[10].aRuleIdentifier);
    
  }

  private WtProofreadingError createError(String id, int start, int length, int numSuggestion) {
    WtProofreadingError error = new WtProofreadingError();
    error.aRuleIdentifier = id;
    error.nErrorStart = start;
    error.nErrorLength = length;
    error.aSuggestions = new String[numSuggestion];
    return error;
  }
  
}

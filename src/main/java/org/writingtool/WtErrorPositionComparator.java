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

import java.util.Comparator;

/**
 * A simple comparator for sorting errors by their position.
 */
public class WtErrorPositionComparator implements Comparator<WtProofreadingError> {

  @Override
  public int compare(WtProofreadingError match1, WtProofreadingError match2) {
    int error1pos = match1.nErrorStart;
    int error2pos = match2.nErrorStart;
    int error1len = match1.nErrorLength;
    int error2len = match2.nErrorLength;
    if (error1pos == error2pos) {
      if (error2len < error1len) {
        return 1;
      } else if (error1len < error1len) {
        return -1;
      }
    }
    if (error2pos > error1pos && error2pos < error1pos + error1len) {
      return 1;
    }
    if (error1pos > error2pos && error1pos < error2pos + error2len) {
      return -1;
    }
    if (error1pos > error2pos) {
      return 1;
    } else if (error1pos < error2pos) {
      return -1;
    } else {
      if (error1len > error2len) {
        return 1;
      }
      if (error1len < error2len) {
        return -1;
      }
      return Integer.compare(match2.aSuggestions.length, match1.aSuggestions.length);
/*
      if (match1.aSuggestions.length == 0 && match2.aSuggestions.length > 0) {
        return 1;
      }
      if (match2.aSuggestions.length == 0 && match1.aSuggestions.length > 0) {
        return -1;
      }
      if (match1.aSuggestions.length != 0 && match2.aSuggestions.length != 0
          && match1.aSuggestions.length != match2.aSuggestions.length) {
        return Integer.compare(match1.aSuggestions.length, match2.aSuggestions.length);
      }
*/      
    }
//    return match1.aRuleIdentifier.compareTo(match2.aRuleIdentifier);
//    return 0;
  }

}

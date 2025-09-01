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

import java.io.Serializable;

import com.sun.star.beans.PropertyValue;
import com.sun.star.linguistic2.SingleProofreadingError;

/**
 * Class of serializable proofreading errors
 * @author Fred Kruse
 * @since 1.0
 */
public class WtProofreadingError implements Serializable {

  private static final long serialVersionUID = 1L;
  public int nErrorStart;
  public int nErrorLength;
  public int nErrorType;
  public boolean bDefaultRule;
  public boolean bStyleRule;
  public boolean bPunctuationRule;
  public String aFullComment;
  public String aRuleIdentifier;
  public String aShortComment;
  public String[] aSuggestions;
  public WtPropertyValue[] aProperties = null;
  
  public WtProofreadingError() {
  }

  public WtProofreadingError(WtProofreadingError error) {
    nErrorStart = error.nErrorStart;
    nErrorLength = error.nErrorLength;
    nErrorType = error.nErrorType;
    aFullComment = error.aFullComment;
    aRuleIdentifier = error.aRuleIdentifier;
    aShortComment = error.aShortComment;
    aSuggestions = error.aSuggestions;
    bDefaultRule = error.bDefaultRule;
    bStyleRule = error.bStyleRule;
    bPunctuationRule = error.bPunctuationRule;
    if (error.aProperties != null) {
      aProperties = new WtPropertyValue[error.aProperties.length];
      for (int i = 0; i < error.aProperties.length; i++) {
        aProperties[i] = new WtPropertyValue(error.aProperties[i]);
      }
    }
  }
  
  public WtProofreadingError(SingleProofreadingError error) {
    nErrorStart = error.nErrorStart;
    nErrorLength = error.nErrorLength;
    nErrorType = error.nErrorType;
    aFullComment = error.aFullComment;
    aRuleIdentifier = error.aRuleIdentifier;
    aShortComment = error.aShortComment;
    aSuggestions = error.aSuggestions;
    if (error.aProperties != null) {
      aProperties = new WtPropertyValue[error.aProperties.length];
      for (int i = 0; i < error.aProperties.length; i++) {
        aProperties[i] = new WtPropertyValue(error.aProperties[i]);
      }
    }
  }
  
  public SingleProofreadingError toSingleProofreadingError () {
    SingleProofreadingError error = new SingleProofreadingError();
    error.nErrorStart = nErrorStart;
    error.nErrorLength = nErrorLength;
    error.nErrorType = nErrorType;
    error.aFullComment = aFullComment;
    error.aRuleIdentifier = aRuleIdentifier;
    error.aShortComment = aShortComment;
    error.aSuggestions = aSuggestions;
    if (aProperties != null) {
      error.aProperties = new PropertyValue[aProperties.length];
      for (int i = 0; i < aProperties.length; i++) {
        error.aProperties[i] = aProperties[i].toPropertyValue();
      }
    } else {
      error.aProperties = null;
    }
    return error;
  }

  /**
   * Define equal WtProofreadingError entries
   *//*
  @Override
  public boolean equals(Object o) {
    if (o == null || !(o instanceof WtProofreadingError)) {
      return false;
    }
    WtProofreadingError e = (WtProofreadingError) o;
    if (((aRuleIdentifier == null && e.aRuleIdentifier != null) || (aRuleIdentifier != null && e.aRuleIdentifier == null))
        || ((aFullComment == null && e.aFullComment != null) || (aFullComment != null && e.aFullComment == null))
        || ((aShortComment == null && e.aShortComment != null) || (aShortComment != null && e.aShortComment == null)) ) {
      return false;
    }
    if (nErrorStart == e.nErrorStart && nErrorLength == e.nErrorLength && nErrorType == e.nErrorType
        && bDefaultRule == e.bDefaultRule && bPunctuationRule == e.bPunctuationRule && bStyleRule == e.bStyleRule
        && ((aRuleIdentifier == null && e.aRuleIdentifier == null) || aRuleIdentifier.equals(e.aRuleIdentifier))
        && ((aShortComment == null && e.aShortComment == null) || aShortComment.equals(e.aShortComment)) 
        && ((aFullComment == null && e.aFullComment == null) || aFullComment.equals(e.aFullComment)) ) {
      return true;
    }
    return false;
  }
*/
  /**
   * Define equal WtProofreadingError entries (only ID and Position)
   */
  public boolean equalsIdAndPos(WtProofreadingError e) {
    if (e == null || e.aRuleIdentifier == null || aRuleIdentifier == null) {
      return false;
    }
    if (nErrorStart == e.nErrorStart && nErrorLength == e.nErrorLength && aRuleIdentifier.equals(e.aRuleIdentifier)) {
      return true;
    }
    return false;
  }

}

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

}

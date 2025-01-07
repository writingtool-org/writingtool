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

import com.sun.star.beans.PropertyState;
import com.sun.star.beans.PropertyValue;

/**
 * Class of serializable property values
 * @author Fred Kruse
 * @since 1.0
 */
public class WtPropertyValue implements Serializable {

  private static final long serialVersionUID = 1L;
  public String name;
  public Object value;
  
  public WtPropertyValue(WtPropertyValue properties) {
    name = properties.name;
    value = properties.value;
  }
  
  public WtPropertyValue(String name, Object value) {
    this.name = name;
    this.value = value;
  }
  
  public WtPropertyValue(PropertyValue properties) {
    name = properties.Name;
    value = properties.Value;
  }
  
  PropertyValue toPropertyValue() {
    PropertyValue properties = new PropertyValue();
    properties.Name = name;
    properties.Value = value;
    properties.Handle = -1;
    properties.State = PropertyState.DIRECT_VALUE;
    return properties;
  }
}

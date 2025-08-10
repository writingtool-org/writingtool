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

import java.util.ArrayList;
import java.util.List;

/**
 * Class to store the chapter as a node in hierarchy
 * 
 * @since 25.10
 * @author Fred Kruse
 */
public class WtChapter {
  public String name;
  public int from;
  public int to;
  public int hierarchy;
  public WtChapter parent;
  public List<WtChapter> children = null;
  public int weight;
  public String summary;
  
  public WtChapter(String name, int from, int to, int hierarchy, WtChapter parent) {
    this(name, from, to, hierarchy, parent, null, -1, null);
  }

  public WtChapter(String name, int from, int to, int hierarchy, WtChapter parent, int weight) {
    this(name, from, to, hierarchy, parent, null, weight, null);
  }

  public WtChapter(String name, int from, int to, int hierarchy, WtChapter parent, String summary) {
    this(name, from, to, hierarchy, parent, null, -1, summary);
  }

  public WtChapter(String name, int from, int to, int hierarchy, WtChapter parent, List<WtChapter> children, int weight, String summary) {
    this.name = name;
    this.from = from;
    this.to = to;
    this.hierarchy = hierarchy;
    this.weight = weight;
    this.parent = parent;
    if (children != null) {
      this.children = new ArrayList<>(children);
    }
  }
  
  public void setChildren(List<WtChapter> children) {
    this.children = new ArrayList<>(children);
  }
  
  public void setWeight(int weight) {
    this.weight = weight;
  }
  
  public void setSummary(String summary) {
    this.summary = summary;
  }
}


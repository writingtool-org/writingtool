/* WritingTool, a LibreOffice Extension based on LanguageTool 
 * Copyright (C) 2024 Fred Kruse (https://fk-es.de)
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
package org.writingtool.config;

import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.languagetool.Language;
import org.languagetool.rules.Rule;
import org.writingtool.tools.WtMessageHandler;
import org.writingtool.tools.WtOfficeTools;
import org.writingtool.tools.WtVersionInfo;
import org.writingtool.WtDocumentsHandler;
import org.writingtool.WtLanguageTool;
import org.writingtool.dialogs.WtConfigurationDialog;

/**
 * A thread that shows the configuration dialog which lets the
 * user enable/disable rules.
 * 
 * @author Fred Kruse
 */
public class WtConfigThread extends Thread {

  private final Language docLanguage;
  private final WtConfiguration config;
  private final WtLanguageTool lt;
  private final WtDocumentsHandler documents;
  private final WtConfigurationDialog cfgDialog;
  
  public WtConfigThread(Language docLanguage, WtConfiguration config, WtLanguageTool lt, WtDocumentsHandler documents) {
    if (config.getDefaultLanguage() == null) {
      this.docLanguage = docLanguage;
    } else {
      this.docLanguage = config.getDefaultLanguage();
    }
    this.config = config;
    this.lt = lt;
    this.documents = documents;
    String title = WtOfficeTools.getMessageBundle().getString("guiWtConfigWindowTitle") 
        + " (WT " + WtVersionInfo.getWtInformation()
        + ", LT " + WtVersionInfo.getLtInformation() + ")";
    cfgDialog = new WtConfigurationDialog(null, WtOfficeTools.getLtImage(), title, config);
  }

  @Override
  public void run() {
    if(!documents.javaVersionOkay()) {
      return;
    }
    if (!documents.isJavaLookAndFeelSet()) {
      WtDocumentsHandler.setJavaLookAndFeel();
    }
    documents.setConfigurationDialog(cfgDialog);
    try {
      List<Rule> allRules = lt.getAllRules();
      Set<String> disabledRulesUI = WtDocumentsHandler.getDisabledRules(docLanguage.getShortCodeWithCountryAndVariant());
      config.addDisabledRuleIds(disabledRulesUI);
      boolean configChanged = cfgDialog.show(allRules);
      if (configChanged) {
        Set<String> disabledRules = config.getDisabledRuleIds();
        Set<String> tmpDisabledRules = new HashSet<>(disabledRulesUI);
        for (String ruleId : tmpDisabledRules) {
          if(!disabledRules.contains(ruleId)) {
            disabledRulesUI.remove(ruleId);
          }
        }
        documents.setDisabledRules(docLanguage.getShortCodeWithCountryAndVariant(), disabledRulesUI);
        config.removeDisabledRuleIds(disabledRulesUI);
        config.saveConfiguration(docLanguage);
        documents.resetDocumentCaches();
        documents.resetConfiguration();
      } else {
        config.removeDisabledRuleIds(WtDocumentsHandler.getDisabledRules(docLanguage.getShortCodeWithCountryAndVariant()));
      }
    } catch (Throwable e) {
      WtMessageHandler.showError(e);
    }
    documents.setConfigurationDialog(null);
  }
  
}

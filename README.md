# WritingTool (a LibreOffice Extension based on LanguageTool) 



WritingTool is a LibreOffice extension for LibreOffice that adds a writing assistant to text editing. It is designed for creating and editing extensive texts (e.g. for literature, science or business).

In addition to a spelling, grammar and style checker based on LanguageTool, it provides statistical methods for optimizing texts as well as AI support (in development).

Compared to using the internal LanguageTool in LibreOffice, the extension offers the following advantages:

* The application runs on the local computer. Therefore, no remote server is required (the use of a remote server including the LanguageTool premium service is still supported).
* All rules that work at the full-text level are also supported across paragraph boundaries (e.g. opening and closing quotation marks are recognized even if they are more than one paragraph apart).
* In addition to that of LibreOffice, the LanguageTool spell checker is used.
* A custom caching mechanism is used, which significantly speeds up repeated work on long texts. It avoids rechecking unchanged sections, saving time. When a text that has already been checked is loaded, all previously found errors are displayed almost immediately.
* A custom WritingTool check dialog is used, designed to be highly performant. It has been optimized for WritingTool's internal caching system to ensure fast and efficient operation.
* Grammar checking of Impress and Calc documents is supported (Manual checking only through the LT check dialog. Automatic checking is not supported by LibreOffice.)
* Using a configuration dialog, users can change the following settings, for example:
  * Easily activate/deactivate optional rules
  * Easily deactivate/reactivate standard rules
  * Define custom colors and styles for rule groups or individual rules
  * Change parameters for some special rules
  * Define profiles for checking different document types
* The extension offers the possibility of statistical analyses, such as frequently used words, filler words, etc. (So far only for individual languages).
* Support for (local) AI:
  * Supplement to the grammar check
  * Generation, improvement, rewriting, expansion and translation of paragraphs
  * Creating illustrations for a paragraph
  * supporting OpenAI like API (Local applications such as https://localai.io are particularly supported)

The nightly snapshots contain the current LanguageTool snapshots, see here: https://writingtool.org/writingtool/snapshots/
The releases contain the current LanguageTool releases, see here: https://writingtool.org/writingtool/releases/

You can find more information (requirements, licence, downloads, etc.) here: https://writingtool.org



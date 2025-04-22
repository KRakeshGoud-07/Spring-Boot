Replace Placeholder with HTML in DOCX using Docx4j (Spring Boot)
This Spring Boot REST API demonstrates how to replace a placeholder in a DOCX file with HTML content (including styled tables) using the powerful Docx4j library.

 Features :
   Replaces a specific placeholder (data) in a .docx file
   Injects well-formatted HTML content (tables, headings, styles)
   Converts HTML (XHTML-compliant) to Word-compatible content using XHTMLImporterImpl
   Saves the modified DOCX to a new file
   Deletes the temporary XHTML file after processing

 How It Works
	The user places a DOCX file (1.docx) in their home directory with a placeholder text (e.g., data).
	The API /replace-placeholder-with-html reads this file.
	HTML content is defined inside the code (can be loaded from an external file or DB).
	The HTML is converted to Word-compatible elements using Docx4j’s XHTMLImporterImpl.
	The placeholder paragraph in the DOCX is located and replaced with the imported content.
	The result is saved to 2.docx in the home directory.

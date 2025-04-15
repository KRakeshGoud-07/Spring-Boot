package org.example.docx4j.test;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.wml.*;
import org.docx4j.XmlUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileOutputStream;

import java.nio.charset.StandardCharsets;
import java.net.URL;
import java.util.Iterator;
import java.util.List;

@RestController
public class A1 {

    @GetMapping("/replace-placeholder-with-html")
    public String replacePlaceholderWithHtml() {
        try {
            // Load the DOCX file
            String inputPath = System.getProperty("user.home") + "/1.docx";
            File inputFile = new File(inputPath);
            if (!inputFile.exists()) {
                return "❌ File not found at: " + inputPath;
            }

            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(inputFile);
            MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();

            // The HTML content you want to convert to DOCX (replace "$nbsp" if needed)
            String htmlContent = "<html>\n" +
                    "<head>\n" +
                    "<style>\n" +
                    "table {\n" +
                    "  font-family: arial, sans-serif;\n" +
                    "  border-collapse: collapse;\n" +
                    "  width: 100%;\n" +
                    "}\n" +
                    "\n" +
                    "td, th {\n" +
                    "  border: 1px solid #dddddd;\n" +
                    "  text-align: left;\n" +
                    "  padding: 8px;\n" +
                    "}\n" +
                    "\n" +
                    "tr:nth-child(even) {\n" +
                    "  background-color: #dddddd;\n" +
                    "}\n" +
                    "</style>\n" +
                    "</head>\n" +
                    "<body>\n" +
                    "\n" +
                    "<h2>HTML Table</h2>\n" +
                    "\n" +
                    "<table>\n" +
                    "  <tr>\n" +
                    "    <th>Company</th>\n" +
                    "    <th>Contact</th>\n" +
                    "    <th>Country</th>\n" +
                    "  </tr>\n" +
                    "  <tr>\n" +
                    "    <td>Alfreds Futterkiste</td>\n" +
                    "    <td>Maria Anders</td>\n" +
                    "    <td>Germany</td>\n" +
                    "  </tr>\n" +
                    "  <tr>\n" +
                    "    <td>Centro comercial Moctezuma</td>\n" +
                    "    <td>Francisco Chang</td>\n" +
                    "    <td>Mexico</td>\n" +
                    "  </tr>\n" +
                    "  <tr>\n" +
                    "    <td>Ernst Handel</td>\n" +
                    "    <td>Roland Mendel</td>\n" +
                    "    <td>Austria</td>\n" +
                    "  </tr>\n" +
                    "  <tr>\n" +
                    "    <td>Island Trading</td>\n" +
                    "    <td>Helen Bennett</td>\n" +
                    "    <td>UK</td>\n" +
                    "  </tr>\n" +
                    "  <tr>\n" +
                    "    <td>Laughing Bacchus Winecellars</td>\n" +
                    "    <td>Yoshi Tannamuri</td>\n" +
                    "    <td>Canada</td>\n" +
                    "  </tr>\n" +
                    "  <tr>\n" +
                    "    <td>Magazzini Alimentari Riuniti</td>\n" +
                    "    <td>Giovanni Rovelli</td>\n" +
                    "    <td>Italy</td>\n" +
                    "  </tr>\n" +
                    "</table> </body>\n" +
                    "</html>\n";  // Ensure it's XHTML format
            htmlContent = htmlContent.replace("&nbsp;", "&#160;");
            // Create a temporary file to hold the XHTML content
            File tempFile = new File(System.getProperty("user.home") + "/temp.xhtml");
            FileOutputStream outputStream = new FileOutputStream(tempFile);
            outputStream.write(htmlContent.getBytes(StandardCharsets.UTF_8));
            outputStream.close();

            // Use XHTMLImporterImpl to convert HTML (XHTML) to DOCX elements
            URL xhtmlFileUrl = tempFile.toURI().toURL();
            XHTMLImporterImpl importer = new XHTMLImporterImpl(wordMLPackage);
            List<Object> htmlDocElements = importer.convert(xhtmlFileUrl);

            // The placeholder to be replaced in the DOCX file
            String placeholder = "data";

            // Search for the placeholder and replace it with the HTML content
            List<Object> content = mdp.getContent();
            Iterator<Object> contentIterator = content.iterator();
            boolean replaced = false;

            while (contentIterator.hasNext()) {
                Object unwrapped = XmlUtils.unwrap(contentIterator.next());
                if (unwrapped instanceof P paragraph) {
                    StringBuilder paragraphText = new StringBuilder();

                    for (Object rObj : paragraph.getContent()) {
                        Object rUnwrapped = XmlUtils.unwrap(rObj);
                        if (rUnwrapped instanceof R run) {
                            for (Object tObj : run.getContent()) {
                                Object tUnwrapped = XmlUtils.unwrap(tObj);
                                if (tUnwrapped instanceof Text t) {
                                    paragraphText.append(t.getValue());
                                }
                            }
                        }
                    }

                    // If the paragraph contains the placeholder, replace it
                    if (paragraphText.toString().contains(placeholder)) {
                        contentIterator.remove();  // Remove the paragraph with the placeholder

                        // Add the converted HTML content to the document (this replaces the placeholder)
                        mdp.getContent().addAll(htmlDocElements);
                        replaced = true;
                        break;  // Only replace the first occurrence
                    }
                }
            }

            // If no placeholder was found or replaced
            if (!replaced) {
                return "❌ Placeholder not found in the document.";
            }

            // Save the modified DOCX
            String outputPath = System.getProperty("user.home") + "/2.docx";
            wordMLPackage.save(new File(outputPath));

            // Delete the temporary file after use
            tempFile.delete();

            return "✅ HTML content inserted and placeholder replaced. Saved to: " + outputPath;

        } catch (Exception e) {
            e.printStackTrace();
            return "❌ Error: " + e.getMessage();
        }
    }
}

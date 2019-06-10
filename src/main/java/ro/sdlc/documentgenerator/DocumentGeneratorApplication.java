package ro.sdlc.documentgenerator;

import org.apache.commons.io.FileUtils;
import org.docx4j.XmlUtils;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.exceptions.InvalidFormatException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.wml.RFonts;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import javax.xml.bind.JAXBException;
import java.io.File;
import java.io.IOException;

@SpringBootApplication
public class DocumentGeneratorApplication {

	public static void main(String[] args) {

	    String inputfilepath = "C:\\Work\\SDLC-Farming\\generated-docs\\welcome.html";
        String baseURL = "file:C:\\Work\\SDLC-Farming\\generated-docs";

		String stringFromFile = null;
		try {
			stringFromFile = FileUtils.readFileToString(new File(inputfilepath), "UTF-8");
		} catch (IOException e) {
			e.printStackTrace();
		}

		String unescaped = stringFromFile;

		System.out.println("Unescaped: " + unescaped);

        // Setup font mapping
        RFonts rfonts = Context.getWmlObjectFactory().createRFonts();
        rfonts.setAscii("Century Gothic");
        XHTMLImporterImpl.addFontMapping("Century Gothic", rfonts);

        // Create an empty docx package
        WordprocessingMLPackage wordMLPackage = null;
        try {
            wordMLPackage = WordprocessingMLPackage.createPackage();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }

        NumberingDefinitionsPart ndp = null;
        try {
            ndp = new NumberingDefinitionsPart();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        try {
            wordMLPackage.getMainDocumentPart().addTargetPart(ndp);
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        try {
            ndp.unmarshalDefaultNumbering();
        } catch (JAXBException e) {
            e.printStackTrace();
        }

        // Convert the XHTML, and add it into the empty docx we made
        XHTMLImporterImpl XHTMLImporter = new XHTMLImporterImpl(wordMLPackage);

        XHTMLImporter.setHyperlinkStyle("Hyperlink");
        try {
            wordMLPackage.getMainDocumentPart().getContent().addAll(
                    XHTMLImporter.convert(unescaped, baseURL) );
        } catch (Docx4JException e) {
            e.printStackTrace();
        }

        System.out.println(XmlUtils.marshaltoString(wordMLPackage.getMainDocumentPart().getJaxbElement(), true, true));

        try {
            wordMLPackage.save(new File("C:\\Work\\SDLC-Farming\\generated-docs\\welcome.docx"));
        } catch (Docx4JException e) {
            e.printStackTrace();
        }

        System.out.println("Finished.");
	}
}



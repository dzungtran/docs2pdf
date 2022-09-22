package com.yeokhengmeng.docstopdfconverter;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PrintStream;
import java.io.StringWriter;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.convert.in.xhtml.XHTMLImporterImpl;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.util.XMLHelper;

import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.w3c.dom.Document;
import org.w3c.tidy.Tidy;

public class DocToPDFConverter extends Converter {

	public DocToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages,
			boolean closeStreamsWhenComplete) {
		super(inStream, outStream, showMessages, closeStreamsWhenComplete);
	}

	@Override
	public void convert() throws Exception {

		loading();
		// Disable stdout temporarily as Doc convert produces alot of output
		PrintStream originalStdout = System.out;
		System.setOut(new PrintStream(new OutputStream() {
			public void write(int b) {
				// DO NOTHING
			}
		}));

		HWPFDocument document = new HWPFDocument(inStream);
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.createPackage();

		processing();
		String outHtml = getHtmlText(document);
		XHTMLImporterImpl XHTMLImporter = new XHTMLImporterImpl(wordMLPackage);
		wordMLPackage.getMainDocumentPart().getContent().addAll(XHTMLImporter.convert(outHtml, null));

		Docx4J.toPDF(wordMLPackage, outStream);
		System.setOut(originalStdout);

		finished();

	}

	private static String getHtmlText(HWPFDocument hwpfDocument) throws Exception {
		Document newDocument = XMLHelper.newDocumentBuilder().newDocument();
		WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(newDocument);

		wordToHtmlConverter.processDocument(hwpfDocument);

		StringWriter stringWriter = new StringWriter();
		Transformer transformer = XMLHelper.newTransformer();
		transformer.setOutputProperty(OutputKeys.METHOD, "html");
		transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
		transformer.transform(
				new DOMSource(wordToHtmlConverter.getDocument()),
				new StreamResult(stringWriter));

		Tidy tidy = new Tidy();
		tidy.setShowWarnings(false);
		tidy.setXmlTags(false);
		tidy.setInputEncoding("UTF-8");
		tidy.setOutputEncoding("UTF-8");
		tidy.setXHTML(true);//
		tidy.setMakeClean(true);

		ByteArrayInputStream inputStream = new ByteArrayInputStream(stringWriter.toString().getBytes("UTF-8"));
		ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
		tidy.parseDOM(inputStream, outputStream);

		return outputStream.toString();
	}

}

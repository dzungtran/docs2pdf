package com.yeokhengmeng.docstopdfconverter;

import java.io.InputStream;
import java.io.OutputStream;

import org.docx4j.Docx4J;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class DocxToPDFConverter extends Converter {

	public DocxToPDFConverter(InputStream inStream, OutputStream outStream, boolean showMessages,
			boolean closeStreamsWhenComplete) {
		super(inStream, outStream, showMessages, closeStreamsWhenComplete);
	}

	@Override
	public void convert() throws Exception {
		loading();
		
		WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(inStream);
		
		processing();
		
		Docx4J.toPDF(wordMLPackage, outStream);
		
		finished();
	}
}

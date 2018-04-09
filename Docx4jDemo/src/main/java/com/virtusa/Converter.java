package com.virtusa;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.commons.io.FilenameUtils;
import org.docx4j.Docx4J;
import org.docx4j.convert.in.Doc;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.OpcPackage;
import org.docx4j.openpackaging.packages.PresentationMLPackage;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

/**
 * Hello world!
 *
 */
public class Converter 
{
	public static final String DOC_EXTENSION = "doc";
	public static final String DOCX_EXTENSION = "docx";
	public static final String PPTX_EXTENSION = "pptx";
	public static void main( String[] args ) throws FileNotFoundException, Docx4JException
	{
		System.out.println("starting..");
		String filePathInput="doc/sample.doc";		
		String extension = FilenameUtils.getExtension(filePathInput);
		String baseName = FilenameUtils.getBaseName(filePathInput);
		File file = new File(filePathInput);
		String filePathOutput = "pdf/"+ baseName+".pdf";
		InputStream iStream = new FileInputStream(file);
		OutputStream outStream = new FileOutputStream(new File(filePathOutput));

		switch (extension.toLowerCase()) {
		case DOC_EXTENSION:
			docConverter(iStream,outStream);
			break;
		case DOCX_EXTENSION:
			docxConverter(iStream,outStream);
			break;
		default:
			break;
		}







	}


	public static void docConverter(InputStream iStream, OutputStream outStream){
		try {
			WordprocessingMLPackage wordMLPackage = Doc.convert(iStream);
			Docx4J.toPDF(wordMLPackage, outStream);
			System.out.println("successful");
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			System.out.println("failure");
			e.printStackTrace();
		} catch (Exception e) {
			e.printStackTrace();
		}		
	}


	public static void docxConverter(InputStream iStream, OutputStream outStream){
		try {
			WordprocessingMLPackage wordMLPackage = Docx4J.load(iStream);
			Docx4J.toPDF(wordMLPackage, outStream);
			System.out.println("successful");
		} catch (Docx4JException e) {
			// TODO Auto-generated catch block
			System.out.println("failure");
			e.printStackTrace();
		}		
	}


}

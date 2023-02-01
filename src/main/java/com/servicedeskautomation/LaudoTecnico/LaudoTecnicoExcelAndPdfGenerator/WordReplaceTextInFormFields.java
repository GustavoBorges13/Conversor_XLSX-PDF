package com.servicedeskautomation.LaudoTecnico.LaudoTecnicoExcelAndPdfGenerator;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.SimpleValue;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;

public class WordReplaceTextInFormFields {

	private static void replaceFormFieldText(XWPFDocument document, String ffname, String text) {
		boolean foundformfield = false;
		for (XWPFParagraph paragraph : document.getParagraphs()) {
			for (XWPFRun run : paragraph.getRuns()) {
				XmlCursor cursor = run.getCTR().newCursor();
				cursor.selectPath(
						"declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:fldChar/@w:fldCharType");
				while (cursor.hasNextSelection()) {
					cursor.toNextSelection();
					XmlObject obj = cursor.getObject();
					if ("begin".equals(((SimpleValue) obj).getStringValue())) {
						cursor.toParent();
						obj = cursor.getObject();
						obj = obj.selectPath(
								"declare namespace w='http://schemas.openxmlformats.org/wordprocessingml/2006/main' .//w:ffData/w:name/@w:val")[0];
						if (ffname.equals(((SimpleValue) obj).getStringValue())) {
							foundformfield = true;
						} else {
							foundformfield = false;
						}
					} else if ("end".equals(((SimpleValue) obj).getStringValue())) {
						if (foundformfield)
							return;
						foundformfield = false;
					}
				}
				if (foundformfield && run.getCTR().getTList().size() > 0) {
					run.getCTR().getTList().get(0).setStringValue(text);
					foundformfield = false;
					// System.out.println(run.getCTR());
				}
			}
		}
	}

	public static void main(String[] args) throws Exception {
		// Abrindo o arquivo...
		String fileName = "modelo laudo.docx";
		String userHome = System.getProperty("user.home");
		String pathRestante = "/Documents/ConversorXLSX-PDF/data/";
		File file = new File(userHome + pathRestante + fileName);
		FileInputStream fis;
		fis = new FileInputStream(file.getAbsolutePath());
		XWPFDocument document = new XWPFDocument(fis);

		// Pegando dados de outras classes
		String laudo = Principal.laudo.get(GerarLaudoPDF.linhasSelecionadas[0]);
		String analise = GerarLaudoPDF.editorPaneAnalise.getText();
		String consideracoes = GerarLaudoPDF.editorPaneConsideracoesTecnicas.getText();

		// Subistituindo as palavras do bookmarks...
		replaceFormFieldText(document, "Text1", laudo);
		replaceFormFieldText(document, "Text2", analise);
		replaceFormFieldText(document, "Text3", consideracoes);

		String[] ativo = Principal.ativo.get(GerarLaudoPDF.linhasSelecionadas[0]).split(" ");
		// Local onde será baixado
		File folder = new File(userHome + pathRestante + "backup");
		folder.mkdirs();
		try (FileOutputStream out = new FileOutputStream(
				folder.getPath() + "\\" + Principal.laudo.get(GerarLaudoPDF.linhasSelecionadas[0]) + " - " + ativo[0]
						+ " - " + Principal.nomeSolicitante.get(GerarLaudoPDF.linhasSelecionadas[0]) + ".docx")) {
			document.write(out);
			out.close();
			document.close();
		}
	}
}